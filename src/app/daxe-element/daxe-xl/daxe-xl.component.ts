import { Component, Inject, LOCALE_ID, OnDestroy, OnInit } from '@angular/core';
import { Subscription } from 'rxjs';
import { DaxeXLBaseEvents, daxeXlEvent, DaxeXlEventType, Field, LayoutDetail, MdoRecord, RequestType } from 'daxe-ui-sdk';
import { codeFields, data as generatedElementData, dynamicClassificationSheet, fieldData } from './daxe-header-data';
import { CommonModule } from '@angular/common';
import { ExcelReaderService } from '../../services/excel-reader.service';
import { UtilsService } from '../../services/utils.service';
import { ExcelDataWriterService } from '../../services/excel-data-writer.service';
import { SharedService } from '../../services/shared.service';
import { ExcelEventHandlerService } from '../../services/excel-event-handler.service';
import { DaxeXLService } from '../../services/daxe-xl.service';
import { TranslationService } from '../../services/translation.service';
import { TranslatePipe } from '../../pipes/translate.pipe';
import { SheetContext, SheetContextStore } from '../../services/sheet-context-store.service';
import { mocrRootContexts } from './mocr-root-contexts';
import { multiSheetContexts } from './multi-sheet-contexts';

type XfuseDaxeXlAction = 'generateDescription' | 'applyDataRules' | 'validateRecords' | 'submitRecords' | 'sendForCorrection' | 'importRecordsFromExcel';

const XFUSE_DAXE_XL_ACTION_REQUEST = 'xFuseDaxeXlActionRequest';
// Host-side GlobalVariables are intentionally not bundled into a generated
// DAXE element. This request lets ngx-mdo decide whether the current form has
// description generation enabled whenever the user changes worksheet.
const DAXE_XL_DESCRIPTION_GENERATOR_STATE_REQUEST = 'daxeXLDescriptionGeneratorStateRequest';

@Component({
  selector: 'daxe-xl',
  imports: [CommonModule,TranslatePipe],
  templateUrl: './daxe-xl.component.html',
  styleUrls: ['./daxe-xl.component.scss'],
  standalone: true,
})
export class DAXEXLComponent implements OnInit, OnDestroy {
  //Subscription object
  subscription: Subscription | undefined;
  isSendForCorrectionVisible: boolean = false;
  private mocrRootIndex = -1;
  private multiSheetIndex = -1;
  private readonly isMocrWorkbook = mocrRootContexts.length > 0;
  private readonly isMultiSheetWorkbook = multiSheetContexts.length > 0;
  private readonly worksheetActivatedHandler = (event: any) => {
    void this.publishDescriptionGeneratorState(event?.worksheetId);
  };
  private worksheetActivationRegistered = false;
  /** Configuration is generated per element; never infer it from another element's flow. */
  private readonly generationProfile = (generatedElementData as any)?.generationProfile || {
    process: 'change',
    autoImportBeforeValidation: true,
    changeImportModes: [],
    acceptsHostImportResponse: true,
    descriptionGeneration: false,
    classificationEnabled: false,
    staticClassificationAttributes: false,
    dynamicClassificationAttributes: false
  };
  private importResponseResolver?: (success: boolean) => void;

  constructor(protected readonly excelReader: ExcelReaderService,
    protected readonly excelWriter: ExcelDataWriterService,
    protected readonly utilsService: UtilsService,
    protected readonly sharedService: SharedService,
    private readonly eventHandler: ExcelEventHandlerService,
    protected readonly daxeXlService: DaxeXLService,
    @Inject(LOCALE_ID) public locale: string, protected readonly translationService: TranslationService,
    private readonly sheetContextStore: SheetContextStore) {

  }
  ngOnInit(): void {
    // Header data is copied for each generated element. Apply only its profile
    // to this element-local SharedService instance.
    this.sharedService.data.classificationEnabled = this.generationProfile.classificationEnabled === true;
    (this.sharedService.data as any).generationProfile = this.generationProfile;
    void this.registerWorksheetActivationHandler();
    console.log("current application language", this.locale)
    if (daxeXlEvent && !this.subscription) {
      console.log("subscribing to events")
      this.subscription = daxeXlEvent.subscribe(async (event: DaxeXlEventType) => {
        try {
          const actionEvent = event as any;
          if (actionEvent.type === XFUSE_DAXE_XL_ACTION_REQUEST) {
            await this.handleXfuseAction(actionEvent.payload?.action);
            return;
          }
          switch (event.type) {
            case DaxeXLBaseEvents.Ready:
              daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSelectedFlowDetailsRequest, payload: {} as any });
              break;
            case DaxeXLBaseEvents.XFUSELLayoutDetailsResponse:
              console.log('Received XFUSELLayoutDetailsResponse event ', event);
              this.sharedService.data.layoutDetails = event.data;
              if (this.isMocrWorkbook && this.mocrRootIndex >= 0) {
                const root = mocrRootContexts[this.mocrRootIndex];
                this.createFieldMap(event.data, root.fieldData, root.codeFields);
                try {
                  await this.utilsService.toggleRuntimeEvents(false);
                  await this.excelWriter.createHeader(true, true, root.sheetName);
                  // Each MOCR root is protected, but never activated merely to
                  // apply protection. This leaves the user's visible sheet intact.
                  await this.utilsService.protectWorksheet(root.sheetName);
                } finally {
                  await this.utilsService.toggleRuntimeEvents(true);
                }
                this.sheetContextStore.snapshotContext(root.sheetName, this.sharedService);
                await this.requestNextMocrRootLayout();
                break;
              }
              if (this.isMultiSheetWorkbook && this.multiSheetIndex >= 0) {
                const sheetContext = multiSheetContexts[this.multiSheetIndex];
                this.createFieldMap(event.data, sheetContext.fieldData, sheetContext.codeFields);
                try {
                  await this.utilsService.toggleRuntimeEvents(false);
                  await this.excelWriter.createHeader(true, true, sheetContext.sheetName);
                  // Write directly to the named worksheet; never activate each
                  // sheet while initialising a multi-sheet workbook.
                  await this.utilsService.protectWorksheet(sheetContext.sheetName);
                } finally {
                  await this.utilsService.toggleRuntimeEvents(true);
                }
                this.sheetContextStore.snapshotContext(sheetContext.sheetName, this.sharedService);
                await this.requestNextMultiSheetLayout();
                break;
              }
              this.sharedService.data.uniqueStructures = this.utilsService.getUniqueStructureList(event.data).map(m => m.structure);
              this.createFieldMap(event.data);
              try {
                await this.utilsService.toggleEvents(false);
                await this.excelWriter.createHeader();
                await this.createDynamicClassificationSheet();
              } finally {
                await this.utilsService.toggleEvents(true);
              }
              break;
            case DaxeXLBaseEvents.XFUSESelectedFlowDetailsResponse:
              console.log('Received XFUSESelectedFlowDetailsResponse event ', event);
              const flowData: FlowModel = event.data;
              if (flowData.isMocrEnabled === true && this.isMocrWorkbook) {
                await this.startMocrWorkbook(flowData);
                break;
              }
              if (flowData.isMocrEnabled !== true && this.isMultiSheetWorkbook) {
                await this.startMultiSheetWorkbook(flowData);
                break;
              }
              this.sharedService.isMocrRequest = false;
              this.sharedService.data.flowId = flowData.flowId;
              this.sharedService.data.stepId = flowData.stepId;
              this.sharedService.data.controlData.moduleId = Object.keys(flowData.parentMapping)?.[0];
              this.sharedService.data.controlData.layoutId = this.sharedService.data.formId = flowData.parentMapping[this.sharedService.data.controlData.moduleId].formId;
              this.sharedService.data.formDesc = flowData.parentMapping[this.sharedService.data.controlData.moduleId].formDesc;
              const sanitizedName = this.sharedService.data.formDesc.replace(/[\\\/\?\*\[\]:]/g, '');
              const newSheetName = sanitizedName.substring(0, 31);
              this.sharedService.data.baseStructure = [{
                dataSetId: flowData.parentMapping ? Object.keys(flowData.parentMapping)[0] : '',
                formId: flowData.parentMapping ? flowData.parentMapping[Object.keys(flowData.parentMapping)[0]]?.formId : '',
                StructureList: [{
                  structure: '1',
                  sheetName: newSheetName
                },]
              },
              {
                dataSetId: flowData.parentMapping ? Object.keys(flowData.parentMapping)[0] : '',
                formId: flowData.childMapping?.[0]?.[Object.keys(flowData.parentMapping || {})[0]]?.formId ?? '',
                StructureList: [{
                  structure: '1',
                  sheetName: `${newSheetName}_Child`
                },]
              }
            ];
              const childMapping: FormMapping | undefined = flowData?.childMapping?.find(m => Object.keys(m).includes(this.sharedService.data.controlData.moduleId))?.[this.sharedService.data.controlData.moduleId];
              this.sharedService.data.childFormId = childMapping ? childMapping.formId : '';
              daxeXlEvent?.next({ type: DaxeXLBaseEvents.DAXEXLLayoutDetailsRequest, payload: { layoutId: this.sharedService.data.formId, moduleId: this.sharedService.data.controlData.moduleId } });
              daxeXlEvent?.next({ type: DaxeXLBaseEvents.DAXEXLBusinessRulesRequest, payload: { flowId: this.sharedService.data.flowId, stepId: this.sharedService.data.stepId } as any });
              daxeXlEvent?.next({ type: DaxeXLBaseEvents.DAXEXLControlDataRequest });
              daxeXlEvent?.next({ type: DaxeXLBaseEvents.DAXEXLGlobalConfigRequest });
              this.sharedService.initPersistence();
              this.validateSendForCorrection();
              break;
            case DaxeXLBaseEvents.XFUSEGlobalConfigResponse: {
              console.log('Received XFUSEGlobalConfigResponse Data ', event);
              this.sharedService.globalConfig = event.data;
              this.utilsService.globalConfig = event.data;
              this.translationService.use(event.data?.UserProfile?.selfServiceUserModel?.lang || 'en');
              break;
            }
            case DaxeXLBaseEvents.XFUSEControlDataResponse:
              console.log('Received XFUSEControlDataResponse Data ', event);
              // Hosts in use return control data either as flat fields or nested
              // under controlData. Resolve both shapes once so every generated
              // MOCR root receives the same mass, tenant, user and MOCR context.
              const controlDataResponse = event.data as typeof event.data & {
                mocrId?: string;
                controlData?: {
                  massId?: string;
                  userId?: string;
                  tenantId?: string;
                  mocrId?: string;
                };
              };
              const nestedControlData = controlDataResponse.controlData;
              const massId = controlDataResponse.massId || nestedControlData?.massId || '';
              const userId = controlDataResponse.userId || nestedControlData?.userId || '';
              const tenantId = controlDataResponse.tenantId || nestedControlData?.tenantId || '';
              const mocrId = controlDataResponse.mocrId || nestedControlData?.mocrId || '';
              this.sharedService.data.controlData.massId = massId;
              this.sharedService.data.controlData.userId = userId;
              this.sharedService.data.controlData.tenantId = tenantId;
              (this.sharedService.data.controlData as any).mocrId = mocrId;
              if (this.isMocrWorkbook || this.isMultiSheetWorkbook) {
                this.sheetContextStore.updateControlData({
                  massId,
                  userId,
                  tenantId,
                  mocrId
                });
              }
              break;
            case DaxeXLBaseEvents.XFUSEDropDownSelectionsResponse:
              try {
                await this.utilsService.toggleEvents(false);
                await this.excelWriter.setSelectedDropDownValues(event.data);
                if ((event.data as any)?.field?.isDynamicClassificationCharacteristic === true) {
                  await this.rememberDynamicCharacteristicValue((event.data as any)?.value?.itemObject);
                }
              } finally {
                await this.utilsService.toggleEvents(true);
              }
              break;
            case DaxeXLBaseEvents.XFUSEBusinessRulesResponse:
              console.log('Received XFUSEBusinessRulesResponse Rules ', event);
              this.sharedService.data.businessRules = event.data;
              break;
            case DaxeXLBaseEvents.XFUSEClassAttributesResponse:
              console.log("Class Attributes Response received:", event);
              //this.sharedService.attributesMap = event.data.response;
              // Only the Description Class owns attribute columns in this
              // sheet. Dynamic CLASS_TYPE attributes belong to their separate
              // generated dynamic-class sheet and must never be appended here.
              // Older hosts omit fieldId; their response is safe only when
              // this project has no dynamic-class sheet.
              const classResponse = event.data as any;
              // The picker response carries the real class type/UUID. Cache it
              // before processing the sheet because a template Description
              // Class column intentionally has no fixed CLASS_TYPE metadata.
              this.sharedService.cacheClassAttributes(classResponse?.response);
              const dynamicConfig = dynamicClassificationSheet as any;
              if (dynamicConfig?.classFieldId && classResponse?.fieldId === dynamicConfig.classFieldId) {
                await this.rememberDynamicClassAttributes(
                  event.data.response?.charList?.filter((attribute: any) => String(attribute?.status).toLowerCase() === 'true') || []
                );
                break;
              }
              const isDescriptionClassResponse = classResponse?.fieldId === 'CLASSIFICATION_CLASS'
                || (!classResponse?.fieldId && this.generationProfile.dynamicClassificationAttributes !== true);
              const responseRowIndex = await this.resolveDescriptionClassRowIndex(classResponse?.rowIndex);
              if (this.sharedService.classificationEnabled === true && isDescriptionClassResponse && responseRowIndex !== null) {
                try {
                  await this.utilsService.displayStatus(this.translationService.translate('Loading class attributes...'));
                  await this.utilsService.toggleEvents(false);
                  await Excel.run({ delayForCellEdit: true }, async (context) => {
                    await this.excelWriter.createAttributesInSheetBatch([{
                      rowIndex: responseRowIndex,
                      attrs: event.data.response.charList?.filter(m => m.status?.toString()?.toLowerCase() === 'true') || []
                    }], context);
                    //await this.daxeXlService.generateDescription(responseRowIndex, responseRowIndex, context, false);
                  });
                } finally {
                  await this.utilsService.toggleEvents(true);
                  await this.utilsService.displayStatus(null);
                }
              }
              break;

            case DaxeXLBaseEvents.XFUSESubmitRecordResponse:
              console.log("Submit Response received:", event.data.records);
              this.sharedService.isProcessRunning = false;

              // Update Status & Message columns based on each record's controlData
              if (event.data?.records?.length && this.isMocrWorkbook) {
                await this.updateMocrRootStatuses(event.data.records);
                this.openValidationErrorSideSheet(event.data.records);
                break;
              }
              if (event.data?.records?.length && this.isMultiSheetWorkbook) {
                await this.updateMultiSheetStatuses(event.data.records);
                this.openValidationErrorSideSheet(event.data.records);
                break;
              }
              if (event.data?.records?.length) {
                await this.utilsService.toggleEvents(false);
                await this.utilsService.displayStatus(this.translationService.translate('Updating status and messages...'));
                await this.excelWriter.updateStatusAndMessages(event.data.records).catch(err => {
                  console.error('Failed updating status/message cells:', err);
                });
                this.openValidationErrorSideSheet(event.data.records);
                await this.utilsService.displayStatus(null);
                await this.utilsService.toggleEvents(true);
              }
              // Here we can write a code to use event.data to create your own view
              // when data send by MDO specific to Daxebackend API called triggered
              break;
            case DaxeXLBaseEvents.XFUSEGetSendForCorrectionResponse:
              //excelresp
              await this.utilsService.toggleEvents(false);
              console.log("sent for correction resp", event.data.excelresp);
              try {
                await this.excelWriter.updateStatusAndMessages(event.data.excelresp, RequestType.SentForCorrection).catch(err => {
                  console.error('Failed updating status/message cells:', err);

                });
              }
              catch (err) {
                console.log(err);
              }
              finally {
                await this.utilsService.toggleEvents(true);
              }

              // Here we can write a code to use event.data to create your own view
              // when data send by MDO specific to Daxebackend API called triggered
              break;
            case DaxeXLBaseEvents.XFUSEApplyDataRuleResponse:
              console.log("Apply Data Rules Response received:", event.data);
              this.sharedService.isProcessRunning = false;
              await this.utilsService.toggleEvents(false);
              await this.utilsService.displayStatus(this.translationService.translate('Applying datarules...'));
              if (this.isMocrWorkbook) {
                await this.updateMocrRootDataRuleResponse(event.data.result || []);
              } else if (this.isMultiSheetWorkbook) {
                await this.updateMultiSheetDataRuleResponse(event.data.result || []);
              } else {
                await this.excelWriter.parseAndwriteApplyDataRuleResponse(event.data.result);
              }
              await this.utilsService.displayStatus(null);
              await this.utilsService.toggleEvents(true);
              // Here we can write a code to use event.data to create your own view
              // when data send by MDO specific to Daxebackend API called triggered
              break;
            case DaxeXLBaseEvents.XFUSEModuleReferenceRecordResponse:
              console.log("Module Reference Record Response received:", event.data);
              if (!event.data?.rowIndex) {
                this.sharedService.lock(async () => {
                  await this.utilsService.displayStatus(this.translationService.translate('Updating Sap number'));
                  await this.excelWriter.updateModuleReferenceRecordData(event.data);
                  await this.utilsService.displayStatus(null);
                });
              }
              break;

            case DaxeXLBaseEvents.XFUSEImportDataResponse:
              await this.utilsService.toggleEvents(false);
              try {
                await this.utilsService.displayStatus(this.translationService.translate('Loading data...'));
                console.log("Import data response received::::", event.data);
                this.sharedService.recordsImported = true;
                const fieldData = [...this.sharedService.fieldsMap.values()];
                const uniqueCodes = this.sharedService.collectAllUniqueClassCodes(event.data);
                console.log("CLASS unique codes::::::;", uniqueCodes)
                this.sharedService.saveRecordsImported(event.data);
                //await this.sharedService.updateAttributeMap(uniqueCodes);
                const isMDOConfigurationExist = await this.utilsService.isMDOConfigurationExist()
               
                  await this.excelWriter.writeFlattenedDataToSheet(event.data, "", fieldData, 2);
                
                // Resolve only after the imported records have been written;
                // Change validation reads those cells immediately afterwards.
                this.importResponseResolver?.(true);
                this.importResponseResolver = undefined;
              } catch (err) {
                console.error("Error writing imported data to sheet:", err);
                this.importResponseResolver?.(false);
                this.importResponseResolver = undefined;
              }
              finally {
                await this.utilsService.displayStatus(null);
              }
              await this.utilsService.toggleEvents(true);
              break;
            
            case DaxeXLBaseEvents.Die:
              // Calling ngDestroy to make sure all subscription and cleaning logic should execute
              this.ngOnDestroy();
              // This is event get triggered show we can unsubcribe/unload our component.
              break;
            case DaxeXLBaseEvents.XFUSEErrorNavigationResponse:
              await this.utilsService.toggleEvents(false);
              try {
                await Excel.run({ delayForCellEdit: true }, async (context) => {
                  const sheet = context.workbook.worksheets.getActiveWorksheet();
                  const fieldId = event.data?.fieldId;

                  if (!fieldId) return;

                  const fieldMeta = this.sharedService.staticFieldsObject.find(
                    (f: any) => f.fieldId === fieldId
                  );

                  if (!fieldMeta) return;

                  const columnIndex = fieldMeta.columnNumber;
                  const usedRange = sheet.getUsedRange();
                  usedRange.load(['rowCount', 'rowIndex']);
                  await context.sync();

                  const rowIndex = parseInt(event.data?.arbitraryRecords[0]) + 1;
                  const targetCell = sheet.getCell(rowIndex, columnIndex);
                  targetCell.load(['address']);
                  await context.sync();
                  targetCell.select();
                  await context.sync();
                });
              } catch (err) {
                console.error('Failed to navigate to error cell:', err);
              } finally {
                await this.utilsService.toggleEvents(true);
              }
            break;  
          }
        }
        catch (error) {
          console.error("Error handling event:", error);
        }
      });
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSelectedFlowDetailsRequest, payload: {} as any });
    }
    else {
      this.utilsService.toggleEvents(true);
      console.log("event not binded correctly with AE")
    }
    this.eventHandler.registerExcelEvents();
  }

  createFieldMap(tabs: LayoutDetail[], configuredFields: any[] = fieldData, configuredCodeFields: string[] = codeFields) {
    this.sharedService.codeFields = [...configuredCodeFields];
    this.sharedService.extentionHeaderFields = configuredFields.map(field => field.fieldId);
    this.sharedService.staticFieldsObject = this.getTranslatedFields(configuredFields);
    const fieldsMap = new Map<string, any>();

    const addGridField = (parentField: any, gridMeta: any) => {
      const gridFieldId: string | undefined = gridMeta?.fieldId ?? gridMeta?.fId;
      if (!gridFieldId) return;
      if (fieldsMap.has(gridFieldId)) return;

      const synthesized = {
        fieldId: gridFieldId,
        fieldType: parentField?.fieldType ?? 'Grid',
        metadata: [gridMeta],
        layoutId: parentField?.layoutId,
        ...this.sharedService.staticFieldsObject.find(f => f.fieldId === gridFieldId)
      };
      fieldsMap.set(gridFieldId, synthesized);
    };

    for (const tab of tabs ?? []) {
      const tabFields = tab?.fields ?? [];
      for (const field of tabFields) {
        if (field?.fieldId) {
          const configuredField = this.sharedService.staticFieldsObject.find(f => f.fieldId === field.fieldId);
          // CLASS_TYPE controls are returned by some layouts with `metadata:
          // null`.  Always expose an array to the Excel event handler and
          // retain the generated class type metadata that ngx-mdo needs for
          // the Class dropdown.
          const metadata = Array.isArray(field.metadata) && field.metadata.length > 0
            ? field.metadata
            : (Array.isArray(configuredField?.metadata) && configuredField.metadata.length > 0
              ? configuredField.metadata.map((meta: any) => ({
                ...meta,
                moduleId: meta.moduleId || field.moduleId || generatedElementData.controlData.moduleId
              }))
              : ((field.fieldType === 'CLASS_TYPE' || configuredField?.isClassificationField === true)
                ? [{
                  fieldId: field.fieldId,
                  description: configuredField?.label || field.description || 'Class',
                  pickList: '1',
                  structureId: 1,
                  isGridColumn: false,
                  moduleId: field.moduleId || generatedElementData.controlData.moduleId,
                  classCtrl: field.classCtrl ?? configuredField?.classCtrl,
                  classTypeCtrl: field.classTypeCtrl ?? configuredField?.classTypeCtrl,
                  classType: field.classType ?? configuredField?.classType
                }]
                : []));
          fieldsMap.set(field.fieldId, {
            ...field,
            ...configuredField,
            metadata
          });
        }

        const metas = field?.metadata ?? [];
        for (const meta of metas) {
          const gridArray = meta?.grid ?? [];
          if (!Array.isArray(gridArray)) continue;
          for (const gridMeta of gridArray) {
            addGridField(field, gridMeta);
          }
        }
      }
    }

    // Static-class characteristics are expanded from classification metadata
    // when the project is generated. They are not returned as normal layout
    // fields, so materialise their ngx-compatible metadata here. Without this
    // step a visible template characteristic would be ignored by the reader.
    for (const configuredField of this.sharedService.staticFieldsObject) {
      // A template Descriptions/Class column is a legacy payload helper, not
      // a form layout field. Materialise it so ngx-mdo opens its normal
      // description-class picker instead of treating it as a CLASS_TYPE.
      if (configuredField?.fieldId === 'CLASSIFICATION_CLASS' && !fieldsMap.has(configuredField.fieldId)) {
        fieldsMap.set(configuredField.fieldId, {
          ...configuredField,
          fieldType: 'Static',
          metadata: Array.isArray(configuredField.metadata) && configuredField.metadata.length > 0
            ? configuredField.metadata.map((meta: any) => ({
              ...meta,
              // ngx-mdo's legacy description-class lookup calls
              // metadata[0].moduleId.toString(). Header metadata generated
              // from an Excel template therefore needs the active module.
              moduleId: meta.moduleId || generatedElementData.controlData.moduleId
            }))
            : [{
              fieldId: configuredField.fieldId,
              description: configuredField.label || 'Class',
              pickList: '1',
              structureId: 1,
              isGridColumn: false,
              moduleId: generatedElementData.controlData.moduleId
            }]
        });
        continue;
      }
      if (!configuredField?.isStaticClassCharacteristics || fieldsMap.has(configuredField.fieldId)) continue;
      const attributeMetadata = configuredField.attributeMetadata || {};
      fieldsMap.set(configuredField.fieldId, {
        ...configuredField,
        fieldType: configuredField.fieldType || 'Static',
        metadata: [{
          fieldId: configuredField.fieldId,
          description: configuredField.label || attributeMetadata.charDesc || configuredField.fieldId,
          dataType: configuredField.dataType || attributeMetadata.dataType,
          pickList: configuredField.pickList || (attributeMetadata.fieldType === 'DROPDOWN' ? '1' : '0'),
          structureId: configuredField.structureId || 1,
          isGridColumn: false,
          moduleId: generatedElementData.controlData.moduleId,
          structDesc: configuredField.group || 'Static Class Attributes'
        }],
        attributeMetadata
      });
    }

    this.sharedService.fieldsMap = fieldsMap;
    this.sharedService.isExtentionProcess = true;
    console.log('Fields Map created:', this.sharedService.fieldsMap);
  }

  /**
   * New hosts echo the DAXE-owned row index. The active-selection fallback is
   * retained only for elements generated before that host contract existed.
   */
  private async resolveDescriptionClassRowIndex(rowIndex: unknown): Promise<number | null> {
    if (typeof rowIndex === 'number' && Number.isInteger(rowIndex) && rowIndex >= 2) {
      return rowIndex;
    }
    return Excel.run({ delayForCellEdit: true }, async (context) => {
      const selectedRange = context.workbook.getSelectedRange();
      selectedRange.load('rowIndex');
      await context.sync();
      return selectedRange.rowIndex >= 2 ? selectedRange.rowIndex : null;
    });
  }

  /**
   * Dynamic classes are row-based classification data, not columns in the
   * parent form. Create the sheet without replacing the primary worksheet and
   * retain the real CLASS_TYPE metadata for the host dropdown.
   */
  private async createDynamicClassificationSheet(): Promise<void> {
    const config = dynamicClassificationSheet as any;
    if (!config?.sheetName || !Array.isArray(config.fields) || !config.fields.length) return;

    const primaryFields = this.sharedService.staticFieldsObject;
    const primaryCodeFields = this.sharedService.codeFields;
    this.sharedService.dynamicClassificationSheet = config;
    for (const field of config.fields) {
      this.sharedService.fieldsMap.set(field.fieldId, {
        ...field,
        layoutId: generatedElementData.controlData.layoutId,
        metadata: field.metadata || []
      });
    }
    try {
      this.sharedService.staticFieldsObject = this.getTranslatedFields(config.fields);
      this.sharedService.codeFields = [];
      await this.excelWriter.createHeader(true, true, config.sheetName);
      await this.utilsService.protectWorksheet(config.sheetName);
    } finally {
      // The main reader and validation flow always operate on the parent form.
      // Dynamic-sheet rows are joined separately by Record Number.
      this.sharedService.staticFieldsObject = primaryFields;
      this.sharedService.codeFields = primaryCodeFields;
    }
  }

  /** Stores the class-specific characteristic list for the selected dynamic-sheet row. */
  private async rememberDynamicClassAttributes(attributes: any[]): Promise<void> {
    const config = dynamicClassificationSheet as any;
    if (!config?.sheetName) return;
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = context.workbook.getActiveCell();
      sheet.load(['id', 'name']);
      cell.load('rowIndex');
      await context.sync();
      if (sheet.name === config.sheetName && cell.rowIndex >= 2) {
        this.sharedService.rememberDynamicClassAttributes(sheet.id, cell.rowIndex, attributes);
      }
    });
  }

  /** Binds the selected Characteristic to the Value cell in the same dynamic-sheet row. */
  private async rememberDynamicCharacteristicValue(attribute: any): Promise<void> {
    const config = dynamicClassificationSheet as any;
    if (!config?.sheetName || !attribute?.uuid) return;
    const valueColumn = config.fields?.find((field: any) => field.fieldId === 'VALUE')?.columnNumber;
    if (!Number.isFinite(valueColumn)) return;
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const cell = context.workbook.getActiveCell();
      sheet.load(['id', 'name']);
      cell.load('rowIndex');
      await context.sync();
      if (sheet.name === config.sheetName && cell.rowIndex >= 2) {
        this.sharedService.rememberRuntimeClassificationAttribute(sheet.id, cell.rowIndex, valueColumn, attribute);
      }
    });
  }

    getTranslatedFields(fieldDataV2: any[]) {
    // Apply translations (if available) to a cloned static fields array so templates reflect localized labels/groups
    const translatedStaticFields = (fieldDataV2 || []).map(fd => {
      const f = { ...fd } as any;

      // Translate label if it exists
      if (fd.label) {
        const translatedLabel = this.translationService.translate(fd.label);
        if (translatedLabel && translatedLabel !== fd.label) {
          f.label = translatedLabel;
        }
      }

      // Translate group if it exists
      if (fd.group) {
        const translatedGroup = this.translationService.translate(fd.group);
        if (translatedGroup && translatedGroup !== fd.group) {
          f.group = translatedGroup;
        }
      }

      return f;
    });
    return translatedStaticFields;
  }

  private async handleXfuseAction(action?: XfuseDaxeXlAction): Promise<void> {
    if (this.sharedService.isProcessRunning) {
      return;
    }

    switch (action) {
      case 'generateDescription':
        if (this.generationProfile.descriptionGeneration === true) {
          await this.daxeXlService.generateDescription(0, 0, null, true);
        }
        break;
      case 'applyDataRules':
        this.validateAndSubmit(RequestType.Defaults, false);
        break;
      case 'validateRecords':
        this.validateAndSubmit(RequestType.Validation, true);
        break;
      case 'submitRecords':
        this.validateAndSubmit(RequestType.Export, true);
        break;
      case 'sendForCorrection':
        if (this.enableSendForCorrection) {
          this.sendForCorrection();
        }
        break;
      case 'importRecordsFromExcel':
        if (this.canRequestSheetRecordImport()) {
          await this.importData();
        }
        break;
    }
  }

  async validateAndSubmit(requestType: RequestType, submitOrValidate: boolean) {
    if (this.sharedService.isProcessRunning) {
      return;
    }
    this.sharedService.isProcessRunning = true;
    if (this.shouldAutoImportBeforeValidation(requestType, submitOrValidate) && !this.sharedService.recordsImported) {
      const imported = await this.importData();
      if (!imported) {
        this.sharedService.isProcessRunning = false;
        return;
      }
    }
    if (submitOrValidate) {
      if (this.isMocrWorkbook) {
        await this.daxeXlService.validateMocrRecordsMass(this.sheetContextStore, requestType);
      } else if (this.isMultiSheetWorkbook) {
        await this.daxeXlService.validateWorkbookRecordsMass(this.sheetContextStore, requestType);
      } else {
        await this.daxeXlService.validateRecordsMass(requestType);
      }
    }
    else {
      if (this.isMultiSheetWorkbook) {
        await this.daxeXlService.applyWorkbookDataRules(this.sheetContextStore);
      } else {
        await this.daxeXlService.applyDataRules(this.isMocrWorkbook ? this.sheetContextStore : undefined);
      }
    }
  }

  /** Initialises each configured MOCR root as an independent parent worksheet. */
  private async startMocrWorkbook(flowData: FlowModel): Promise<void> {
    this.sharedService.isMocrRequest = true;
    this.sharedService.data.flowId = flowData.flowId;
    this.sharedService.data.stepId = flowData.stepId;
    this.mocrRootIndex = -1;
    for (const root of mocrRootContexts) {
      const context: SheetContext = {
        sheetName: root.sheetName,
        flowId: flowData.flowId,
        stepId: root.stepId || flowData.stepId,
        formId: root.layoutId,
        childFormId: '',
        schemaId: '',
        moduleId: root.moduleId,
        layoutId: root.layoutId,
        formDesc: root.formDesc,
        massId: '', userId: '', tenantId: '', mocrId: '',
        fieldsMap: new Map<string, any>(),
        staticFieldsObject: [],
        codeFields: [...root.codeFields],
        extensionHeaderFields: root.fieldData.map(field => field.fieldId),
        businessRules: [],
        baseStructure: [{ dataSetId: root.moduleId, formId: root.layoutId, StructureList: [{ structure: '1', sheetName: root.sheetName }]}]
      };
      this.sheetContextStore.set(root.sheetName, context);
    }
    await this.requestNextMocrRootLayout();
  }

  private async requestNextMocrRootLayout(): Promise<void> {
    this.mocrRootIndex++;
    const root = mocrRootContexts[this.mocrRootIndex];
    if (!root) {
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLBusinessRulesRequest, payload: { flowId: this.sharedService.data.flowId, stepId: this.sharedService.data.stepId } as any });
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLControlDataRequest });
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLGlobalConfigRequest });
      this.sharedService.initPersistence();
      this.validateSendForCorrection();
      return;
    }
    this.sheetContextStore.restoreSheetContext(root.sheetName, this.sharedService);
    // MOCR has independent roots. Tell ngx-mdo which workbook sheet owns this
    // layout so grid validation errors are mapped back to the same root sheet.
    daxeXlEvent.next({
      type: DaxeXLBaseEvents.DAXEXLLayoutDetailsRequest,
      payload: {
        layoutId: root.layoutId,
        moduleId: root.moduleId,
        rootSheetName: root.sheetName
      } as any
    });
  }

  /**
   * Initialises one normal form across several workbook sheets. Unlike MOCR
   * roots, these sheets contribute to the same record and are joined by
   * Record Number by DaxeXLService before the request is sent.
   */
  private async startMultiSheetWorkbook(flowData: FlowModel): Promise<void> {
    this.sharedService.isMocrRequest = false;
    this.sharedService.data.flowId = flowData.flowId;
    this.sharedService.data.stepId = flowData.stepId;
    this.multiSheetIndex = -1;
    for (const sheet of multiSheetContexts) {
      this.sheetContextStore.set(sheet.sheetName, {
        sheetName: sheet.sheetName,
        sheetRole: sheet.role,
        flowId: flowData.flowId,
        stepId: sheet.stepId || flowData.stepId,
        formId: sheet.layoutId,
        childFormId: '',
        schemaId: '',
        moduleId: sheet.moduleId,
        layoutId: sheet.layoutId,
        formDesc: sheet.formDesc,
        massId: '', userId: '', tenantId: '', mocrId: '',
        fieldsMap: new Map<string, any>(),
        staticFieldsObject: [],
        codeFields: [...sheet.codeFields],
        extensionHeaderFields: sheet.fieldData.map(field => field.fieldId),
        businessRules: [],
        baseStructure: [{ dataSetId: sheet.moduleId, formId: sheet.layoutId, StructureList: [{ structure: '1', sheetName: sheet.sheetName }]}]
      });
    }
    await this.requestNextMultiSheetLayout();
  }

  private async requestNextMultiSheetLayout(): Promise<void> {
    this.multiSheetIndex++;
    const sheet = multiSheetContexts[this.multiSheetIndex];
    if (!sheet) {
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLBusinessRulesRequest, payload: { flowId: this.sharedService.data.flowId, stepId: this.sharedService.data.stepId } as any });
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLControlDataRequest });
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLGlobalConfigRequest });
      this.sharedService.initPersistence();
      this.validateSendForCorrection();
      return;
    }
    this.sheetContextStore.restoreSheetContext(sheet.sheetName, this.sharedService);
    daxeXlEvent.next({
      type: DaxeXLBaseEvents.DAXEXLLayoutDetailsRequest,
      payload: { layoutId: sheet.layoutId, moduleId: sheet.moduleId, rootSheetName: sheet.sheetName } as any
    });
  }

  /** Writes each joined-record result to every source sheet that owns its row. */
  private async updateMultiSheetStatuses(records: MdoRecord[]): Promise<void> {
    await this.utilsService.toggleRuntimeEvents(false);
    try {
      for (const sheetName of this.sheetContextStore.getAllSheetNames()) {
        const sheetRecords = this.daxeXlService.getWorkbookResponseRecords(this.sheetContextStore, sheetName, records);
        if (!sheetRecords.length || !this.sheetContextStore.restoreSheetContext(sheetName, this.sharedService)) {
          continue;
        }
        await this.excelWriter.updateStatusAndMessages(sheetRecords, undefined, sheetName);
        this.sheetContextStore.snapshotContext(sheetName, this.sharedService);
      }
    } finally {
      await this.utilsService.toggleRuntimeEvents(true);
    }
  }

  /** Routes default/data-rule changes back to the sheet that contains the field. */
  private async updateMultiSheetDataRuleResponse(records: MdoRecord[]): Promise<void> {
    for (const sheetName of this.sheetContextStore.getAllSheetNames()) {
      const sheetRecords = this.daxeXlService.getWorkbookResponseRecords(this.sheetContextStore, sheetName, records);
      if (!sheetRecords.length || !this.sheetContextStore.restoreSheetContext(sheetName, this.sharedService)) {
        continue;
      }
      await this.excelWriter.parseAndwriteApplyDataRuleResponse(sheetRecords, sheetName);
      this.sheetContextStore.snapshotContext(sheetName, this.sharedService);
    }
  }

  private async updateMocrRootStatuses(records: MdoRecord[]): Promise<void> {
    const grouped = new Map<string, MdoRecord[]>();
    for (const record of records) {
      const key = this.getMocrContextKey(record);
      grouped.set(key, [...(grouped.get(key) || []), record]);
    }
    await this.utilsService.toggleRuntimeEvents(false);
    try {
      for (const sheetName of this.sheetContextStore.getAllSheetNames()) {
        const context = this.sheetContextStore.get(sheetName);
        const rootRecords = grouped.get(this.getMocrContextKey(context)) || [];
        if (!rootRecords.length || !this.sheetContextStore.restoreSheetContext(sheetName, this.sharedService)) continue;
        // Update the target sheet directly; do not switch the visible sheet.
        await this.excelWriter.updateStatusAndMessages(rootRecords, undefined, sheetName);
        this.sheetContextStore.snapshotContext(sheetName, this.sharedService);
      }
    } finally {
      await this.utilsService.toggleRuntimeEvents(true);
    }
  }

  /**
   * Sends the complete workbook structure to ngx-mdo. A validation response
   * can contain records for several MOCR roots, whereas SharedService holds
   * only the context most recently restored by the writer.
   */
  private openValidationErrorSideSheet(records: MdoRecord[]): void {
    daxeXlEvent.next({
      type: DaxeXLBaseEvents.DAXEXLViewErrorSideSheet,
      payload: {
        records,
        baseStructureModels: this.getWorkbookBaseStructures()
      }
    });
  }

  /**
   * A structure is `1` for a flat worksheet. Split workbook layouts may add
   * grid field IDs and hierarchy structure IDs for their own worksheets.
   */
  private getWorkbookBaseStructures(): any[] {
    const structuresByRoot = new Map<string, any>();

    for (const sheetName of this.sheetContextStore.getAllSheetNames()) {
      const context = this.sheetContextStore.get(sheetName);
      for (const baseStructure of context?.baseStructure || []) {
        const key = `${baseStructure.dataSetId}|${baseStructure.formId}`;
        const existing = structuresByRoot.get(key) || {
          dataSetId: baseStructure.dataSetId,
          formId: baseStructure.formId,
          StructureList: []
        };
        const knownStructures = new Set(existing.StructureList.map((item: any) => item.structure));

        for (const structure of baseStructure.StructureList || []) {
          if (!knownStructures.has(structure.structure)) {
            existing.StructureList.push({ ...structure });
            knownStructures.add(structure.structure);
          }
        }
        structuresByRoot.set(key, existing);
      }
    }

    return structuresByRoot.size
      ? Array.from(structuresByRoot.values())
      : this.sharedService.data.baseStructure || [];
  }

  /** Writes each root's data-rule result back to that root's worksheet only. */
  private async updateMocrRootDataRuleResponse(records: MdoRecord[]): Promise<void> {
    const grouped = new Map<string, MdoRecord[]>();
    for (const record of records || []) {
      const key = this.getMocrContextKey(record);
      grouped.set(key, [...(grouped.get(key) || []), record]);
    }
    for (const sheetName of this.sheetContextStore.getAllSheetNames()) {
      const context = this.sheetContextStore.get(sheetName);
      const rootRecords = grouped.get(this.getMocrContextKey(context)) || [];
      if (!rootRecords.length || !this.sheetContextStore.restoreSheetContext(sheetName, this.sharedService)) {
        continue;
      }
      await this.excelWriter.parseAndwriteApplyDataRuleResponse(rootRecords, sheetName);
      this.sheetContextStore.snapshotContext(sheetName, this.sharedService);
    }
  }

  /**
   * Backend responses normally return moduleId/layoutId. Parent/dataset/form
   * fallbacks keep response routing correct for equivalent response shapes.
   */
  private getMocrContextKey(source: any): string {
    const controlData = source?.controlData || source || {};
    const moduleId = controlData.moduleId || controlData.parentModuleId || controlData.datasetId || '';
    const layoutId = controlData.layoutId || controlData.parentLayoutId || controlData.formId || '';
    return `${String(moduleId)}|${String(layoutId)}`;
  }

  sendForCorrection() {
    this.daxeXlService.sendForCorrection();
  }

  /**
 * unsubscribing the events on component destroy
 */
  ngOnDestroy(): void {
    console.log("ngDestroy Called upon URL removed")
    if (this.subscription) {
      this.subscription.unsubscribe();
    }
    void this.unregisterWorksheetActivationHandler();
  }

  /**
   * Updating the host toolbar on activation must not emulate a cell selection:
   * doing that would reset dropdown state in ngx-mdo. The sheet context is
   * already the authoritative form/module source for multi-sheet workbooks.
   */
  private async registerWorksheetActivationHandler(): Promise<void> {
    if (this.worksheetActivationRegistered) return;
    try {
      await Office.onReady();
      await Excel.run(async (context: any) => {
        context.workbook.worksheets.onActivated.add(this.worksheetActivatedHandler);
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load('id');
        await context.sync();
        this.worksheetActivationRegistered = true;
        await this.publishDescriptionGeneratorState(activeSheet.id);
      });
    } catch (error) {
      console.warn('Unable to register DAXE XL worksheet activation handler.', error);
    }
  }

  private async unregisterWorksheetActivationHandler(): Promise<void> {
    if (!this.worksheetActivationRegistered) return;
    try {
      await Excel.run(async (context: any) => {
        context.workbook.worksheets.onActivated.remove(this.worksheetActivatedHandler);
        await context.sync();
      });
    } catch (error) {
      // Office can dispose the workbook before the custom element. There is
      // nothing to recover in that case, and the element subscription is gone.
      console.debug('Unable to remove DAXE XL worksheet activation handler.', error);
    } finally {
      this.worksheetActivationRegistered = false;
    }
  }

  private async publishDescriptionGeneratorState(worksheetId?: string): Promise<void> {
    try {
      let sheetName = '';
      await Excel.run(async (context: any) => {
        const sheet = worksheetId
          ? context.workbook.worksheets.getItem(worksheetId)
          : context.workbook.worksheets.getActiveWorksheet();
        sheet.load('name');
        await context.sync();
        sheetName = sheet.name;
      });
      const sheetContext = this.sheetContextStore.get(sheetName);
      const formId = sheetContext?.formId || this.sharedService.data.formId;
      daxeXlEvent.next({
        type: DAXE_XL_DESCRIPTION_GENERATOR_STATE_REQUEST as any,
        payload: {
          formId,
          // The currently generated sheet-context contract has only form-data
          // roles (header/grid/hierarchy/record). Do not compare it with a
          // future-only "classification" role: Angular correctly rejects that
          // impossible union comparison during a generated-project build.
          isDataSheet: true
        }
      } as any);
    } catch (error) {
      console.debug('Unable to publish DAXE XL description-generator state.', error);
    }
  }

  /** Validates if "Send for Correction" should be enabled by checking for MDOConfiguration sheet */
  async validateSendForCorrection(): Promise<void> {
    try {
      await Office.onReady();
      await Excel.run({ delayForCellEdit: true }, async (context) => {
        const worksheets = context.workbook.worksheets;
        worksheets.load("items/name");
        await context.sync();
        this.isSendForCorrectionVisible = worksheets.items.map(sheet => sheet.name).includes("MDOConfiguration");
      });
    } catch (error) {
      console.error("Error in validateSendForCorrection: ", error);
      this.isSendForCorrectionVisible = false;
    }
  }
  get enableSendForCorrection(): boolean {
    return this.generationProfile.process === 'approve'
      && this.isSendForCorrectionVisible
      && this.utilsService?.globalConfig?.isSendForCorrectionEnable;
  }

  private canRequestSheetRecordImport(): boolean {
    return this.generationProfile.process === 'change'
      && (this.generationProfile.autoImportBeforeValidation === true
        || (this.generationProfile.changeImportModes || []).includes('sheet-record-number'));
  }

  private shouldAutoImportBeforeValidation(requestType: RequestType, submitOrValidate: boolean): boolean {
    return submitOrValidate
      && requestType !== RequestType.Defaults
      && this.generationProfile.process === 'change'
      && this.generationProfile.autoImportBeforeValidation === true;
  }

  async importData(): Promise<boolean> {
    try {
      this.utilsService.displayStatus(this.translationService.translate('Preparing records for {{requestType}}...', { requestType: String(RequestType.Import) }));
      const values = await this.excelReader.getSapNumbers()

      if (values && values.length == 0) {
        this.utilsService.displayStatus(null);
        this.sharedService.isProcessRunning = false;
        return false;
      }
      const payload = {
        "structures": this.sharedService.data.uniqueStructures,
        "filters": [
          {
            "moduleId": this.sharedService.data.controlData.moduleId,
            "fieldId": "mdo_number",
            "fieldDescri": "Record Number",
            "fieldType": "text",
            "dataType": "ALTN",
            "operator": "EQUAL",
            "values": values,
            "startValue": null,
            "endValue": null,
            "esFieldPath": "hdvs.mdo_number",
            "structureId": 1
          }
        ]
      }

      const response = new Promise<boolean>((resolve) => {
        this.importResponseResolver = resolve;
      });
      daxeXlEvent.next({
        type: DaxeXLBaseEvents.DAXEXLImportDataRequest,
        payload: payload
      });
      // The host returns selected records through XFUSEImportDataResponse. Do
      // not validate a Change element until that payload has populated it.
      const timeout = new Promise<boolean>((resolve) => setTimeout(() => resolve(false), 30000));
      const imported = await Promise.race([response, timeout]);
      if (!imported) {
        this.importResponseResolver = undefined;
        await this.utilsService.displayStatus(null);
      }
      return imported;
    } catch (error) {
      this.sharedService.isProcessRunning = false;
      console.log(error)
      this.importResponseResolver?.(false);
      this.importResponseResolver = undefined;
      return false;
    }
  }
}

export interface FlowModel {
  flowId: string;
  flowDesc: string;
  stepId: string;
  parentMapping: { [ModuleId: string]: FormMapping };
  childMapping: { [ModuleId: string]: FormMapping }[];
  isRestrictedModules: false
  isSendForCorrection: boolean;
  isMocrEnabled?: boolean;
}
export interface FormMapping {
  datasetDesc: string;
  formId: string;
  formDesc: string;
  refernceDataSet: RefernceDataSet[];
}
export interface RefernceDataSet {
  referenceDatasetId: string;
  formId: string;
  refdataSetDesc: string;
  formDesc: string;
}
