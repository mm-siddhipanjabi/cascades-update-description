import { Injectable } from '@angular/core';
import { ClassificationDescription, ControlData, daxeXlEvent, DaxeXlEventType, Field, FieldTypes, Grid, MdoRecord, MdoRecordES, RecordField, RequestType, Structure, Value } from 'daxe-ui-sdk';
import { Subject } from 'rxjs';
import { SharedService } from './shared.service';
import { UtilsService } from './utils.service';
import { TranslationService } from './translation.service';
import { Md5 } from 'ts-md5';

@Injectable({
  providedIn: 'root',
})
export class ExcelReaderService {
  event: Subject<DaxeXlEventType> = daxeXlEvent;
  private readonly ROW_BATCH_SIZE = 5000;
  private readonly RECORD_PUSH_BATCH = 5000;

  constructor(private readonly sharedService: SharedService, private readonly utilsService: UtilsService, private readonly translationService: TranslationService) { }

  /**
   * Reads records from the supplied worksheet when one is provided. MOCR uses
   * this to read every root sheet without changing the user's active sheet.
   */
  async buildRecordList(requestType: RequestType, sheetName?: string) {
    let recordList: MdoRecord[] = [];
    try {
      await this.processSheetRecordWise(requestType, async ({ totalRecords, batchRecordCount, recordList: batchRecordList }) => {
        if (batchRecordList && batchRecordList.length > 0) {
          recordList.push(...batchRecordList);
        }
      }, sheetName);
    }
    catch (ex) {
      console.error('Error processing sheet record wise:', ex);
      return null;
    }

    return recordList;
  }
  async PrepareRecordsToSendForCorrection() {
    let recordsJSON: Array<MdoRecord> = [];
    await Excel.run({ delayForCellEdit: true }, async (ctx) => {
      const sheet: Excel.Worksheet = ctx.workbook.worksheets.getActiveWorksheet();
      let usedRange = sheet.getUsedRange(true);
      usedRange.load(["columnCount", "rowCount"]);
      await ctx.sync();
      let selectedRanges = ctx.workbook.getSelectedRanges();
      selectedRanges.areas.load(['items']);
      await ctx.sync();
      for (let selectedRange of selectedRanges.areas.items) {
        selectedRange.load(["rowCount", "rowIndex", "columnCount"]);
        let selectedRowcount = 0;
        let selectedColumCount = 0;
        await ctx.sync();
        selectedRowcount = selectedRange.rowCount;
        selectedColumCount = selectedRange.columnCount
        let newRecordJson: MdoRecord[] = [];
        if (selectedColumCount > 6) {
          newRecordJson = await this.prepareSendForCorrectionRecordArray(sheet, selectedRange.rowIndex, selectedRowcount, usedRange.columnCount);
        }
        else if (usedRange.rowCount > 2) {
          newRecordJson = await this.prepareSendForCorrectionRecordArray(sheet, 2, usedRange.rowCount - 2, usedRange.columnCount);
        }
        recordsJSON = [...recordsJSON, ...newRecordJson];
      }

      return ctx.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });

    return recordsJSON;
  }

  async prepareSendForCorrectionRecordArray(sheet: Excel.Worksheet, startRow: number, Rowcount: number, ColumCount: number) {
    let recordList: MdoRecord[] = []
    let ctx = sheet.context;
    let dataRange = sheet.getRangeByIndexes(startRow, 0, Rowcount, ColumCount);
    dataRange.load("text");
    let recordsData: string[][] = [[]];
    await ctx.sync().then(() => { recordsData = dataRange.text; });
    let recordIndex = startRow - 1;
    for (const dataRow of recordsData) {
      let recNo = dataRow[2];
      if (recordList.find(m => m.controlData.recordNumber == recNo)) {
        continue;
      }
      //#region  Building controldata
      const recordExist = this.sharedService.records?.find(record => record.controlData.recordNumber == recNo);
      let mdorecord: MdoRecord = new MdoRecord();
      let controlData: ControlData = {
        ...recordExist?.controlData || this.sharedService.controlDataMap.get(recNo) || {},
      } as ControlData
      const data = this.sharedService.data;
      controlData.moduleId = data.controlData.moduleId;
      controlData.processFlowId = data.flowId;
      controlData.massId = data.controlData.massId;
      controlData.stepId = data.stepId;
      controlData.layoutId = data.formId || '';
      controlData.tenantId = data.controlData.tenantId;
      controlData.userId = data.controlData.userId;
      (controlData as any).mocrId = (data.controlData as any).mocrId || '';
      controlData.excelRow = recordIndex
      mdorecord.controlData = controlData;
      //#endregion Building controldata
      recordList.push(mdorecord);
      recordIndex++
    }
    return recordList;
  }

  async buildIndividualRecordsFromRows(
    recNo: string,
    rows: string[][],
    rowIndexes: number[],
    requestType: RequestType,
    classDetails: Map<string, ClassificationDescription>,
    isModuleReference: boolean = false
  ): Promise<MdoRecord[]> {
    const recordList: MdoRecord[] = [];
    const recordExist = this.sharedService.records?.find(record => record.controlData.recordNumber == recNo);

    for (const [rowIdx, row] of rows.entries()) {
      try {
        console.log("preparing individual record for", row);
        const excelRowNumber = rowIndexes[rowIdx] + 1;

        // Create fresh hdvs, hyvs, gvs for this row only
        const hdvs: { [FieldId: string]: RecordField } = this.sharedService.isExtentionProcess ? (recordExist?.mdoRecordES.hdvs || {} as unknown as {
          [FieldId: string]: RecordField
        }) : {} as unknown as {
          [FieldId: string]: RecordField;
        };
        const hyvs: { [structureId: string]: Structure } = {} as unknown as { [structureId: string]: Structure };
        const gvs: { [GridId: string]: Grid } = {} as unknown as { [GridId: string]: Grid };

        const hyvsRowObjs: Record<number, any> = {};
        const gvsRowObjs: Record<string, any> = {};

        // Process all fields for this row
        for (const field of this.sharedService.staticFieldsObject) {
          const colIndex = field.columnNumber;
          const cellValue = row[colIndex] ?? '';
          const fieldMeta: Field | undefined = this.sharedService.fieldsMap.get(field.fieldId);
          if (!fieldMeta) continue;

          const structId = Number(fieldMeta.metadata?.[0]?.structureId);
          const isKeyField = fieldMeta.metadata?.[0]?.isKeyField ?? false;
          const isGridField = fieldMeta.metadata?.[0]?.isGridColumn ?? false;
          const parentField = fieldMeta.metadata?.[0]?.parentField ?? '';

          if (this.isHdvsField(fieldMeta, structId, isGridField)) {
            await this.processHdvsField(hdvs, fieldMeta, cellValue, recNo);
          } else if (isGridField) {
            await this.processGvsField(gvsRowObjs, fieldMeta, cellValue, recNo, parentField);
          } else if (structId !== 1 && !isGridField) {
            await this.processHyvsField(hyvsRowObjs, hyvs, fieldMeta, cellValue, recNo, structId, isKeyField);
          }
        }

        hdvs['ROW_UUID'] = { fId: 'ROW_UUID', vc: [{ c: excelRowNumber, t: recNo }] } as RecordField;

        // Append hyvs and gvs rows for this single row
        this.appendHyvsRows(hyvs, hyvsRowObjs, excelRowNumber);
        this.appendGvsRows(gvs, gvsRowObjs, excelRowNumber);

        // Keep the complete imported/generated descriptions payload. Workbook
        // columns expose only the configured class fields, so rebuilding this
        // array from the visible cells would erase classes/attributes not shown
        // by this element (especially in Change flows).
        const isGenerateDescription = (requestType as any) === 'generateDescriptions';
        const importedDescriptions = recordExist?.mdoRecordES.descriptions || [];
        // Generate Description is a snapshot of the workbook. It must never
        // take a class, an attribute value, or description text from retained
        // imported/generated state. That state is only for Validate/Submit.
        let storedDescriptions = isGenerateDescription
          ? []
          : this.sharedService.getDescriptions(recNo, importedDescriptions);
        const selectedClassColumn = this.getConfiguredClassificationField()
          ? this.sharedService.staticFieldsObject.find(field => field.fieldId === this.getConfiguredClassificationField()?.fieldId)?.columnNumber
          : undefined;
        if (!isGenerateDescription && this.sharedService.classificationEnabled
          && selectedClassColumn !== undefined
          && !String(row[selectedClassColumn] ?? '').trim()) {
          this.sharedService.clearDescriptionClass(recNo);
          storedDescriptions = this.sharedService.getDescriptions(recNo);
        }
        // Static classifications are business data for validation/submit.
        // They must not replace the Description Class context supplied to the
        // Generate Description API.
        const staticDescriptions = isGenerateDescription
          ? []
          : this.buildStaticClassDescriptions(row, storedDescriptions);
        if (staticDescriptions.length) {
          this.sharedService.mergeDescriptions(recNo, staticDescriptions);
        }
        if (this.sharedService.classificationEnabled && isGenerateDescription) {
          this.sharedService.clearGeneratedDescription(recNo);
        }
        let descriptionClass = this.sharedService.classificationEnabled
          ? (isGenerateDescription
            ? classDetails.get(recNo)
            : this.sharedService.generatedDescriptions.get(recNo) || classDetails.get(recNo)) as ClassificationDescription
          : undefined;
        if (descriptionClass) {
          descriptionClass = { ...descriptionClass, isDesc: true } as ClassificationDescription;
        }
        descriptionClass = this.mergeTemplateDescriptionValues(
          row,
          descriptionClass,
          isGenerateDescription ? [] : storedDescriptions
        );
        // Generate Description is driven exclusively by the Description Class
        // chosen in this workbook. Existing/static classifications belong to
        // the validation/export payload and must not become the generator's
        // class context.
        const descriptions = this.sharedService.classificationEnabled
          ? (isGenerateDescription
            ? (descriptionClass ? [descriptionClass] : [])
            : this.sharedService.mergeWorkbookClass(recNo, descriptionClass))
          : importedDescriptions;

        // Create individual record for this row
        const record = {
          childRecord: [], // No child records in individual row records
          mdoRecordES: {
            hdvs,
            hyvs,
            gvs,
            descriptions,
            id: `${recNo}_${rowIdx}`,
            strId: 1,
          } as MdoRecordES,
          controlData: {
            ...this.sharedService.controlDataMap.get(recNo) || recordExist?.controlData || {},
          } as ControlData,
        };
        const data = this.sharedService.data;
        record.controlData.arbitraryRecordNo = recNo;
        record.controlData.moduleId = data.controlData.moduleId;
        record.controlData.processFlowId = data.flowId;
        record.controlData.massId = data.controlData.massId;
        record.controlData.stepId = data.stepId;
        record.controlData.layoutId = data.formId || '';
        record.controlData.tenantId = data.controlData.tenantId;
        record.controlData.userId = data.controlData.userId;
        (record.controlData as any).mocrId = (data.controlData as any).mocrId || '';
        record.controlData.excelRow = excelRowNumber;
        record.controlData.validationRequired = true;

        recordList.push(record);
      } catch (ex) {
        console.error(`Error processing row ${rowIdx}:`, ex);
      }
    }

    return recordList;
  }

  async buildRecordJson(
    recNo: string,
    rows: string[][],
    rowIndexes: number[],
    requestType: RequestType,
    classDetails: Map<string, ClassificationDescription>,
    isModuleReference: boolean = false
  ): Promise<MdoRecord> {
    const recordExist = this.sharedService.records?.find(record => record.controlData.recordNumber == recNo);
    const hdvs: { [FieldId: string]: RecordField } = this.sharedService.isExtentionProcess ? (recordExist?.mdoRecordES.hdvs || {} as unknown as {
      [FieldId: string]: RecordField
    }) : {} as unknown as {
      [FieldId: string]: RecordField;
    };
    const hyvs: { [structureId: string]: Structure } = {} as unknown as {
      [structureId: string]: Structure
    };
    const gvs: { [GridId: string]: Grid } = {} as unknown as {
      [GridId: string]: Grid
    };
    let index = 0;
    const hierarchyContextRowObjs: Record<number, any> = {};
    for (const [rowIdx, row] of rows.entries()) {
      try {
        console.log("preparing for", row)
        const excelRowNumber = rowIndexes[index] + 1;
        const hyvsRowObjs: Record<number, any> = {};
        const gvsRowObjs: Record<string, any> = {};
        for (const field of this.sharedService.staticFieldsObject) {
          const colIndex = field.columnNumber;
          const cellValue = row[colIndex] ?? '';
          const fieldMeta: Field | undefined = this.sharedService.fieldsMap.get(field.fieldId);
          if (!fieldMeta) continue;

          const structId = Number(fieldMeta.metadata?.[0]?.structureId);
          const isKeyField = fieldMeta.metadata?.[0]?.isKeyField ?? false;
          const isGridField = fieldMeta.metadata?.[0]?.isGridColumn ?? false;
          const parentField = fieldMeta.metadata?.[0]?.parentField ?? '';
          if (index == 0 && this.isHdvsField(fieldMeta, structId, isGridField)) {
            await this.processHdvsField(hdvs, fieldMeta, cellValue, recNo);
          } else if (isGridField) {
            await this.processGvsField(
              gvsRowObjs,
              fieldMeta,
              cellValue,
              recNo,
              parentField
            );
          } else if (structId !== 1 && !isGridField) {
            await this.processHyvsField(
              hyvsRowObjs,
              hyvs,
              fieldMeta,
              cellValue,
              recNo,
              structId,
              isKeyField
            );
          }
        }

        await this.addParentKeyFieldsToHyvs(hyvsRowObjs, hyvs);
        this.updateHierarchyContext(hierarchyContextRowObjs, hyvsRowObjs);
        await this.addParentKeyFieldsToGvs(hyvsRowObjs, gvsRowObjs, hierarchyContextRowObjs);

        // ROW_UUID is a platform row identifier. Business-field values come
        // exclusively from the generated metadata for the selected form.
        hdvs['ROW_UUID'] = { fId: 'ROW_UUID', vc: [{ c: rowIndexes[0] + 1 }] } as RecordField;

        this.appendHyvsRows(hyvs, hyvsRowObjs, excelRowNumber);
        this.appendGvsRows(gvs, gvsRowObjs, excelRowNumber);
      } catch (ex) {
        console.error(ex);
      }
      index++
    }
    // See the row-wise reader above: source data not represented as Excel
    // columns must survive each validation/submit serialization.
    const isGenerateDescription = (requestType as any) === 'generateDescriptions';
    const importedDescriptions: ClassificationDescription[] = recordExist?.mdoRecordES.descriptions || [];
    // A Generate Description request must reflect only current worksheet
    // values. Retained descriptions are deliberately used only for requests
    // that need to preserve imported, non-visible classification fields.
    let storedDescriptions = isGenerateDescription
      ? []
      : this.sharedService.getDescriptions(recNo, importedDescriptions);
    const selectedClassColumn = this.getConfiguredClassificationField()
      ? this.sharedService.staticFieldsObject.find(field => field.fieldId === this.getConfiguredClassificationField()?.fieldId)?.columnNumber
      : undefined;
    if (!isGenerateDescription && this.sharedService.classificationEnabled
      && selectedClassColumn !== undefined
      && !String((rows[0] || [])[selectedClassColumn] ?? '').trim()) {
      this.sharedService.clearDescriptionClass(recNo);
      storedDescriptions = this.sharedService.getDescriptions(recNo);
    }
    // Static class characteristics are normal form fields, not a synthetic
    // classification sheet. Convert only the configured static fields into a
    // description entry and merge it with imported/generated state. Fields not
    // represented by this workbook remain untouched in the stored payload.
    // Generate Description uses the selected Description Class only. Static
    // class attributes are preserved for validation/export, not sent as the
    // description-generator classification context.
    const staticDescriptions = isGenerateDescription
      ? []
      : this.buildStaticClassDescriptions(rows[0] || [], storedDescriptions);
    if (staticDescriptions.length) {
      this.sharedService.mergeDescriptions(recNo, staticDescriptions);
    }
    if (this.sharedService.classificationEnabled && isGenerateDescription) {
      this.sharedService.clearGeneratedDescription(recNo);
    }
    let descriptionClass = this.sharedService.classificationEnabled
      ? (isGenerateDescription
        ? classDetails.get(recNo)
        : this.sharedService.generatedDescriptions.get(recNo) || classDetails.get(recNo)) as ClassificationDescription
      : undefined;
    if (descriptionClass) {
      descriptionClass = { ...descriptionClass, isDesc: true } as ClassificationDescription;
    }
    descriptionClass = this.mergeTemplateDescriptionValues(
      rows[0] || [],
      descriptionClass,
      isGenerateDescription ? [] : storedDescriptions
    );
    // See the row-wise reader above. The Generate Description endpoint must
    // receive only the selected Description Class, never static/imported
    // classification entries from this record.
    const descriptions = this.sharedService.classificationEnabled
      ? (isGenerateDescription
        ? (descriptionClass ? [descriptionClass] : [])
        : this.sharedService.mergeWorkbookClass(recNo, descriptionClass))
      : importedDescriptions;
    const record = {
      childRecord: [],
      mdoRecordES: {
        hdvs,
        hyvs,
        gvs,
        descriptions,
        id: recNo,
        strId: 1,
      } as MdoRecordES,
      controlData: {
        ...this.sharedService.controlDataMap.get(recNo) || recordExist?.controlData || {},
      } as ControlData,
    };

    const data = this.sharedService.data;
    record.controlData.arbitraryRecordNo = recNo;
    record.controlData.moduleId = data.controlData.moduleId;
    record.controlData.processFlowId = data.flowId;
    record.controlData.massId = data.controlData.massId;
    record.controlData.stepId = data.stepId;
    record.controlData.layoutId = data.formId || '';
    record.controlData.tenantId = data.controlData.tenantId;
    record.controlData.userId = data.controlData.userId;
    (record.controlData as any).mocrId = (data.controlData as any).mocrId || '';
    record.controlData.excelRow = rowIndexes[0] + 1;
    record.controlData.validationRequired = true;
    return record;
  }
  private isHdvsField(
    field: Field,
    structId: number,
    isGridField: boolean
  ): boolean {
    return (
      structId === 1 &&
      !isGridField &&
      field.fieldType != FieldTypes.Static &&
      !field.metadata[0].grid &&
      !field.isStaticClassCharacteristics &&
      !field.isStaticClassCharacteristicsUoM &&
      field.fieldType != 'CLASS_TYPE' &&
      field.fieldType != FieldTypes.Reference
    );
  }

  private async processHdvsField(
    hdvs: { [FieldId: string]: RecordField },
    field: Field,
    cellValue: string,
    recNo: string
  ) {
    if (cellValue !== '' || !hdvs[field.fieldId]) {
      hdvs[field.fieldId] = {
        vc: await this.prepareValueCollection(field, cellValue, recNo),
        fId: field.fieldId,
        ls: field.description,
      } as RecordField;
    }
  }

  private async processGvsField(
    gvsRowObjs: Record<string, any>,
    field: Field,
    cellValue: string,
    recNo: string,
    parentField: string
  ) {
    // MOCR needs an explicit empty value so a blank grid cell is validated and
    // can clear an existing value. Existing non-MOCR payloads keep their
    // historical sparse-grid behaviour.
    if (cellValue !== '' || this.sharedService.isMocrRequest) {
      if (!gvsRowObjs[parentField]) gvsRowObjs[parentField] = {};
      gvsRowObjs[parentField][field.fieldId] = {
        vc: await this.prepareValueCollection(field, cellValue, recNo),
        fId: field.fieldId,
        ls: field.metadata[0].description,
      };
    }
  }

  private async processHyvsField(
    hyvsRowObjs: Record<number, any>,
    hyvs: { [structureId: string]: Structure },
    field: Field,
    cellValue: string,
    recNo: string,
    structId: number,
    isKeyField: boolean
  ) {
    if (cellValue !== '') {
      if (!hyvsRowObjs[structId]) hyvsRowObjs[structId] = {};
      hyvsRowObjs[structId][field.fieldId] = {
        vc: await this.prepareValueCollection(field, cellValue, recNo),
        fId: field.fieldId,
        ls: field.description,
      };

      if (isKeyField) {
        if (!hyvs[structId]) {
          hyvs[structId] = { rows: [], keyFields: [], def: {} };
        }
        if (!hyvs[structId].keyFields.includes(field.fieldId)) {
          hyvs[structId].keyFields.push(field.fieldId);
        }
      }
    }
  }

  private async addParentKeyFieldsToHyvs(
    hyvsRowObjs: Record<number, any>,
    hyvs: { [structureId: string]: Structure },
  ): Promise<void> {
    for(let structId in hyvs) {
      let currentStructureFields = Array.from(this.sharedService.fieldsMap.values()).filter(f => f?.metadata[0]?.structureId == Number(structId)) as Field[];
      if(currentStructureFields.length > 0) {
        let parentStructure = currentStructureFields[0]?.metadata[0]?.parentHierarchy;
        if(parentStructure !== currentStructureFields[0]?.metadata[0]?.structureId) {
          const parentKeyFields = Array.from(this.sharedService.fieldsMap.values()).filter(f => f?.metadata[0]?.structureId == parentStructure && f?.metadata[0]?.isKeyField) as Field[];
          for(const parentKeyField of parentKeyFields) {
            if(!hyvsRowObjs[Number(structId)]) {
              continue;
            }
            if(!hyvsRowObjs[Number(structId)][parentKeyField.fieldId]) {
              hyvsRowObjs[Number(structId)][parentKeyField.fieldId] = hyvsRowObjs[Number(parentStructure)]?.[parentKeyField.fieldId];
            }
          }
        }
      }
    }
  }

  private async addParentKeyFieldsToGvs(
    hyvsRowObjs: Record<number, any>,
    gvsRowObjs: Record<string, any>,
    hierarchyContextRowObjs: Record<number, any> = {},
  ): Promise<void> {
    for(let gridId in gvsRowObjs) {
      const gridMeta = this.sharedService.fieldsMap.get(gridId);
      const gridParentStructure = gridMeta?.metadata[0]?.structureId;
      const parentStructurekeyFields = Array.from(this.sharedService.fieldsMap.values()).filter(f => f?.metadata[0]?.structureId == gridParentStructure && f?.metadata[0]?.isKeyField) as Field[];
      const parentRowObj = hyvsRowObjs[Number(gridParentStructure)] || hierarchyContextRowObjs[Number(gridParentStructure)];

      for (const parentKeyField of parentStructurekeyFields) {
        if (!gvsRowObjs[gridId]) {
          continue;
        }
        if (!gvsRowObjs[gridId][parentKeyField.fieldId] && this.hasRecordFieldValue(parentRowObj?.[parentKeyField.fieldId])) {
          gvsRowObjs[gridId][parentKeyField.fieldId] = parentRowObj[parentKeyField.fieldId];
        }
      }

    }
  }

  private updateHierarchyContext(
    hierarchyContextRowObjs: Record<number, any>,
    hyvsRowObjs: Record<number, any>,
  ): void {
    for (const structId of Object.keys(hyvsRowObjs)) {
      const rowObj = hyvsRowObjs[Number(structId)];
      if (this.hasAnyRecordFieldValue(rowObj)) {
        hierarchyContextRowObjs[Number(structId)] = {
          ...(hierarchyContextRowObjs[Number(structId)] || {}),
          ...rowObj,
        };
      }
    }
  }

  private hasAnyRecordFieldValue(rowObj: any): boolean {
    return !!rowObj && Object.values(rowObj).some(field => this.hasRecordFieldValue(field as RecordField));
  }

  private hasRecordFieldValue(field: RecordField | undefined): boolean {
    return !!field?.vc?.some(value => value?.c !== undefined && value?.c !== null && value?.c !== '');
  }

  private appendHyvsRows(
    hyvs: { [structureId: string]: Structure },
    hyvsRowObjs: Record<number, any>,
    excelRowNumber: number
  ) {
    Object.keys(hyvsRowObjs).forEach((sIdStr) => {
      if (!hyvs[sIdStr]) hyvs[sIdStr] = { rows: [], keyFields: [], def: {} };
      hyvs[sIdStr].rows.push({
        ...hyvsRowObjs[+sIdStr],
        ROW_UUID: {
          vc: [{ c: excelRowNumber.toString() }],
          fId: 'ROW_UUID',
          ls: 'excel row',
        },
      });
    });
  }

  private appendGvsRows(
    gvs: { [GridId: string]: Grid },
    gvsRowObjs: Record<string, any>,
    excelRowNumber: number
  ) {
    Object.keys(gvsRowObjs).forEach((parentField) => {
      if (!gvs[parentField]) gvs[parentField] = new Grid();
      gvs[parentField].rows.push({
        ...gvsRowObjs[parentField],
        ROW_UUID: {
          vc: [{ c: excelRowNumber.toString() }],
          fId: 'ROW_UUID',
          ls: 'excel row',
        },
        UUID: {
          vc: [{ c: crypto.randomUUID() }],
          fId: 'UUID',
          ls: 'UUID',
        },
      });
    });
  }

  async prepareValueCollection(
    field: Field,
    cellValue: string,
    ParentFilterKey: string,
    validation: boolean = true
  ): Promise<Value[]> {
    let value: Value = {
      t: null,
      c: null,
      u: null,
      isInError: false,
      errmsgs: [],
      excelRow: 0,
    };

    value.c = cellValue;
    if (field.metadata?.[0]?.dataType === 'DEC' && typeof value.c === 'string') {
      value.c = value.c.replace(",", ".");
    }
    let ValueCollection: Value[] = [];
    ValueCollection.push(value);

    if (field.metadata[0].isCheckList) {
      let values = cellValue
        .split(",")
        .map((m) => {
          return { c: m } as Value;
        });
      ValueCollection = values;
    }
    if (field.metadata[0].pickList == '38') {
      let attachMentFileNameList = cellValue?.split('|').map((m) => {
        return { fileName: m };
      });
      try {
        const globalVariableString = localStorage.getItem('GlobalVariables');
        let GlobalVariables = globalVariableString ? JSON.parse(globalVariableString) : {};
        ValueCollection = attachMentFileNameList.map(m => {
          const fileObj = GlobalVariables.AttachmentList.find((a: any) => a.fileName == m.fileName);
          return { c: fileObj?.id, t: fileObj?.fileName } as Value;
        });
      }
      catch (ex) {
        console.error('Error processing attachments from localStorage:', ex);
      }

    }
    if (field.metadata[0].pickList == "52") {
      const newCellValue = await this.utilsService.convertDateToMilliseconds(cellValue) || '';
      ValueCollection = [{ c: newCellValue } as Value];
    }

    if (["1", "4", "18", "37"].includes(field.metadata[0].pickList || '') && !this.sharedService.codeFields.includes(field.fieldId) && cellValue) {
      const newCellValue = await this.sharedService.getDropDownCodeOrText({ field: field, fieldId: field.fieldId, cellValue: cellValue, datsetId: field.metadata?.[0]?.moduleId as string, formId: this.sharedService.data.formId });
      if (newCellValue) {
        console.log("New Cell Value::", newCellValue, " for Field::", field.fieldId, " Cell Value::", cellValue);

        ValueCollection = [{ c: newCellValue } as Value];
      }
      else {
        ValueCollection = [{ c: cellValue } as Value];
      }
    }
    return ValueCollection;
  }

  async processSheetRecordWise(
    requestType: RequestType,
    pushCallback: (batchJson: ({ totalRecords: any, batchRecordCount: any, recordList: any })) => Promise<void>,
    sheetName?: string
  ) {
    await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const usedRange = sheet.getUsedRange();
      usedRange.load(['rowCount', 'columnCount']);
      await context.sync();

      const totalRows = usedRange.rowCount;
      const totalCols = usedRange.columnCount;
      if(totalRows <= 2) {return;} // No data rows to process

      const objectColumnLetter = 'C'
      const recordCol = sheet.getRange(
        `${objectColumnLetter}3:${objectColumnLetter}${totalRows}`
      );
      recordCol.load('values');
      await context.sync();
      const recordNumbers = recordCol.values.map((r) => r[0]);
      const missingRecordNumber = recordNumbers.some(val => val === "");

      if (missingRecordNumber) {
        recordNumbers.forEach((val, index) => {
          if (val === '') {
            const excelRow = index + 3;
            const statusCell = sheet.getRange(`A${excelRow}`);
            const errorCell = sheet.getRange(`B${excelRow}`);
            const mdoNumberCell = sheet.getRange(`C${excelRow}`);
            const errorText = this.translationService.translate('RecordNumberMissing');
            const statusText = this.translationService.translate('Error').toUpperCase();
            statusCell.values = [[statusText]];
            errorCell.values = [[errorText]];
            mdoNumberCell.format.fill.color = "#ffcccc";
          }
        });

        await context.sync();
        return;
      }

      const recordLastIndexMap = this.getRecordLastIndexMap(recordNumbers);
      const incompleteRecords = new Map<
        string,
        { rows: string[][]; rowIndexes: number[] }
      >();
      const pushBuffer: MdoRecord[] = [];
      let startRow = 2; // zero based ro index
      let classDetails: Map<string, ClassificationDescription> = new Map<string, ClassificationDescription>();

      while (startRow < totalRows) {
        const rowCount = Math.min(this.ROW_BATCH_SIZE, totalRows - startRow);
        classDetails = this.sharedService.classificationEnabled
          ? await this.processClassAndAttributes(rowCount, context, undefined, false, sheet)
          : new Map<string, ClassificationDescription>();
        console.log("CLASS MAPPING::", classDetails.values());
        const endRow = startRow + rowCount - 1;

        let result: any = [];
        for (const item of classDetails.values()) {
          result.push({
            classCode: item.classCode,
            attributes: item.attributes
          });
        }

        if (!this.sharedService.isExtentionProcess && this.sharedService.classificationEnabled && requestType as any != 'generateDescriptions') {
          const currentDescString = JSON.stringify(result);
          const currentDescHash = Md5.hashStr(currentDescString);
          if (currentDescHash !== this.sharedService.updatedDescHash) {
            await this.utilsService.displayDialog('Error', this.translationService.translate('Description generation Error'));
            this.sharedService.isProcessRunning = false;
            this.utilsService.displayStatus(null)
            await this.pushBufferIfNeeded(totalRows, [], pushCallback);
            return;
          }
        }

        const block = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);
        block.load('text');
        await context.sync();
        const values = block.text;

        this.groupRowsByRecordNumber(values, recordNumbers, incompleteRecords);

        const completed = await this.processCompletedRecords(
          incompleteRecords,
          recordLastIndexMap,
          endRow,
          requestType,
          pushBuffer,
          classDetails,
          false
        );
        this.deleteCompletedRecords(incompleteRecords, completed);
        this.deleteCompletedRecords(incompleteRecords, completed);

        await this.pushBufferIfNeeded(totalRows, pushBuffer, pushCallback);


        startRow += rowCount;
      }
      await this.processIncompleteRecords(
        incompleteRecords,
        requestType,
        pushBuffer,
        classDetails,
        false
      );

      await this.pushRemainingBuffer(totalRows, pushBuffer, pushCallback);

    });
  }

  /**
   * Build MDO records for a specific row range by batching Excel reads.
   * Rows are zero-based and inclusive. Header rows (0,1) are skipped automatically.
   */
  async buildRecordsForRowRange(
    startRow: number,
    endRow: number,
    requestType: any, context: Excel.RequestContext
  ): Promise<MdoRecord[]> {
    console.log(`buildRecordsForRowRange: startRow=${startRow}, endRow=${endRow}`);
    const results: MdoRecord[] = [];
    let effectiveStart = Math.max(2, Math.min(startRow, endRow));
    let effectiveEnd = Math.max(2, Math.max(startRow, endRow));

    let totalCols = 0;
    let totalRows = 0;
    let recordsBlockValues: string[][] = [];
    let changeRecordNumbers: any[] = [];
    let recordLastIndexMap: Map<string, number> = new Map();
    let recordBoundMap: Map<string, { first: number; last: number }> = new Map();


    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load(['columnCount', 'rowCount']);
    await context.sync();
    totalCols = usedRange.columnCount;
    totalRows = usedRange.rowCount;

    // Load full record number column once (C3..C{totalRows}) to compute last index map
    const recColStartRow = 2; // zero-based index for row 3
    const recColRowCount = Math.max(0, totalRows - recColStartRow);
    const recColRange = recColRowCount > 0 ? sheet.getRangeByIndexes(recColStartRow, 2, recColRowCount, 1) : null;
    if (recColRange) recColRange.load('values');
    await context.sync();

    const recordNumbers = recColRange ? (recColRange.values as any[]).map(r => r[0]) : [];


    const rowCount = effectiveEnd - effectiveStart + 1;
    const block = sheet.getRangeByIndexes(effectiveStart, 0, rowCount, totalCols);
    block.load('text');
    await context.sync();
    const values = block.text as string[][];


    const uniqueChangeRecordNumbers = Array.from(new Set(values.map(r => String(r[2] || '')).filter(recNo => recNo !== '')));
    // Use existing grouping util for the block, aligning record numbers slice with values
    recordLastIndexMap = this.getRecordLastIndexMap(recordNumbers as any);

    recordBoundMap = this.utilsService.getRecordIndexBoundsMap(recordNumbers);
    // Derive min first row and max last row among changed records using bounds map
    const candidateFirstIdxs: number[] = [];
    const candidateLastIdxs: number[] = [];
    for (const recNo of uniqueChangeRecordNumbers) {
      const b = recordBoundMap.get(recNo);
      if (b) {
        candidateFirstIdxs.push(b.first); // Excel row number
        candidateLastIdxs.push(b.last);   // Excel row number
      }
    }
    const minFirstRowExcel = candidateFirstIdxs.length ? Math.min(...candidateFirstIdxs) : (effectiveStart + 1); // convert effectiveStart (0-based) to Excel row
    const maxLastRowExcel = candidateLastIdxs.length ? Math.max(...candidateLastIdxs) : (effectiveEnd + 1);
    // Adjust effectiveStart (0-based) and effectiveEnd (Excel row) accordingly
    effectiveStart = Math.min(effectiveStart, Math.max(2, minFirstRowExcel - 1)); // keep header skip
    effectiveEnd = Math.max(effectiveEnd, maxLastRowExcel);
    const recordsBlockRange = sheet.getRangeByIndexes(effectiveStart, 0, effectiveEnd - effectiveStart, totalCols);
    recordsBlockRange.load('values');
    await context.sync();
    recordsBlockValues = recordsBlockRange.values as string[][];
    changeRecordNumbers = recordsBlockValues.map(r => String(r[2] || '')).filter(recNo => recNo !== '');


    const incompleteRecords = new Map<string, { rows: string[][]; rowIndexes: number[] }>();
    this.groupRowsByRecordNumber(recordsBlockValues as any, changeRecordNumbers as any, incompleteRecords, recordBoundMap);
    if (incompleteRecords.size === 0) return results;
    console.log("incompleted records size: ", incompleteRecords);

    // Derive class details for just this block
    let classDetails: Map<string, ClassificationDescription> = new Map();
    try {
      classDetails = this.sharedService.classificationEnabled
        ? await this.processClassAndAttributes(effectiveEnd - effectiveStart + 1, context, undefined, requestType == 'generateDescriptions' ? true : false)
        : new Map<string, ClassificationDescription>();
    } catch (e) {
      console.warn('buildRecordsForRowRange: processClassAndAttributes failed; continuing with empty classDetails', e);
    }

    // Use existing completion logic to build only records completed within the range
    const pushBuffer: MdoRecord[] = [];
    const completed = await this.processCompletedRecords(
      incompleteRecords,
      recordLastIndexMap,
      effectiveEnd,
      requestType,
      pushBuffer,
      classDetails,
      false
    );
    if (completed?.length) results.push(...pushBuffer);

    return results;
  }

  private async pushRemainingBuffer(
    totalRows: number,
    pushBuffer: MdoRecord[],
    pushCallback: (batchJson: {
      totalRecords: any;
      batchRecordCount: any;
      recordList: any;
    }) => Promise<void>
  ) {
    if (pushBuffer.length > 0) {
      const batchRecordCount = pushBuffer.length;
      await pushCallback({
        totalRecords: totalRows - 2,
        batchRecordCount,
        recordList: pushBuffer,
      });
    }
  }

  private async processCompletedRecords(
    incompleteRecords: Map<string, { rows: string[][]; rowIndexes: number[] }>,
    recordLastIndexMap: Map<string, number>,
    endRow: number,
    requestType: RequestType,
    pushBuffer: MdoRecord[],
    classDetails: Map<string, ClassificationDescription>,
    useIndividualRecords: boolean = false
  ): Promise<string[]> {
    const completed: string[] = [];
    for (const [recNo, recData] of incompleteRecords) {
      const lastIdx = recordLastIndexMap.get(recNo);
      if (lastIdx && lastIdx <= endRow) {
        if (useIndividualRecords) {
          const individualRecords = await this.buildIndividualRecordsFromRows(
            recNo,
            recData.rows,
            recData.rowIndexes,
            requestType,
            classDetails
          );
          pushBuffer.push(...individualRecords);
        } else {
          pushBuffer.push(
            await this.buildRecordJson(
              recNo,
              recData.rows,
              recData.rowIndexes,
              requestType,
              classDetails
            )
          );
        }
        completed.push(recNo);
      }
      else {
        console.log('record not completed yet:', recNo, 'lastIdx:', lastIdx, 'endRow:', endRow);
      }
    }
    return completed;
  }

  private deleteCompletedRecords(
    incompleteRecords: Map<string, { rows: string[][]; rowIndexes: number[] }>,
    completed: string[]
  ) {
    for (const recNo of completed) {
      incompleteRecords.delete(recNo);
    }
  }

  private async pushBufferIfNeeded(
    totalRows: number,
    pushBuffer: MdoRecord[],
    pushCallback: (batchJson: {
      totalRecords: any;
      batchRecordCount: any;
      recordList: any;
    }) => Promise<void>
  ) {
    if (pushBuffer.length >= this.RECORD_PUSH_BATCH) {
      const batchRecordCount = pushBuffer.length;
      // splice(0) removes all items from pushBuffer and returns them as an array
      // This is used to clear the buffer after pushing the batch
      await pushCallback({
        totalRecords: totalRows - 3,
        batchRecordCount,
        recordList: pushBuffer.splice(0),
      });
    }
  }

  private async processIncompleteRecords(
    incompleteRecords: Map<string, { rows: string[][]; rowIndexes: number[] }>,
    requestType: RequestType,
    pushBuffer: MdoRecord[],
    classDetails: Map<string, ClassificationDescription>,
    useIndividualRecords: boolean = false
  ) {
    if (incompleteRecords.size > 0) {
      for (const [recNo, recData] of incompleteRecords) {
        if (useIndividualRecords) {
          const individualRecords = await this.buildIndividualRecordsFromRows(
            recNo,
            recData.rows,
            recData.rowIndexes,
            requestType,
            classDetails
          );
          pushBuffer.push(...individualRecords);
        } else {
          pushBuffer.push(
            await this.buildRecordJson(
              recNo,
              recData.rows,
              recData.rowIndexes,
              requestType,
              classDetails
            )
          );
        }
      }
    }
  }

  private getRecordLastIndexMap(recordNumbers: string[]): Map<string, number> {
    const recordLastIndexMap = new Map<string, number>();
    for (const [idx, recNo] of recordNumbers.entries()) {
      if (recNo) recordLastIndexMap.set(recNo, idx + 3); // +3 since starts at row 4
    }
    return recordLastIndexMap;
  }

  private groupRowsByRecordNumber(
    values: string[][],
    recordNumbers: string[],
    incompleteRecords: Map<string, { rows: string[][]; rowIndexes: number[] }>,
    recordBoundMap?: Map<string, { first: number; last: number }>
  ) {

    // 1️⃣ Find records that must be skipped
    const recordsToSkip = new Set<string>();

    for (let i = 0; i < values.length; i++) {
      if (values[i]?.[0]?.toLowerCase() === 'sent for correction') {
        const recNo = recordNumbers[i];
        if (recNo) {
          recordsToSkip.add(recNo);
        }
      }
    }

    for (let idx = 0; idx < values.length; idx++) {
      const row = values[idx];
      const recNo = recordNumbers[idx];

      if (!recNo) continue;
      if (recordsToSkip.has(recNo)) continue;
      const recData = incompleteRecords.get(recNo) || {
        rows: [],
        rowIndexes: [],
      };
      let excelRow = recordBoundMap ? (recordBoundMap.get(recNo)?.first || 1) - 3 : idx;
      if (recData.rows.length > 0) {
        excelRow = recordBoundMap ? excelRow + recData.rows.length : idx;
      }
      recData.rows.push(row);

      recData.rowIndexes.push(excelRow);
      incompleteRecords.set(String(recNo), recData);
    }
  }

  /**
 * Reads and prepares a single MdoRecord for the record number of the selected cell in Excel.
 * Uses groupRowsByRecordNumber and avoids loading the entire used range in memory.
 * @returns Promise<MdoRecord | null>
 */
  async prepareSingleMdoRecordFromSelectedCell(context: Excel.RequestContext): Promise<MdoRecord | null> {
    console.log("Preparing single MdoRecord from selected cell...");

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(['rowIndex']);
    await context.sync();

    const selectedRowIndex = selectedRange.rowIndex;

    // Get the record number from column B (index 1) of the selected row
    const recordNumberRange = sheet.getRangeByIndexes(selectedRowIndex, 2, 1, 1);
    recordNumberRange.load('values');
    await context.sync();
    const recordNumber = String(recordNumberRange.values[0][0] ?? '');

    const classDetailsMap = await this.processClassAndAttributes(500, context, selectedRowIndex);

    if (!recordNumber) return null;

    // Get the used range dimensions
    const usedRange = sheet.getUsedRange();
    usedRange.load(['rowCount', 'columnCount']);
    await context.sync();

    const totalRows = usedRange.rowCount;
    const totalCols = usedRange.columnCount;

    // Read only the record number column for all rows
    const recordColRange = sheet.getRangeByIndexes(0, 2, totalRows, 1);
    recordColRange.load('values');
    await context.sync();
    const recordNumbers = recordColRange.values.map(r => String(r[0] ?? ''));

    // Find all row indexes with the selected record number
    const matchingRowIndexes: number[] = [];
    for (let i = 0; i < recordNumbers.length; i++) {
      if (recordNumbers[i] === recordNumber) {
        matchingRowIndexes.push(i);
      }
    }
    if (matchingRowIndexes.length === 0) return null;

    // Read only the matching rows
    const matchingRows: string[][] = [];
    for (const rowIdx of matchingRowIndexes) {
      const rowRange = sheet.getRangeByIndexes(rowIdx, 0, 1, totalCols);
      rowRange.load('text');
      await context.sync();
      matchingRows.push(rowRange.text[0]);
    }
    // Use groupRowsByRecordNumber for consistency
    const incompleteRecords = new Map<string, { rows: string[][]; rowIndexes: number[] }>();
    this.groupRowsByRecordNumber(matchingRows, Array(matchingRows.length).fill(recordNumber), incompleteRecords);

    // Build the MdoRecord using your existing logic
    const recData = incompleteRecords.get(recordNumber);
    if (!recData) return null;

    return await this.buildRecordJson(
      recordNumber,
      recData.rows,
      recData.rowIndexes,
      RequestType.Export, // or pass as parameter if needed
      classDetailsMap
    );

  }

  async processClassAndAttributes(
    batchSize: number,
    context: Excel.RequestContext,
    selectedRow?: number,
    isGenerateDescription: boolean = false,
    worksheet?: Excel.Worksheet
  ): Promise<Map<string, ClassificationDescription>> {
    if (!this.sharedService.classificationEnabled) {
      return new Map<string, ClassificationDescription>();
    }


    let classDetailsMap: (Map<string, ClassificationDescription>) = new Map<string, ClassificationDescription>();


    const sheet = worksheet || context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load(["rowCount", "columnCount"]);
    await context.sync();

    const totalRows = usedRange.rowCount;
    const totalCols = usedRange.columnCount;

    const classificationField = this.getConfiguredClassificationField();
    const classColumnIndex = classificationField
      ? this.sharedService.staticFieldsObject.find(field => field.fieldId === classificationField.fieldId)?.columnNumber
      : undefined;
    const classType = this.getConfiguredClassificationType(classificationField);
    if (classColumnIndex === undefined || classColumnIndex === -1) {
      console.error("Class not found");
      return classDetailsMap;
    }

    const startRowIndex = 2;
    let startRow = startRowIndex;

    if (selectedRow) {
      startRow = selectedRow;
    }

    while (startRow < totalRows) {
      const rowCount = Math.min(batchSize, totalRows - startRow);
      const block = sheet.getRangeByIndexes(startRow, 0, rowCount, totalCols);

      block.load("values");
      await context.sync();

      const rows = block.values;
      console.log("stored classAttributesMap", this.sharedService.classAttributesMap);
      const classCodeList = rows.map(r => ({ classCode: r[classColumnIndex], classType }))
                             .filter(obj => obj.classCode && !this.sharedService.findClassAttributes(obj.classCode, obj.classType));

      if (classCodeList?.length > 0) {
        const uniqueClassCodes = Array.from(new Map(classCodeList.map(item => [(item.classCode + item.classType), item])).values());
        console.log("Unique class codes with types:", uniqueClassCodes);
        if (uniqueClassCodes.length) {
          await this.sharedService.updateAttributeMap(uniqueClassCodes);
        }
      }
      for (const row of rows) {
        const classCode = row[classColumnIndex];
        const classDet = this.sharedService.findClassAttributes(classCode, classType);
        // Do not assume Record Number is always column C. Generated layouts can
        // preserve a template's system-column order, so resolve it from metadata.
        const recordNumberColumn = this.sharedService.staticFieldsObject
          .find(field => field.fieldId === 'mdo_number')?.columnNumber;
        const recordNumber = recordNumberColumn === undefined ? '' : row[recordNumberColumn];


        if (this.isRowEmpty(row) || row.length <= classColumnIndex || !classDet || !recordNumber) {
          continue;
        }
        const languages: string[] = this.sharedService.allLanguages || ['en'];
        const profileLang: string = this.sharedService.profileLanguage || (languages.length ? languages[0] : 'en');

        const attributesObj: any = {};
        for (const l of languages) {
          attributesObj[l] = {
            shortDesc: "",
            longDesc: "",
            attrs: {}
          };
        }

        const classObject: ClassificationDescription = {
          classCode: classDet.code,
          // `mod` is the class mode returned by Classification. It is part of
          // the class identity, so ABRASIVE and ABRASIVE--BRUSH remain distinct
          // in the Generate Description payload.
          classMode: String((classDet as any).mod ?? (classDet as any).classMode ?? ''),
          isDesc: false,
          classType: classDet.classType,
          classId: classDet.uuid,
          attributes: attributesObj
        };

        for (let col = classColumnIndex + 2; col < row.length; col += 2) {
          const attribute = row[col - 1];
          const value = row[col];
          if (!attribute) break;

          // Find all matching characteristics by description (case-insensitive)
          const matchingFields = classDet.charList.filter((f: any) => {
            let name = String(f.charDesc || '').trim();
            if (f.labels) {
              name = f.labels.find((label: any) => label.language == this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang)?.label || name;
            }
            return name.toLowerCase() === String(attribute).toLowerCase();
          });
          if (!matchingFields || matchingFields.length === 0) continue;

          // Prefer a matching field whose uuid is not yet present in attrs.
          let selectedField: any | null = null;
          for (const f of matchingFields) {
            if (!f || !f.uuid) continue;
            if (!classObject.attributes[profileLang].attrs[f.uuid]) {
              selectedField = f;
              break;
            }
          }
          // If all matching uuids are already present, pick the first occurrence (append value below)
          if (!selectedField) selectedField = matchingFields.find((f: any) => f && f.uuid) || null;
          if (!selectedField || !selectedField.uuid) continue;

          // Each attribute has only one value — set for profile language and blank for others
          if (!classObject.attributes[profileLang]) {
            classObject.attributes[profileLang] = { shortDesc: '', longDesc: '', attrs: {} };
          }
          classObject.attributes[profileLang].attrs[selectedField.uuid] = {
            fId: selectedField.charCode,
            vc: [{ c: value }]
          };

          // Ensure all other languages have an entry with blank value
          for (const langKey of Object.keys(classObject.attributes)) {
            if (langKey === profileLang) continue;
            if (!classObject.attributes[langKey]) {
              classObject.attributes[langKey] = { shortDesc: '', longDesc: '', attrs: {} };
            }
            if (!classObject.attributes[langKey].attrs[selectedField.uuid]) {
              classObject.attributes[langKey].attrs[selectedField.uuid] = {
                fId: selectedField.charCode,
                vc: [{ c: '' }]
              };
            }
          }
        }
        classDetailsMap.set(String(recordNumber), classObject);
      }
      startRow += rowCount;
    }

    return classDetailsMap;
  }

  /**
   * Classification type is optional form metadata. It must never be inferred
   * from a field ID belonging to a different generated element.
   */
  /**
   * The template Description/Class column is independent of a form's dynamic
   * class-control field.  When it is present, it is the only class that may
   * drive Generate Description; a dynamic class belongs to its own sheet.
   */
  private getConfiguredClassificationField(): Field | undefined {
    const visibleFieldIds = new Set(this.sharedService.staticFieldsObject.map(field => field.fieldId));
    const templateDescriptionField = this.sharedService.staticFieldsObject
      .find(field => field.fieldId === 'CLASSIFICATION_CLASS');
    if (templateDescriptionField) {
      return this.sharedService.fieldsMap.get('CLASSIFICATION_CLASS') || templateDescriptionField as Field;
    }

    // Workbooks without a Description/Class section retain the legacy support
    // for a real form class-control field.
    const formField = Array.from(this.sharedService.fieldsMap.values()).find((field: any) =>
      visibleFieldIds.has(field?.fieldId)
      && (field?.fieldType === 'CLASS_TYPE'
        || Boolean(field?.classCtrl)
        || Boolean(field?.metadata?.[0]?.classCtrl))
    ) as Field | undefined;
    if (formField) return formField;

    return undefined;
  }

  private getConfiguredClassificationType(classField = this.getConfiguredClassificationField()): string {
    const metadata = classField?.metadata?.[0] as any;
    return String(
      (classField as any)?.classTypeCtrl
      ?? (classField as any)?.classType
      ?? metadata?.classTypeId
      ?? metadata?.classType
      ?? ''
    );
  }

  /**
   * Builds the static-class portion of `mdoRecordES.descriptions` directly
   * from form metadata. The implementation deliberately has no flow-specific
   * IDs: every generated element reads only fields belonging to its form.
   */
  private buildStaticClassDescriptions(
    row: any[],
    existingDescriptions: ClassificationDescription[]
  ): ClassificationDescription[] {
    const staticFields = this.sharedService.staticFieldsObject
      .map(column => ({ column, field: this.sharedService.fieldsMap.get(column.fieldId) as any }))
      .filter(({ field }) => Boolean(field?.isStaticClassCharacteristics || field?.isStaticClassCharacteristicsUoM));

    if (!staticFields.length) return [];

    const language = this.sharedService.profileLanguage || this.sharedService.allLanguages?.[0] || 'en';
    const descriptionsByClass = new Map<string, ClassificationDescription>();

    for (const { column, field } of staticFields) {
      if (field?.isStaticClassCharacteristicsUoM) continue;

      const metadata = field.metadata?.[0] || {};
      const attributeMetadata = field.attributeMetadata || metadata.attributeMetadata || {};
      const classId = String(field.classId || metadata.classId || attributeMetadata.classId || '');
      if (!classId) continue;

      const classKey = classId;
      let description = descriptionsByClass.get(classKey);
      if (!description) {
        const existing = existingDescriptions.find(item => item?.classId === classId);
        description = existing
          ? JSON.parse(JSON.stringify(existing)) as ClassificationDescription
          : {
              classId,
              classCode: String(field.classCode || metadata.classCode || attributeMetadata.classCode || ''),
              classType: String(field.classType || metadata.classType || attributeMetadata.classType || ''),
              classMode: String(field.classMode || metadata.classMode || attributeMetadata.classMode || ''),
              isDesc: false,
              attributes: {}
            } as ClassificationDescription;
        description.attributes = description.attributes || {};
        description.attributes[language] = {
          ...(description.attributes[language] || {}),
          attrs: { ...(description.attributes[language]?.attrs || {}) }
        } as any;
        descriptionsByClass.set(classKey, description);
      }

      const value = row[column.columnNumber] ?? '';
      const uomColumn = staticFields.find(({ field: candidate }) =>
        candidate?.isStaticClassCharacteristicsUoM
        && (candidate.fieldId === `${field.fieldId}-UOM` || candidate.parentFieldId === field.fieldId)
      )?.column;
      const characteristicId = String(attributeMetadata.uuid || field.fieldId);
      const characteristicCode = String(attributeMetadata.charCode || field.charCode || field.fieldId);
      const attribute: any = {
        fId: characteristicCode,
        vc: [{ c: value }]
      };
      if (uomColumn) {
        attribute.uom = { vc: [{ c: row[uomColumn.columnNumber] ?? '' }] };
      }
      (description.attributes[language] as any).attrs[characteristicId] = attribute;
    }

    return Array.from(descriptionsByClass.values());
  }

  /**
   * Short/long description columns are template presentation fields rather
   * than layout fields. Their values must therefore be merged into the
   * selected/imported classification object instead of being emitted as HDVS
   * with synthetic field IDs. This preserves hidden class metadata on Change
   * flows while allowing an uploaded template to contain a final description
   * section with any visual styling.
   */
  private mergeTemplateDescriptionValues(
    row: any[],
    selectedDescription: ClassificationDescription | undefined,
    storedDescriptions: ClassificationDescription[]
  ): ClassificationDescription | undefined {
    const shortColumn = this.findDescriptionColumn('SHORT_DESCRIPTION', 'short description');
    const longColumn = this.findDescriptionColumn('LONG_DESCRIPTION', 'long description');
    if (!shortColumn && !longColumn) return selectedDescription;

    const base = selectedDescription
      || storedDescriptions.find(description => Boolean(description?.isDesc));
    if (!base) {
      // Description text without an associated class is not a valid platform
      // classification payload. Leave it untouched until the user selects a
      // class (or imports a record carrying the existing description object).
      return selectedDescription;
    }

    const language = this.sharedService.profileLanguage || this.sharedService.allLanguages?.[0] || 'en';
    const merged: any = JSON.parse(JSON.stringify(base));
    merged.isDesc = true;
    merged.attributes = merged.attributes || {};
    merged.attributes[language] = {
      ...(merged.attributes[language] || {}),
      attrs: { ...(merged.attributes[language]?.attrs || {}) }
    };
    if (shortColumn) merged.attributes[language].shortDesc = String(row[shortColumn.columnNumber] ?? '');
    if (longColumn) merged.attributes[language].longDesc = String(row[longColumn.columnNumber] ?? '');
    return merged as ClassificationDescription;
  }

  private findDescriptionColumn(fieldId: string, label: string): any | undefined {
    const normalizedLabel = label.replace(/\s/g, '').toLowerCase();
    return this.sharedService.staticFieldsObject.find(field =>
      field.fieldId === fieldId
      || String(field.label || '').replace(/\s/g, '').toLowerCase() === normalizedLabel
    );
  }

  isRowEmpty(row: any[]) {
    return !row || row.every(cell => cell === null || cell === undefined || cell === "");
  }

  async getSourceTargetMappingOfSelectedField(
    selectedRangeRowIndex: number,
    currentSelectedField: Field, context: Excel.RequestContext
  ): Promise<{ searchString: string; parent: Record<string, any> }> {
    const parentSet: { searchString: string; parent: Record<string, any> } = {
      searchString: "",
      parent: {},
    };

    try {
      if (!currentSelectedField || !currentSelectedField.metadata?.length) {
        console.warn("No currentSelectedField or metadata provided.");
        return parentSet;
      }



      const workbook = context.workbook;
      const activeSheet = workbook.worksheets.getActiveWorksheet();

      // load used range columnCount (single sync)
      const activeUsedRange = activeSheet.getUsedRange(true);
      activeUsedRange.load(["columnCount"]);
      await context.sync();

      const columnCount = activeUsedRange.columnCount ?? 0;


      // load the selected row range values
      const selectedRowRange = activeSheet.getRangeByIndexes(selectedRangeRowIndex, 0, 1, columnCount);
      selectedRowRange.load("values");
      await context.sync();

      const selectedRowValues = selectedRowRange.values?.[0] ?? [];

      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();

      const sheetExists = sheets.items.some(ws => ws.name === "DropDataDump");

      let dropDumpWorksheet: Excel.Worksheet;

      if (sheetExists) {
        dropDumpWorksheet = sheets.getItem("DropDataDump");
      } else {
        return parentSet;
      }
      const usedRange = dropDumpWorksheet.getUsedRange(true);
      usedRange.load(["rowCount"]);
      await context.sync();
      const fieldRange = dropDumpWorksheet.getRangeByIndexes(0, 0, usedRange.rowCount, 1);
      fieldRange.load("values");
      await context.sync();

      const fList: string[] = (fieldRange.values || []).map((r) => String(r[0]));
      const fieldIdToRowIndex = new Map<string, number>();
      for (let idx = 0; idx < fList.length; idx++) {
        fieldIdToRowIndex.set(fList[idx], idx);
      }

      const index = fieldIdToRowIndex.has(currentSelectedField.fieldId)
        ? fieldIdToRowIndex.get(currentSelectedField.fieldId)!
        : -1;

      let sourceFieldIdList: string[] = [];
      if (index >= 0) {
        const valueRange = dropDumpWorksheet.getRangeByIndexes(index, 1, 1, 1);
        valueRange.load("values");
        await context.sync();

        const raw = valueRange.values?.[0]?.[0];
        try {
          sourceFieldIdList = raw ? JSON.parse(raw) : [];
        } catch (err) {
          console.error("Failed to parse sourceFieldIdList JSON from DropDataDump:", err);
          sourceFieldIdList = [];
        }
      }

      if (!sourceFieldIdList?.length) {
        // nothing mapped
        return parentSet;
      }


      // iterate over each source field id
      for (const sourceFieldId of sourceFieldIdList) {
        const sourceField = this.sharedService.staticFieldsObject.find(f => f.fieldId == sourceFieldId);
        if (!sourceField) {
          // source field not in same layout - skip or log
          console.warn(`Source field ${sourceFieldId} not found in current layout fields`);
          continue;
        }

        let sourceFieldValue: any = undefined;
        // source is on same sheet: directly read from selectedRowValues using column number
        const parentFieldColumnNumber = this.sharedService.staticFieldsObject.find(m => m.fieldId == sourceFieldId)?.columnNumber;
        const selectedClassType = this.getConfiguredClassificationType();
        if (!parentFieldColumnNumber && parentFieldColumnNumber !== 0) {
          console.warn(`Source field ${sourceFieldId} column number not found in fieldData`);
          continue;
        }
        sourceFieldValue = selectedRowValues[parentFieldColumnNumber];
        try {
          const fieldmeta = this.sharedService.fieldsMap.get(sourceFieldId) || {} as Field;
          sourceFieldValue = await this.sharedService.getDropDownCodeOrText({ field: fieldmeta, fieldId: fieldmeta.fieldId, cellValue: sourceFieldValue, datsetId: fieldmeta.metadata?.[0]?.moduleId as string, formId: this.sharedService.data.formId });
          console.log(`Source Field ID: ${sourceFieldId}, Value:`, sourceFieldValue);
        }
        catch (er) {
          console.error(er);
        }
        // special case: CLASSIFICATION_CLASS - map to uuid
        if (sourceFieldId === "CLASSIFICATION_CLASS" && typeof sourceFieldValue === "string") {
          let classUuid = "";
          let classType = "";
          try {
            const GlobalVariables = JSON.parse(localStorage.getItem("GlobalVariables") || "{}");
            if (sourceFieldValue.includes("--")) {
              let classDetails = GlobalVariables.ClassDropData.find((m: any) => (m?.code + "--" + m?.mod) === sourceFieldValue && (!selectedClassType || selectedClassType === m?.classTypeId));
              classUuid = classDetails?.uuid;
              classType = classDetails?.classTypeId;

            } else {
              let classDetails = GlobalVariables.ClassDropData.find((m: any) => m?.code === sourceFieldValue && (!selectedClassType || selectedClassType === m?.classTypeId));
              classUuid = classDetails?.uuid;
              classType = classDetails?.classTypeId;
            }
          } catch (err) {
            console.warn("Error resolving CLASSIFICATION_CLASS uuid:", err);
          }
          if (!classUuid) {
            classType = classType ?? selectedClassType;
            const classKey = sourceFieldValue + classType;
            const classUniqueCode = { classCode: sourceFieldValue, classType: classType}
            if (!this.sharedService.classAttributesMap.has(classKey)) {
              await this.sharedService.updateAttributeMap([classUniqueCode]);
            }
            classUuid = this.sharedService.classAttributesMap.get(classKey)?.uuid || "";
          }
          sourceFieldValue = classUuid ?? sourceFieldValue;
        }
        if (typeof sourceFieldValue !== "undefined" && sourceFieldValue !== null) {
          parentSet.parent[sourceFieldId] = sourceFieldValue;
        }
      } // end for each sourceFieldId

    } catch (err) {
      console.error("getSourceTargetMappingOfSelectedField failed:", err);
    }

    return parentSet;
  }

  isSapNumberSequenceValid(sapNumbers: string[]): boolean {
    const seen = new Set<string>();
    let previous: string | null = null;

    for (const value of sapNumbers) {
      if (value === null || value === '') {
        continue;
      }

      const current = String(value).trim();

      // Same as previous, allowed (contiguous duplicate)
      if (current === previous) {
        continue;
      }

      // If seen before but not contiguous --> invalid
      if (seen.has(current)) {
        return false;
      }

      seen.add(current);
      previous = current;
    }

    return true;
  }


  async getSapNumbers(): Promise<number[]> {
    let sapNumbers: number[] = [];

    const resultSet = new Set<number>();
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      const startRow = 2;  // Row 3
      const colCIndex = 2; // Column C

      const usedRange = sheet.getUsedRange();
      usedRange.load('rowCount');
      await context.sync();

      const totalRows = usedRange.rowCount;
      const rowCount = Math.max(0, totalRows - startRow);

      if (rowCount === 0) {
        return [];
      }

      // Load both Column C and Column D together
      const range = sheet.getRangeByIndexes(
        startRow,
        colCIndex,
        rowCount,
        2
      );

      range.load('values');
      await context.sync();

      const allSapNumbers: string[] = range.values
        .map(row => row[0])

      range.values.forEach(row => {
        const colCValue = row[0];
        resultSet.add(colCValue);
      });
      return null
    })
    sapNumbers = Array.from(resultSet);

    return sapNumbers;
  }
}
