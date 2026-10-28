import { Injectable } from '@angular/core';
import { SharedService } from './shared.service';
import { ExcelReaderService } from './excel-reader.service';
import { DaxeXLBaseEvents, daxeXlEvent, Field, MdoRecord, MetaData, RequestType } from 'daxe-ui-sdk';
import { UtilsService } from './utils.service';
import { firstValueFrom, filter, take, timeout, catchError, of, throwError } from 'rxjs';
import { ExcelDataWriterService } from './excel-data-writer.service';
import { TranslationService } from './translation.service';

@Injectable({
  providedIn: 'root'
})
export class ExcelEventHandlerService {

  constructor(private readonly excelReader: ExcelReaderService,
    private readonly sharedService: SharedService,
    private readonly utilsService: UtilsService,
    private readonly excelWriter: ExcelDataWriterService,
    private readonly translationService: TranslationService) {
    this.handleSelectionChange = this.handleSelectionChange.bind(this);
    this.handleCellEditComplete = this.handleCellEditComplete.bind(this);
  }

  //#region handle excel cell selection change event
  private debounceTimeout: any;
  // Queue all change events to ensure each is processed
  private readonly changeEventQueue: Excel.WorksheetChangedEventArgs[] = [];
  private isFlushingAggregates = false;
  private isSelectionChangeRunning = false;
  private scheduleProcessingTimeout: any;
  // Track the last pushed change row index to avoid bouncing repeated edits
  private previousChangeRowIndex: number | null = null;
  async handleSelectionChange(event: Excel.WorksheetSelectionChangedEventArgs) {
    // if (this.isFlushingAggregates || this.isSelectionChangeRunning) {
    //   console.log('Currently flushing aggregates; ignoring selection change.');
    //   return;
    // }
    if (this.debounceTimeout) clearTimeout(this.debounceTimeout);
    this.debounceTimeout = setTimeout(async () => {
      try {
        this.isSelectionChangeRunning = true;
        await Excel.run({ delayForCellEdit: true }, async (context) => {
          const currentSheet = context.workbook.worksheets.getActiveWorksheet();
          currentSheet.load(['id', 'name']);
          let selectedRange = currentSheet.getRange(event.address);
          selectedRange.load(['rowIndex', 'rowCount', 'columnIndex', 'columnCount', 'values']);
          await context.sync();
          if (selectedRange.rowIndex < 2) {
            console.log("Header row selected, ignoring.");
            return;
          }
          if (selectedRange.rowCount > 1 || selectedRange.columnCount > 1) {
            console.log("Multiple rows/columns selected, ignoring.");
            return;
          }
          let selectedColumnRange = currentSheet.getCell(1, selectedRange.columnIndex);
          selectedColumnRange.load("values");
          await context.sync();
          let selectedFieldlabel = selectedColumnRange.values[0][0];
          const dynamicSheet = this.sharedService.dynamicClassificationSheet?.sheetName === currentSheet.name
            ? this.sharedService.dynamicClassificationSheet
            : undefined;
          const activeFields = dynamicSheet?.fields || this.sharedService.staticFieldsObject;
          let selectedFieldId = activeFields.find((f: any) => f.label === selectedFieldlabel)?.fieldId || "";
          console.log("Selected Field Id:", selectedFieldId);
          const selectedRow = selectedRange.rowIndex;
          const rowRange = currentSheet.getRangeByIndexes(selectedRow, 0, 1, dynamicSheet ? 6 : 3);
          rowRange.load("values");
          await context.sync();
          let rowValues = rowRange.values[0];
          let objNo = rowValues[2];
          // Description-class characteristics are inserted after class
          // selection, so their labels live in the data row rather than the
          // static header. Resolve the selected Value cell from the runtime
          // registry and expose it as a normal characteristic field.
          const runtimeAttribute = await this.resolveRuntimeClassificationAttribute(
            currentSheet,
            selectedRange,
            context
          );
          if (runtimeAttribute) {
            selectedFieldId = String(runtimeAttribute.uuid);
          }
          if (selectedFieldId) {
            let fieldObj = runtimeAttribute
              ? ({
                fieldId: String(runtimeAttribute.uuid),
                description: runtimeAttribute.charDesc || selectedFieldlabel || 'Value',
                fieldType: 'Static',
                isStaticClassCharacteristics: true,
                attributeMetadata: runtimeAttribute,
                metadata: [{
                  fieldId: String(runtimeAttribute.uuid),
                  description: runtimeAttribute.charDesc || selectedFieldlabel || 'Value',
                  pickList: runtimeAttribute.fieldType === 'DROPDOWN' ? '1' : '0',
                  structureId: 1,
                  isGridColumn: false,
                  moduleId: this.sharedService.data.controlData.moduleId
                }]
              } as Field)
              : this.sharedService.fieldsMap.get(selectedFieldId);
            if (dynamicSheet && selectedFieldId === 'CHARAC') {
              const classValue = String(rowValues[3] ?? '');
              fieldObj = {
                ...fieldObj,
                isDynamicClassificationCharacteristic: true,
                dynamicClassAttributes: this.sharedService.getDynamicClassAttributes(currentSheet.id, selectedRow),
                // This is only host routing context while the characteristic
                // is being selected; it is not a complete AttributeModel.
                attributeMetadata: { class: classValue } as any
              } as Field;
            }
            console.log("Selected Field Object:", fieldObj);
            const parentSet = dynamicSheet
              ? { parent: {}, child: {}, reference: {} }
              : await this.excelReader.getSourceTargetMappingOfSelectedField(selectedRow, fieldObj as Field, context);
            const isReadOnly = fieldObj?.isReadOnly === true;
            // The host performs the async class lookup. Keep the DAXE-owned
            // selected row in the request so its response cannot be written to
            // a stale host selection (or be discarded by the batch writer).
            let payload = { fieldId: selectedFieldId, fieldObj: isReadOnly ? null : fieldObj, businessRules: this.sharedService.data.businessRules.map(rule => rule.ruleId), stepId: this.sharedService.data.stepId, objectNumber: objNo, parentSet, formId: this.sharedService.data.formId, rowIndex: selectedRow };
            // A CLASS_TYPE control can legitimately have null layout metadata.
            // Do not let that prevent the dropdown event from reaching ngx-mdo.
            const fieldMetadata = Array.isArray(fieldObj?.metadata) ? fieldObj.metadata[0] : undefined;
            const isClassificationField = this.sharedService.classificationEnabled && (
              (fieldObj as any)?.isClassificationField === true
              || fieldObj?.fieldType === 'CLASS_TYPE'
            );
            if (isClassificationField) {
              // The class picker needs the selected record as context to load
              // the form's configured class type and available attributes.
              const record = dynamicSheet
                ? {} as MdoRecord
                : await this.excelReader.prepareSingleMdoRecordFromSelectedCell(context);
              // Layout descriptions can be rich-text HTML. ngx-mdo shows a
              // dynamic-class field's description above the picker, so prefer
              // the configured plain class name and never render `<p>…</p>`.
              const classField = {
                ...fieldObj,
                description: (fieldObj as any)?.classTypeCtrl?.className
                  || String((fieldObj as any)?.description || '').replace(/<[^>]*>/g, '').trim()
              };
              daxeXlEvent.next({
                type: DaxeXLBaseEvents.DAXEXLShowDropdownRequest,
                payload: { ...payload, fieldObj: isReadOnly ? null : classField, mdoRecord: record }
              });
            }
            else if (this.sharedService.classificationEnabled && selectedFieldId == 'CLASSIFICATION_CLASS') {
              console.log("Preparing MDO Record for Description Class Field");
              const record = await this.excelReader.prepareSingleMdoRecordFromSelectedCell(context);
              fieldObj = fieldObj || { fieldId: selectedFieldId, fieldType: "Static", metadata: [{ structureId: "1", moduleId: this.sharedService.data.controlData.moduleId, description: "Class", pickList: "1" } as MetaData], layoutId: this.sharedService.data.controlData.layoutId, } as Field;
              daxeXlEvent.next({
                type: DaxeXLBaseEvents.DAXEXLShowDropdownRequest, payload: { ...payload, mdoRecord: record, fieldObj }
              });
            }
            else if (fieldMetadata?.pickList == '30' || (fieldMetadata?.pickList == '1' && fieldMetadata?.isExternalLookup)) {
              const record = await this.excelReader.prepareSingleMdoRecordFromSelectedCell(context);
              daxeXlEvent.next({
                type: DaxeXLBaseEvents.DAXEXLShowDropdownRequest, payload: { ...payload, fieldObj: isReadOnly ? null : fieldObj, mdoRecord: record }
              });
            }
            else {
              daxeXlEvent.next({
                type: DaxeXLBaseEvents.DAXEXLShowDropdownRequest,
                payload: {
                  ...payload,
                  cellValue: String(selectedRange.values?.[0]?.[0] ?? ''),
                  fieldObj: isReadOnly ? null : fieldObj,
                  mdoRecord: {} as MdoRecord
                }
              });
            }
          }
        }).catch((err) => {
          console.error("Error in selection change handler:", err);
        });
      }
      finally {
        this.isSelectionChangeRunning = false;
      }
    }, 400);
  };

  async handleCellEditComplete(event: Excel.WorksheetChangedEventArgs) {
    // Collectively aggregate quick successive changes, then process once
    try {
      // Try to resolve the changed range indices so we can filter "bounce" events
      let rowIndex: number | null = null;
      let columnIndex: number | null = null;
      const classColIndex = this.sharedService.classificationEnabled
        ? this.getConfiguredClassificationField()?.columnNumber
        : undefined;
      try {
        await Excel.run({ delayForCellEdit: true }, async (ctx) => {
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();
          const eventSheet = ctx.workbook.worksheets.getItem(event.worksheetId);
          sheet.load('id');
          eventSheet.load('id');
          await ctx.sync();
          if (sheet.id !== eventSheet.id) {
            throw new Error("Change event on non-active sheet; ignoring for bounce check.")
          }
          const rng = sheet.getRange(event.address);
          rng.load(['rowIndex', 'columnIndex', 'rowCount', 'columnCount', 'address']);
          await ctx.sync();
          console.log("Changed range info:", rng.address, "Rows:", rng.rowCount, "Cols:", rng.columnCount);
          rowIndex = rng.rowIndex;
          columnIndex = rng.columnIndex;

          if (event.changeType === Excel.DataChangeType.rowInserted) {
            const address = event.address;

            const range = sheet.getRange(address);
            range.load('rowIndex');
            await ctx.sync();
            // Temporarily unprotect
            sheet.protection.unprotect();

            // Unlock inserted rows
            range.format.protection.locked = false;
            range.format.fill.clear();

            const protectedCols = [...this.sharedService.rowInsertReadonly, ...this.sharedService.headerColumns];
            for (let col of protectedCols) {
              const remainingRange = sheet.getRangeByIndexes(range.rowIndex, col, 1, 1);
              remainingRange.format.fill.color = '#d3d3d3';
              remainingRange.format.protection.locked = true;
            }

            // Re-protect
            sheet.protection.protect({
              allowAutoFilter: true,
              allowDeleteColumns: true,
              allowDeleteRows: true,
              allowEditObjects: true,
              allowEditScenarios: true,
              allowFormatCells: true,
              allowFormatColumns: true,
              allowFormatRows: true,
              allowInsertColumns: true,
              allowInsertHyperlinks: true,
              allowInsertRows: true,
              allowPivotTables: true,
              allowSort: true,
            });

            await ctx.sync();
          }
        });
      } catch (rangeErr) {
        // If we couldn't resolve the range, continue and enqueue as before
        console.warn('Could not read range info for change event; enqueueing without bounce check', rangeErr);
        return;
      }

      // If this change is in columns >= 37 and the row index matches the previous
      // change we processed, ignore it (bounce) to avoid duplicate handling.
      if (this.sharedService.classificationEnabled && columnIndex != null && classColIndex !== undefined && columnIndex >= classColIndex && rowIndex != null && this.previousChangeRowIndex === rowIndex) {
        console.log(`Bouncing duplicate change at row ${rowIndex}, col ${columnIndex}`);
        return; // do not enqueue
      }

      // Enqueue and remember row index
      this.changeEventQueue.push(event);
      this.previousChangeRowIndex = rowIndex;
      if (this.scheduleProcessingTimeout || this.isSelectionChangeRunning) {
        console.log('Processing already scheduled or selection change running; skipping new schedule.');
        clearTimeout(this.scheduleProcessingTimeout);
      }
      // A class is committed before Excel raises this event, so it does not
      // need the general one-second aggregation delay used for bulk edits.
      // This makes manually typed Description Class values populate promptly.
      const delayMs = this.sharedService.classificationEnabled && columnIndex === classColIndex ? 75 : 1000;
      this.scheduleProcessingTimeout = setTimeout(async () => {
        this.scheduleProcessingTimeout = null;
        this.sharedService.lock(async () => {
          await this.flushAggregatedChanges();
        })
      }, delayMs);
    } catch (err) {
      console.error('Failed to schedule aggregated flush; will attempt fallback:', err);
      // Fallback: if scheduling failed, try immediate flush to avoid missing changes
      try {
        this.sharedService.lock(async () => {
          await this.flushAggregatedChanges();
        })
      } catch (innerErr) {
        console.error('Immediate fallback flush failed:', innerErr);
      }
    }
  }

  /**
   * Runtime description-class attributes are written as label/value pairs in
   * the row, so their metadata cannot be recovered from the static header.
   * The in-memory map is only a performance cache and is empty after an add-in
   * refresh. Rebuild the selected value-cell binding from the persisted sheet
   * values and the class-characteristics API when needed.
   */
  private async resolveRuntimeClassificationAttribute(
    sheet: Excel.Worksheet,
    selectedRange: Excel.Range,
    context: Excel.RequestContext
  ): Promise<any | undefined> {
    const rowIndex = selectedRange.rowIndex;
    const columnIndex = selectedRange.columnIndex;
    const cached = this.sharedService.getRuntimeClassificationAttribute(sheet.id, rowIndex, columnIndex);
    if (cached) return cached;

    const classificationField: any = this.sharedService.staticFieldsObject
      .find(field => field.fieldId === 'CLASSIFICATION_CLASS');
    const classColumnIndex = classificationField?.columnNumber;
    // Generated attributes are [label, value] pairs immediately after Class.
    if (!Number.isFinite(classColumnIndex) || columnIndex <= classColumnIndex + 1 ||
      (columnIndex - classColumnIndex) % 2 !== 0) {
      return undefined;
    }

    const classCell = sheet.getCell(rowIndex, classColumnIndex);
    const attributeLabelCell = sheet.getCell(rowIndex, columnIndex - 1);
    classCell.load('values');
    attributeLabelCell.load('values');
    await context.sync();

    const classValue = String(classCell.values?.[0]?.[0] ?? '').trim();
    const attributeLabel = String(attributeLabelCell.values?.[0]?.[0] ?? '').trim();
    if (!classValue || !attributeLabel) return undefined;

    const configuredClassType = String(
      classificationField?.classType ?? classificationField?.classTypeId ??
      classificationField?.metadata?.[0]?.classType ?? ''
    );
    let classDetails: any = this.sharedService.findClassAttributes(classValue, configuredClassType);
    if (!classDetails) {
      // Preserve `code--mode` from the sheet; ngx-mdo separates it into the
      // class code and modSearchString when resolving the class.
      await this.sharedService.updateAttributeMap([{
        classCode: classValue,
        classType: configuredClassType
      }]);
      classDetails = this.sharedService.findClassAttributes(classValue, configuredClassType);
    }
    if (!classDetails?.charList) return undefined;

    const language = this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang;
    const normalise = (value: unknown) => String(value ?? '').trim().toLocaleLowerCase();
    const attribute = classDetails.charList.find((item: any) => {
      const translatedLabel = item.labels?.find((label: any) => label.language === language)?.label;
      return normalise(translatedLabel || item.charDesc) === normalise(attributeLabel);
    });
    if (attribute?.uuid) {
      this.sharedService.rememberRuntimeClassificationAttribute(sheet.id, rowIndex, columnIndex, attribute);
      return attribute;
    }
    return undefined;
  }

  private async flushAggregatedChanges() {
    if (this.isFlushingAggregates) {
      console.log('Already flushing aggregates, skipping redundant call.');
      return;
    }
    this.isFlushingAggregates = true;
    try {
      // Snapshot and clear current queue to allow new incoming events to accumulate
      const pending = this.changeEventQueue.splice(0, this.changeEventQueue.length);
      if (pending.length === 0) {
        console.log('No pending changes to flush.');
        return;
      }

      let aggregatedStartRow = Number.POSITIVE_INFINITY;
      let aggregatedEndRow = Number.NEGATIVE_INFINITY;
      let shouldProcess = false;
      // Track changed record numbers encountered in pending ranges
      const changedRecordNos = new Set<string>();
      const classColIndex = this.sharedService.classificationEnabled
        ? this.getConfiguredClassificationField()?.columnNumber
        : undefined;
      const recordNumberColIndex = this.sharedService.staticFieldsObject
        .find(field => field.fieldId === 'mdo_number')?.columnNumber ?? 2;
      // First pass inside Excel to compute the union range and do any inline formatting
      await Excel.run({ delayForCellEdit: true }, async (context) => {
        const currentSheet = context.workbook.worksheets.getActiveWorksheet();
        currentSheet.load('id, name');
        await context.sync();
        const usedRange = currentSheet.getUsedRange();
        usedRange.load('columnCount');
        const ranges = pending.map(p => p.worksheetId == currentSheet.id ? currentSheet.getRange(p.address) : null).filter(r => r !== null);
        for (const r of ranges) {
          r.load(['rowIndex', 'rowCount', 'columnIndex', 'columnCount', 'values']);
        }
        await context.sync();
        let classList: { rowIndex: number, value: string, type: string }[] = [];
        let restoreSheetProtection = false;
        let clearingClassAttributes = false;
        for (const r of ranges) {
          try {
            const isHeaderRow = r.rowIndex < 2;
            if (!isHeaderRow) {
              const isColumnC = r.columnIndex === 2;
              if (isColumnC) {
                // Lock generated header fields on continuation rows only.
                try {
                  const headerCols = this.sharedService.headerColumns;
                  // Read all record numbers in sheet to determine first occurrence rows
                  const usedRange = currentSheet.getUsedRange();
                  usedRange.load(['rowCount']);
                  await context.sync();
                  const totalRows_ = usedRange.rowCount;
                  const dataStartRow_ = 2; // zero-based index for Excel row 3
                  const count_ = Math.max(0, totalRows_ - dataStartRow_);
                  let boundsMap: Map<string, { first: number; last: number }> = new Map();
                  if (count_ > 0) {
                    const allRecRange = currentSheet.getRangeByIndexes(dataStartRow_, 2, count_, 1);
                    allRecRange.load('values');
                    await context.sync();
                    const allRecordNumbers: string[] = new Array(count_);
                    for (let idx = 0; idx < count_; idx++) {
                      allRecordNumbers[idx] = (allRecRange.values?.[idx]?.[0] ?? '').toString().trim();
                    }
                    boundsMap = this.utilsService.getRecordIndexBoundsMap(allRecordNumbers);
                  }
                  // For each row in changed span, lock header cells if it's not the first occurrence
                  for (let i = 0; i < r.rowCount; i++) {
                    const targetRow = r.rowIndex + i;
                    const isDataRow = targetRow >= 2;
                    if (!isDataRow) continue;
                    // Read current row's MDO Number directly
                    const currCell = currentSheet.getCell(targetRow, 2);
                    currCell.load('values');
                    await context.sync();
                    const currVal = (currCell.values?.[0]?.[0] ?? '').toString().trim();
                    if (currVal) changedRecordNos.add(currVal);
                    const bounds = boundsMap.get(currVal);
                    // bounds.first is 1-based Excel row; convert to zero-based
                    const firstIdxZero = bounds ? bounds.first - 1 : -1;
                    const isFirstOccurrence = firstIdxZero === targetRow;
                    for (const col of headerCols) {
                      const cell = currentSheet.getCell(targetRow, col);
                      if (isFirstOccurrence && !this.sharedService.isExtentionProcess) {
                        cell.format.fill.clear();
                        cell.format.protection.locked = false;
                      } else if (!this.sharedService.isExtentionProcess) {
                        cell.format.fill.color = '#d3d3d3';
                        cell.format.protection.locked = true;
                      }
                    }
                    for (const col of this.sharedService.alwaysReadOnly) {
                      const cell = currentSheet.getCell(targetRow, col);
                      cell.format.fill.color = '#d3d3d3';
                      cell.format.protection.locked = true;
                    }
                  }
                } catch (error_) {
                  console.warn('Failed to set protection for Header Data cells:', error_);
                }
              }

              if (this.sharedService.classificationEnabled && classColIndex !== undefined && r.columnIndex == classColIndex && r.columnCount == 1) {
                for (let i = 0; i < r.rowCount; i++) {
                  let cellValue = r.values[i][0];
                  if (cellValue) {
                    classList.push({ rowIndex: r.rowIndex + i, value: cellValue, type: this.getConfiguredClassificationType() });
                  }
                  else {
                    const rowIndex = r.rowIndex + i;
                    // Attribute labels are locked once written. Temporarily
                    // unprotect only while clearing a deleted Description
                    // Class; protection is restored below in this same batch.
                    if (!restoreSheetProtection) {
                      currentSheet.protection.unprotect();
                      restoreSheetProtection = true;
                    }
                    clearingClassAttributes = true;
                    await this.utilsService.displayStatus(this.translationService.translate('Removing class attributes...'));
                    const recordNumberCell = currentSheet.getCell(rowIndex, recordNumberColIndex);
                    recordNumberCell.load('values');
                    await context.sync();
                    const recordNumber = String(recordNumberCell.values?.[0]?.[0] ?? '').trim();
                    if (recordNumber) this.sharedService.clearDescriptionClass(recordNumber);
                    this.sharedService.clearRuntimeClassificationAttributes(currentSheet.id, rowIndex, classColIndex + 1);
                    const attributeColumnCount = usedRange.columnCount - (classColIndex + 1);
                    if (attributeColumnCount > 0) {
                      // Disable events in this same request context. Using a
                      // separate Excel.run leaves a gap where programmatic
                      // writes can be seen as user edits.
                      context.runtime.enableEvents = false;
                      await context.sync();
                      try {
                        const attributeRange = currentSheet.getRangeByIndexes(rowIndex, classColIndex + 1, 1, attributeColumnCount);
                        attributeRange.clear();
                        attributeRange.format.protection.locked = false;
                        await context.sync();
                      } finally {
                        context.runtime.enableEvents = true;
                        await context.sync();
                      }
                    }
                  }
                }
              }
            }
          } catch (e) {
            console.warn('Failed to color column F for row:', e);
          }

          // Only consider data rows (skip headers)
          const headerLastRow = 1;
          const startRow = Math.max(r.rowIndex, headerLastRow + 1);
          const endRow = r.rowIndex + r.rowCount - 1;
          if (endRow < startRow) continue;

          // Build triggering condition consistent with previous logic
          const lastChangedCol = r.columnIndex + r.columnCount - 1;
          if (this.sharedService.classificationEnabled && classColIndex !== undefined && lastChangedCol >= classColIndex) {
            shouldProcess = true;
            aggregatedStartRow = Math.min(aggregatedStartRow, startRow);
            aggregatedEndRow = Math.max(aggregatedEndRow, endRow);
          }
        }
        if (restoreSheetProtection) {
          currentSheet.protection.protect({
            allowAutoFilter: true, allowDeleteColumns: true, allowDeleteRows: true,
            allowEditObjects: true, allowEditScenarios: true, allowFormatCells: true,
            allowFormatColumns: true, allowFormatRows: true, allowInsertColumns: true,
            allowInsertHyperlinks: true, allowInsertRows: true, allowPivotTables: true,
            allowSort: true,
          });
        }
        if (this.sharedService.classificationEnabled && classList.length > 0) {
          await this.utilsService.displayStatus(this.translationService.translate('Loading class attributes...'));
          try {
            const classCodeList = classList
              .filter(c => c && !this.sharedService.findClassAttributes(c.value, c.type))
              .map(m => ({ classCode: m.value, classType: m.type }));
            if (classCodeList?.length > 0) {
              const uniqueClassCodes = Array.from(new Map( classCodeList.map(item => [(item.classCode + item.classType), item])).values());
              if (uniqueClassCodes.length) {
                await this.sharedService.updateAttributeMap(uniqueClassCodes);
              }
            }
            const attributeRows = classList.map(classObj => ({
              rowIndex: classObj.rowIndex,
              attrs: this.sharedService.findClassAttributes(classObj.value, classObj.type)?.charList
                ?.filter(m => m.status?.toString()?.toLowerCase() === 'true') || []
            }));
            // Keep event suppression and attribute writes in the same Excel
            // request context; no time/range-based event suppression is needed.
            context.runtime.enableEvents = false;
            await context.sync();
            try {
              await this.excelWriter.createAttributesInSheetBatch(attributeRows, context);
            } finally {
              context.runtime.enableEvents = true;
              await context.sync();
            }
          } finally {
            await this.utilsService.displayStatus(null);
          }
        }
        await context.sync();
        if (clearingClassAttributes) await this.utilsService.displayStatus(null);
        if (!shouldProcess || !Number.isFinite(aggregatedStartRow) || !Number.isFinite(aggregatedEndRow)) {
          console.log('No relevant changes detected for processing.');
          return;
        }

        //await this.generateDescription(aggregatedStartRow, aggregatedEndRow, context);
        await context.sync();
      });



    }
    catch (err) {
      console.error('Failed to flush aggregated changes:', err);
    }
    finally {
      // Clear the previous change tracker now that the flush has completed so
      // subsequent edits to the same row can be processed again.
      this.previousChangeRowIndex = null;
      this.isFlushingAggregates = false;
      // If new events arrived during flush, schedule another pass
      if (this.changeEventQueue.length > 0 && !this.scheduleProcessingTimeout) {
        this.scheduleProcessingTimeout = setTimeout(async () => {
          this.scheduleProcessingTimeout = null;
          await this.flushAggregatedChanges();
        }, 100);
      }
    }
  }


  /**
   * The Description/Class column must win over a dynamic class-control field.
   * Dynamic classification is handled on its separate sheet and must not make
   * the Description generator read a different column from this worksheet.
   */
  private getConfiguredClassificationField(): any {
    const descriptionClassField = this.sharedService.staticFieldsObject
      .find(field => field.fieldId === 'CLASSIFICATION_CLASS');
    if (descriptionClassField) return descriptionClassField;

    return this.sharedService.staticFieldsObject.find(field =>
      field.isClassificationField === true
      || field.fieldType === 'CLASS_TYPE'
    );
  }

  /** Uses classification type supplied by the active generated element. */
  private getConfiguredClassificationType(): string {
    const configuredField = this.getConfiguredClassificationField();
    const classField = configuredField
      ? this.sharedService.fieldsMap.get(configuredField.fieldId)
      : this.sharedService.fieldsMap.get('CLASSIFICATION_CLASS');
    const metadata = classField?.metadata?.[0] as any;
    const typeControl = classField?.classTypeCtrl ?? configuredField?.classTypeCtrl;
    return String(
      metadata?.classTypeId
      ?? metadata?.classType
      ?? typeControl?.classType
      ?? typeControl
      ?? classField?.classType
      ?? configuredField?.classType
      ?? ''
    );
  }

  async generateDescription(aggregatedStartRow: number, aggregatedEndRow: number, context: Excel.RequestContext) {
    if (!this.sharedService.classificationEnabled) {
      return;
    }

    try {
      await this.utilsService.displayStatus(this.translationService.translate('Generating descriptions...'));
      const recordList: MdoRecord[] = await this.excelReader.buildRecordsForRowRange(aggregatedStartRow, aggregatedEndRow, 'generateDescriptions', context);
      console.log('Prepared MDO records for aggregated set:', recordList.length, 'Rows:', aggregatedStartRow, aggregatedEndRow);
      if (recordList.length > 0) {
        const response = await this.sharedService.generateDescriptionsForRecords(recordList);
        await this.excelWriter.processGeneratedDescription(response);
      }
      await this.utilsService.displayStatus(null);
    } catch (error_) {
      console.error('Failed to prepare/emit MDO records for aggregated rows:', error_);
    } finally {
      await this.utilsService.displayStatus(null);
    }
  }

  async registerExcelEvents() {
    console.log("Registering Excel Events");
    await Excel.run({ delayForCellEdit: true }, async (ctx) => {
      ctx.workbook.worksheets.onSelectionChanged.add(this.handleSelectionChange);
      ctx.workbook.worksheets.onChanged.add(this.handleCellEditComplete);
      return ctx.sync();
    });
    // Safety net: lightweight ticker to ensure queued changes eventually flush
    try {
      const tickIntervalMs = 3000;
      setInterval(async () => {
        // If there are pending changes and no scheduled timeout, trigger a flush
        if (this.changeEventQueue.length > 0 && !this.scheduleProcessingTimeout && !this.isFlushingAggregates) {
          try {
            await this.flushAggregatedChanges();
          } catch (tickErr) {
            console.warn('Ticker flush attempt failed:', tickErr);
          }
        }
      }, tickIntervalMs);
    } catch (e) {
      console.warn('Failed to start safety ticker:', e);
    }
  }
}
