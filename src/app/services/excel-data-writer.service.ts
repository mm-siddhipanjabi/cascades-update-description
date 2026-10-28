import { Injectable } from '@angular/core';
import { Attributes, CharacteristicList, ClassificationDescription, ControlData, DaxeUIBaseEvent, DaxeUIBaseEvents, DaxeXLBaseEvents, FileDetailResponse, MdoRecord, MdoRecordES, RecordField, RequestType } from 'daxe-ui-sdk';
import { Field } from 'daxe-ui-sdk';
import { CellValueSaparator, DateTime, UtilsService } from './utils.service';
import { SharedService } from './shared.service';
import { TranslationService } from './translation.service';


/* global Excel, OfficeExtension */
declare const Excel: any;
declare const OfficeExtension: any;

interface FlattenedRowPlan {
  rowIndex: number;
  writeHeader: boolean;
  gridRowIndexesByGridId: Record<string, number>;
  hierarchyRowIndexesByStructureId: Record<string, number>;
}

@Injectable({
  providedIn: 'root'
})
export class ExcelDataWriterService {
  recordFirstRow: Map<string, number> = new Map<string, number>();

  constructor(private readonly utilsService: UtilsService, private readonly sharedService: SharedService, private readonly translationService: TranslationService) { }

  /**
   * Writes flattened data to Excel sheet with grids in parallel structure
   * @param recordsList List of MDO records to write
   * @param sheetName Target sheet name
   * @param allFields All field metadata
   * @param startRow Starting row number (default: 3)
   */
  async writeFlattenedDataToSheet(
    recordsList: MdoRecord[],
    sheetName: string,
    allFields: Field[],
    startRow: number = 2
  ): Promise<void> {
    const BATCH_SIZE = 5000;
    const currentSheetFields = allFields.filter(field => field.ColumnNumber !== -1);

    let rowsArray: Array<string>[] = [];
    let currentRow = startRow;
    let parentRow = currentRow;
    const attachmentList: any[] = [];
    let mdoDescObj: MdoRecordES[] = [];


    const classData = recordsList.map(record => {
      const description: any = record.mdoRecordES.descriptions?.find((d: { isDesc: boolean; }) => d.isDesc);
      const classCode = String(description?.classCode ?? '');
      const classMode = String(description?.classMode ?? '');
      return {
        classId: description?.classId,
        // Keep the modifier in the request/cache key. Class type is optional
        // for Description Class fields and cannot distinguish these classes.
        classCode: classMode && !classCode.includes('--') ? `${classCode}--${classMode}` : classCode,
        classType: description?.classType ?? ''
      };
    });
    const classCodeList = classData
      .filter(item => item.classCode)
      .map(item => ({ classCode: item.classCode, classType: item.classType }));
    console.log("Class code list:", classCodeList);
    await this.sharedService.updateAttributeMap(classCodeList);

    const classIdList = classData
      .filter(item => item.classId && !this.sharedService.findClassAttributes(item.classCode, item.classType))
      .map(item => item.classId as string);

    // Collect unique attachment IDs from all records (both hdvs and gvs)
    const attachmentIdSet = new Set<string>();
    for (const record of recordsList) {
      // Check header values (hdvs)
      const hdvs = record.mdoRecordES?.hdvs;
      if (hdvs) {
        for (const key of Object.keys(hdvs)) {
          if (this.sharedService.fieldsMap.get(key)?.metadata?.[0]?.pickList === "38") {
            const vcList = hdvs[key]?.vc;

            if (Array.isArray(vcList)) {
              for (const vcItem of vcList) {
                const attachmentId = vcItem?.c;

                if (attachmentId && typeof attachmentId === 'string') {
                  attachmentIdSet.add(attachmentId);
                }
              }
            }
          }

        }
      }

      // Check grid values (gvs)
      const gvs = record.mdoRecordES?.gvs;
      if (gvs) {
        for (const gridId of Object.keys(gvs)) {
          const grid = gvs[gridId];
          if (grid?.rows) {
            for (const row of grid.rows) {
              for (const fieldId of Object.keys(row)) {
                if (this.sharedService.fieldsMap.get(fieldId)?.metadata?.[0]?.pickList === "38") {
                  const vcList = row[fieldId]?.vc;

                  if (Array.isArray(vcList)) {
                    for (const vcItem of vcList) {
                      const attachmentId = vcItem?.c;

                      if (attachmentId && typeof attachmentId === 'string') {
                        attachmentIdSet.add(attachmentId);
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }

    // Fetch file details only if attachments exist
    if (attachmentIdSet.size > 0) {
      const attachmentIdList = Array.from(attachmentIdSet);
      const fileDetails = await this.sharedService.sendEventAndWaitForResponse<FileDetailResponse[]>(
        DaxeXLBaseEvents.DAXEXLAttachmentFileNamesRequest,
        DaxeXLBaseEvents.XFUSEAttachmentFileNamesResponse,
        attachmentIdList,
        undefined,
        10000,
        'Attachment file names request error'
      );

      // Update records with file names if response received
      if (fileDetails?.length) {
        const fileDetailsMap = new Map(fileDetails.map(file => [file.docId, file.fileDetails?.fileName]));

        for (const record of recordsList) {
          // Update header values (hdvs)
          const hdvs = record.mdoRecordES?.hdvs;
          if (hdvs) {
            for (const key of Object.keys(hdvs)) {
              if (this.sharedService.fieldsMap.get(key)?.metadata?.[0]?.pickList === "38") {
                const vcList = hdvs[key]?.vc;

                if (Array.isArray(vcList)) {
                  for (const vcItem of vcList) {
                    const attachmentId = vcItem?.c;

                    if (attachmentId && typeof attachmentId === 'string') {
                      const fileName = fileDetailsMap.get(attachmentId);

                      if (fileName) {
                        vcItem.t = fileName;
                      }
                    }
                  }
                }

              }
            }
          }

          // Update grid values (gvs)
          const gvs = record.mdoRecordES?.gvs;
          if (gvs) {
            for (const gridId of Object.keys(gvs)) {
              const grid = gvs[gridId];
              if (grid?.rows) {
                for (const row of grid.rows) {
                  for (const fieldId of Object.keys(row)) {
                    if (this.sharedService.fieldsMap.get(fieldId)?.metadata?.[0]?.pickList === "38") {
                      const vcList = row[fieldId]?.vc;
                      if (Array.isArray(vcList)) {
                        for (const vcItem of vcList) {
                          const attachmentId = vcItem?.c;

                          if (attachmentId && typeof attachmentId === 'string') {
                            const fileName = fileDetailsMap.get(attachmentId);

                            if (fileName) {
                              vcItem.t = fileName;
                            }
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      } else {
        console.warn('No file details received for attachments');
      }
    }

    if (classIdList.length > 0) {
      await this.sharedService.updateAttributeMapBasedOnClassId(classIdList);
    }

    // Process each record
    for (const record of recordsList) {
      mdoDescObj.push(record.mdoRecordES);
      record.mdoRecordES.id = record.controlData.recordNumber;
      this.recordFirstRow.set(record.mdoRecordES.id.toString(), parentRow);

      // Step 1: Find max grid rows for this record
      const maxGridRows = this.getMaxGridRows(record);
      const maxHierarchyRows = this.getMaxHierarchyRows(record);
      const childRows = record?.childRecord?.length ?? 0;
      const rowPlans = this.buildFlattenedRowPlans(record, childRows, maxGridRows, maxHierarchyRows);

      // Step 2: Build flattened rows for this record
      for (let rowIndex = 0; rowIndex < rowPlans.length; rowIndex++) {
        const rowPlan = rowPlans[rowIndex];
        const childIndex = rowIndex - 1;

        const columnArray = await this.buildFlattenedRow(
          record,
          rowPlan,
          currentSheetFields,
          allFields,
          attachmentList,
          childIndex
        );

        rowsArray.push(columnArray);

        // Step 3: Write batch when reaching BATCH_SIZE rows
        if (rowsArray.length >= BATCH_SIZE) {
          await this.processToWriteRecords(rowsArray, currentRow, sheetName, currentSheetFields, mdoDescObj);
          currentRow += rowsArray.length;
          rowsArray = [];
        }
        parentRow++;
      }
    }
    if (attachmentList.length) {
      let storedData = JSON.parse(localStorage.getItem('GlobalVariables') ?? "");
      let globalData;
      if (storedData && typeof storedData === 'string') {
        globalData = JSON.parse(storedData);
      }
      else {
        globalData = storedData
      }
      globalData.AttachmentList = attachmentList;
      localStorage.setItem('GlobalVariables', JSON.stringify(globalData));
    }

    // Step 4: Write remaining rows (< 5000)
    if (rowsArray.length > 0) {

      await this.processToWriteRecords(rowsArray, currentRow, sheetName, currentSheetFields, mdoDescObj);
      mdoDescObj = [];
    }
  }

  async writeExtentionFlattenedDataToSheet(
    recordsList: MdoRecord[],
    allFields: Field[],
  ): Promise<void> {
    const currentSheetFields = allFields.filter(field => field.ColumnNumber !== -1);
    // The existing-row import path performs a second Excel write after it has
    // identified the template rows. Await it so the next Excel.run is never
    // started while that discovery context is still active.
    await this.getExtendedDataCount(recordsList, currentSheetFields);
  }

  private async buildExtentionFlattenedRow(
    record: MdoRecord,
    currentSheetFields: Field[],
  ): Promise<Array<string>> {

    const columnArray: Array<string> = new Array(currentSheetFields.length);
    const hdvs = record.mdoRecordES.hdvs || {};

    const objectNumber = record.controlData?.recordNumber || record.mdoRecordES?.id || '';
    const status = record.mdoRecordES?.status || '';

    this.writeObjectNumberAndStatus(columnArray, currentSheetFields, objectNumber as string, status);
    await this.writeExtentionHeaderFields(columnArray, hdvs, currentSheetFields, objectNumber as string);

    return columnArray;
  }

  private async writeExtentionHeaderFields(
    columnArray: Array<string>,
    hdvs: any,
    currentSheetFields: Field[],
    objectNumber: string,
  ): Promise<void> {

    for (const fieldId of Object.keys(hdvs)) {
      if (!this.sharedService.extentionHeaderFields.includes(fieldId)) {
        continue;
      }

      const field = hdvs[fieldId];
      const fieldMeta = currentSheetFields.find(f => f.fieldId === fieldId);
      if (fieldMeta && !fieldMeta.metadata[0]?.isGridColumn) {
        const cellValue = await this.formatFieldValue(field, fieldMeta, objectNumber.toString(), []);
        const column = this.sharedService.staticFieldsObject.find(f => f.fieldId === fieldId)?.columnNumber;
        if (column !== undefined)
          columnArray[column] = cellValue;
      }
    }
  }

  async getExtendedDataCount(recordsList: MdoRecord[], currentSheetFields: Field[]) {
    let rowsArray: Array<string>[] = [];
    let firstValidIndex = -1;
    let boundsMap = new Map<string, { first: number; last: number }>();
    try {
      await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {

        const sheet: any = ctx.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(['rowCount', 'columnCount']);
        await ctx.sync();

        const totalRows = usedRange.rowCount;
        // A newly generated workbook contains only the two header rows. Do
        // not create a zero-height C:D range in that case: Excel for the web
        // rejects it with RichApi "argument is invalid".
        if (totalRows <= 2 || usedRange.columnCount < 4) return;
        const recRowStartRow = 2;
        const recColStartRow = 2;
        const recColRowCount = Math.max(0, totalRows - recRowStartRow);
        if (recColRowCount <= 0) return;
        const allRecRange = sheet.getRangeByIndexes(recRowStartRow, recColStartRow, recColRowCount, 2);
        let filterdRangeData = [];

        allRecRange.load('values');
        await ctx.sync();

        allRecRange.values.forEach((row: any, index: number) => {
          const colCValue = row[0];
          const colDValue = row[1];

          if (
            (colCValue === null || colCValue === '') &&
            colDValue !== null &&
            colDValue !== ''
          ) {
            if (firstValidIndex === -1) {
              firstValidIndex = index
            }
          }
        });

        if (firstValidIndex !== -1) {
          filterdRangeData = allRecRange.values.slice(firstValidIndex);
        }

        const allRecordNumbers = filterdRangeData
          ? (filterdRangeData as any[]).map(r => String(r?.[1] ?? '')).filter(Boolean)
          : [];
        rowsArray.length = allRecordNumbers.length;
        boundsMap = this.utilsService.getRecordIndexBoundsMap(allRecordNumbers);
      });

      // Do not call processToWriteRecords inside the Excel.run above. It opens
      // its own Excel.run for the bulk write; nesting those contexts produces
      // the RichApi "argument is invalid" error during record import.
      if (firstValidIndex === -1 || rowsArray.length === 0) return;
      for (const record of recordsList) {
        const recordNumber = record.controlData?.recordNumber || record.mdoRecordES?.id;

        const coundmapValues = boundsMap.get(recordNumber as string);
        if (!coundmapValues) {
          continue;
        }
        const columnArray = await this.buildExtentionFlattenedRow(record, currentSheetFields);
        for (let i = coundmapValues.first; i <= coundmapValues.last; i++) {
          rowsArray[i - 3] = columnArray;
        }
      }

      if (rowsArray.length > 0) {
        await this.processToWriteRecords(rowsArray, firstValidIndex + 2, '', currentSheetFields, []);
      }
    } catch (err) {
      console.error('writeExtentionFlattenedDataToSheet failed:', err);
    }
  }

  async processToWriteRecords(rowsArray: string[][], currentRow: number, sheetName: string, currentSheetFields: Field[], mdoDescObj: MdoRecordES[]) {

    const rowsChunk = rowsArray.slice(0);
    const batchStartRow = currentRow;
    await this.bulkWriteToSheet(sheetName, rowsChunk, batchStartRow, currentSheetFields.length);
    try {
      await this.applyAttributeFormatting(batchStartRow, rowsChunk);
    } catch (fmtErr) {
      console.warn('Attribute formatting for batch failed:', fmtErr);
    }
    await this.updateDescription(mdoDescObj);
    await this.applyCellFormatting();
  }

  async applyCellFormatting() {
    try {
      await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
        const sheet: any = ctx.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(['rowCount', 'columnCount']);
        await ctx.sync();

        const totalRows = Math.max(0, (usedRange.rowCount ?? 0) - 2);
        if (totalRows <= 0) return;

        // Do not assume MDO Number is column C. Templates can preserve a
        // different order for system columns, and an invalid range index makes
        // the entire post-import formatting pass fail in Excel for the web.
        const recordNumberColumn = this.sharedService.staticFieldsObject
          .find(field => field.fieldId === 'mdo_number')?.columnNumber;
        if (!Number.isInteger(recordNumberColumn) || recordNumberColumn < 0 || recordNumberColumn >= usedRange.columnCount) {
          console.warn('Skipping cell formatting: MDO Number column is not present on the active sheet.');
          return;
        }
        const allRecRange = sheet.getRangeByIndexes(2, recordNumberColumn, totalRows, 1);
        allRecRange.load('values');
        await ctx.sync();
        const allRecordNumbers: string[] = allRecRange.values.map((r: any) => (r && r[0]) ? r[0].toString() : '');
        const boundsMap = this.utilsService.getRecordIndexBoundsMap(allRecordNumbers);

        // A generated/runtime field can have no physical column on this sheet.
        // Filter before creating Office ranges; undefined/out-of-bounds indexes
        // otherwise surface as RichApi "argument is invalid" errors.
        const headerCols = Array.from(new Set([
          ...this.sharedService.headerColumns,
          ...this.sharedService.readOnlyColumnsOnChange,
          ...this.sharedService.alwaysReadOnly,
          ...this.sharedService.readonlyFieldCols
        ].filter((column): column is number =>
          Number.isInteger(column) && column >= 0 && column < usedRange.columnCount
        )));
        if (!headerCols.length) return;

        // Process each record bounds. Note: bounds.first/last are 1-based Excel rows.
        for (const [recordNumber, bounds] of boundsMap) {
          if (!bounds || typeof bounds.first !== 'number' || typeof bounds.last !== 'number') continue;
          if (bounds.last < bounds.first) {
            console.log(`bounds:-${bounds}, recordnumber:- ${recordNumber}`);
            continue;
          }
          // Convert to zero-based row indices for Excel API
          const startRowZero = bounds.first - 1;
          const endRowZero = bounds.last - 1;
          if (startRowZero < 2 || startRowZero >= usedRange.rowCount) continue;
          const rowCount = Math.min(endRowZero, usedRange.rowCount - 1) - startRowZero + 1;
          if (rowCount <= 0) continue;

          for (const col of headerCols) {
            try {

              this.fillColorAndLockCells(sheet, startRowZero, col, rowCount);

            } catch (colErr) {
              // Non-fatal per-column failure; continue with other columns
              console.warn(`Failed formatting column ${col} for record ${recordNumber}:`, colErr);
            }
          }
        }
        // No explicit sync here — Office JS will sync at end of run
      });
    } catch (err) {
      console.error('applyCellFormatting failed:', err);
    }

  }

  fillColorAndLockCells(sheet: Excel.Worksheet, startRowZero: number, col: number, rowCount: number) {
    // First occurrence: clear + unlock
    const firstCellRange = sheet.getRangeByIndexes(startRowZero, col, 1, 1);
    firstCellRange.format.fill.clear();
    firstCellRange.format.protection.locked = false;
    if (this.sharedService.readOnlyColumnsOnChange.includes(col) || this.sharedService.alwaysReadOnly.includes(col) || this.sharedService.readonlyFieldCols.includes(col)) {
      firstCellRange.format.fill.color = '#d3d3d3';
      firstCellRange.format.protection.locked = true;
    }

    if (rowCount > 1 && (!this.sharedService.isExtentionProcess || this.sharedService.readonlyFieldCols.includes(col))) {
      // Remaining rows: gray + locked (batch range)
      const remainingRange = sheet.getRangeByIndexes(startRowZero + 1, col, rowCount - 1, 1);
      remainingRange.format.fill.color = '#d3d3d3';
      remainingRange.format.protection.locked = true;
    }
  }

  /**
   * Gets the maximum number of grid rows across all grids in a record
   * @param record MDO record
   * @returns Maximum grid row count (minimum 1 for header-only records)
   */
  private getMaxGridRows(record: MdoRecord): number {
    let maxGridRows = 0;
    const recordNumber = record.controlData?.recordNumber || record.mdoRecordES?.id?.toString() || '';

    if (record.mdoRecordES.gvs) {
      for (const gridId in record.mdoRecordES.gvs) {
        if (Object.prototype.hasOwnProperty.call(record.mdoRecordES.gvs, gridId)) {
          const grid = record.mdoRecordES.gvs[gridId];
          if (grid?.rows) {
            // Filter to only count rows that belong to THIS record
            // const matchingRows = grid.rows.filter(row =>
            //   row['ObjectNumber']?.vc?.[0]?.c === recordNumber ||
            //   row['ObjectNumber']?.vc?.[0]?.c === recordNumber
            // );
            maxGridRows = Math.max(maxGridRows, grid.rows.length);
          }
        }
      }
    }

    // If no grids, still create 1 row for header data
    return maxGridRows === 0 ? 1 : maxGridRows;
  }

  private getMaxHierarchyRows(record: MdoRecord): number {
    let maxHierarchyRows = 0;

    if (record.mdoRecordES.hyvs) {
      for (const structureId in record.mdoRecordES.hyvs) {
        if (Object.prototype.hasOwnProperty.call(record.mdoRecordES.hyvs, structureId)) {
          const hierarchy = record.mdoRecordES.hyvs[structureId];
          if (hierarchy?.rows) {
            maxHierarchyRows = Math.max(maxHierarchyRows, hierarchy.rows.length);
          }
        }
      }
    }

    return maxHierarchyRows === 0 ? 1 : maxHierarchyRows;
  }

  private buildFlattenedRowPlans(
    record: MdoRecord,
    childRows: number,
    maxGridRows: number,
    maxHierarchyRows: number
  ): FlattenedRowPlan[] {
    const gvs = record.mdoRecordES.gvs || {};
    const hierarchyGridIds = Object.keys(gvs).filter(gridId => this.isHierarchyOwnedGrid(gridId));

    if (!this.sharedService.isExtentionProcess || hierarchyGridIds.length === 0) {
      const totalRows = this.sharedService.isExtentionProcess
        ? Math.max(maxGridRows, maxHierarchyRows, childRows + 1)
        : Math.max(maxGridRows, childRows + 1);

      return Array.from({ length: totalRows }, (_, rowIndex) => ({
        rowIndex,
        writeHeader: rowIndex === 0,
        gridRowIndexesByGridId: this.getGridRowIndexesForRow(gvs, Object.keys(gvs), rowIndex),
        hierarchyRowIndexesByStructureId: this.sharedService.isExtentionProcess && rowIndex < maxHierarchyRows
          ? this.getHierarchyRowIndexesForRow(record.mdoRecordES.hyvs || {}, rowIndex)
          : {},
      }));
    }

    const hyvs = record.mdoRecordES.hyvs || {};
    const headerGridIds = Object.keys(gvs).filter(gridId => !this.isHierarchyOwnedGrid(gridId));
    const headerGridRows = this.getMaxGridRowsForGridIds(gvs, headerGridIds);
    const hierarchyPlans = this.buildHierarchyGridRowPlans(record, hierarchyGridIds);

    if (hierarchyPlans.length === 0) {
      const totalRows = Math.max(maxGridRows, maxHierarchyRows, childRows + 1);
      return Array.from({ length: totalRows }, (_, rowIndex) => ({
        rowIndex,
        writeHeader: rowIndex === 0,
        gridRowIndexesByGridId: this.getGridRowIndexesForRow(gvs, Object.keys(gvs), rowIndex),
        hierarchyRowIndexesByStructureId: rowIndex < maxHierarchyRows
          ? this.getHierarchyRowIndexesForRow(hyvs, rowIndex)
          : {},
      }));
    }

    const totalRows = Math.max(hierarchyPlans.length, headerGridRows, childRows + 1, 1);
    return Array.from({ length: totalRows }, (_, rowIndex) => {
      const hierarchyPlan = hierarchyPlans[rowIndex];
      return {
        rowIndex,
        writeHeader: rowIndex === 0,
        gridRowIndexesByGridId: {
          ...this.getGridRowIndexesForRow(gvs, headerGridIds, rowIndex),
          ...(hierarchyPlan?.gridRowIndexesByGridId || {}),
        },
        hierarchyRowIndexesByStructureId: hierarchyPlan?.hierarchyRowIndexesByStructureId || {},
      };
    });
  }

  private buildHierarchyGridRowPlans(record: MdoRecord, hierarchyGridIds: string[]): FlattenedRowPlan[] {
    const plans: FlattenedRowPlan[] = [];
    const hyvs = record.mdoRecordES.hyvs || {};
    const gvs = record.mdoRecordES.gvs || {};
    const maxHierarchyRows = this.getActualMaxHierarchyRows(record);
    const usedGridRowIndexesByGridId: Record<string, Set<number>> = {};

    for (let hierarchyRowIndex = 0; hierarchyRowIndex < maxHierarchyRows; hierarchyRowIndex++) {
      const hierarchyRowIndexesByStructureId = this.getHierarchyRowIndexesForRow(hyvs, hierarchyRowIndex);
      const gridRowIndexesByGridId: Record<string, number[]> = {};
      let rowCount = 1;

      for (const gridId of hierarchyGridIds) {
        const grid = gvs[gridId];
        const parentStructureId = this.getGridParentStructureId(gridId);
        const hierarchyRow = hyvs[parentStructureId]?.rows?.[hierarchyRowIndex];
        if (!grid?.rows?.length || !hierarchyRow) {
          continue;
        }

        const keyFields = this.getHierarchyKeyFields(parentStructureId, hyvs[parentStructureId]);
        const matchingRows = this.getMatchingGridRowIndexes(grid.rows, hierarchyRow, keyFields);
        if (matchingRows.length > 0) {
          gridRowIndexesByGridId[gridId] = matchingRows;
          if (!usedGridRowIndexesByGridId[gridId]) {
            usedGridRowIndexesByGridId[gridId] = new Set<number>();
          }
          matchingRows.forEach(rowIndex => usedGridRowIndexesByGridId[gridId].add(rowIndex));
          rowCount = Math.max(rowCount, matchingRows.length);
        }
      }

      for (let gridRowOffset = 0; gridRowOffset < rowCount; gridRowOffset++) {
        const gridRowsForPlan: Record<string, number> = {};
        for (const gridId of Object.keys(gridRowIndexesByGridId)) {
          const gridRowIndex = gridRowIndexesByGridId[gridId][gridRowOffset];
          if (gridRowIndex !== undefined) {
            gridRowsForPlan[gridId] = gridRowIndex;
          }
        }

        plans.push({
          rowIndex: plans.length,
          writeHeader: plans.length === 0,
          gridRowIndexesByGridId: gridRowsForPlan,
          hierarchyRowIndexesByStructureId: gridRowOffset === 0 ? hierarchyRowIndexesByStructureId : {},
        });
      }
    }

    this.appendUnmatchedHierarchyGridRows(plans, gvs, hierarchyGridIds, usedGridRowIndexesByGridId);
    return plans;
  }

  private appendUnmatchedHierarchyGridRows(
    plans: FlattenedRowPlan[],
    gvs: any,
    hierarchyGridIds: string[],
    usedGridRowIndexesByGridId: Record<string, Set<number>>
  ): void {
    const unmatchedGridRowsByGridId: Record<string, number[]> = {};
    let rowCount = 0;

    for (const gridId of hierarchyGridIds) {
      const gridRows = gvs[gridId]?.rows || [];
      const usedGridRows = usedGridRowIndexesByGridId[gridId] || new Set<number>();
      const unmatchedRows = gridRows
        .map((_row: any, rowIndex: number) => rowIndex)
        .filter((rowIndex: number) => !usedGridRows.has(rowIndex));

      if (unmatchedRows.length > 0) {
        unmatchedGridRowsByGridId[gridId] = unmatchedRows;
        rowCount = Math.max(rowCount, unmatchedRows.length);
      }
    }

    for (let rowOffset = 0; rowOffset < rowCount; rowOffset++) {
      const gridRowsForPlan: Record<string, number> = {};
      for (const gridId of Object.keys(unmatchedGridRowsByGridId)) {
        const gridRowIndex = unmatchedGridRowsByGridId[gridId][rowOffset];
        if (gridRowIndex !== undefined) {
          gridRowsForPlan[gridId] = gridRowIndex;
        }
      }

      plans.push({
        rowIndex: plans.length,
        writeHeader: plans.length === 0,
        gridRowIndexesByGridId: gridRowsForPlan,
        hierarchyRowIndexesByStructureId: {},
      });
    }
  }

  private isHierarchyOwnedGrid(gridId: string): boolean {
    return this.getGridParentStructureId(gridId) !== 1;
  }

  private getGridParentStructureId(gridId: string): number {
    const structureId = Number(this.sharedService.fieldsMap.get(gridId)?.metadata?.[0]?.structureId);
    return Number.isFinite(structureId) && structureId > 0 ? structureId : 1;
  }

  private getActualMaxHierarchyRows(record: MdoRecord): number {
    let maxHierarchyRows = 0;
    const hyvs = record.mdoRecordES.hyvs || {};
    for (const structureId of Object.keys(hyvs)) {
      maxHierarchyRows = Math.max(maxHierarchyRows, hyvs[structureId]?.rows?.length || 0);
    }
    return maxHierarchyRows;
  }

  private getMaxGridRowsForGridIds(gvs: any, gridIds: string[]): number {
    return gridIds.reduce((maxRows, gridId) => Math.max(maxRows, gvs[gridId]?.rows?.length || 0), 0);
  }

  private getGridRowIndexesForRow(gvs: any, gridIds: string[], rowIndex: number): Record<string, number> {
    const gridRowIndexesByGridId: Record<string, number> = {};
    for (const gridId of gridIds) {
      if (gvs[gridId]?.rows?.[rowIndex]) {
        gridRowIndexesByGridId[gridId] = rowIndex;
      }
    }
    return gridRowIndexesByGridId;
  }

  private getHierarchyRowIndexesForRow(hyvs: any, rowIndex: number): Record<string, number> {
    const hierarchyRowIndexesByStructureId: Record<string, number> = {};
    for (const structureId of Object.keys(hyvs)) {
      if (hyvs[structureId]?.rows?.[rowIndex]) {
        hierarchyRowIndexesByStructureId[structureId] = rowIndex;
      }
    }
    return hierarchyRowIndexesByStructureId;
  }

  private getHierarchyKeyFields(structureId: number, hierarchy: any): string[] {
    if (hierarchy?.keyFields?.length) {
      return hierarchy.keyFields;
    }

    return Array.from(this.sharedService.fieldsMap.values())
      .filter((field: any) => Number(field?.metadata?.[0]?.structureId) === structureId && field?.metadata?.[0]?.isKeyField)
      .map((field: any) => field.fieldId);
  }

  private getMatchingGridRowIndexes(
    gridRows: any[],
    hierarchyRow: any,
    keyFields: string[]
  ): number[] {
    const matchingKeyFields = this.getMatchingKeyFields(hierarchyRow, gridRows, keyFields);
    const hierarchyKey = this.getRowKey(hierarchyRow, matchingKeyFields);

    if (!hierarchyKey) {
      return [];
    }

    return gridRows.reduce((rows: number[], gridRow, gridRowIndex) => {
      if (this.getRowKey(gridRow, matchingKeyFields) === hierarchyKey) {
        rows.push(gridRowIndex);
      }
      return rows;
    }, []);
  }

  private getMatchingKeyFields(hierarchyRow: any, gridRows: any[], keyFields: string[]): string[] {
    if (hierarchyRow?.['ROW_UUID'] && gridRows.some(gridRow => gridRow?.['ROW_UUID'])) {
      return ['ROW_UUID'];
    }

    const availableKeyFields = keyFields.filter(fieldId =>
      hierarchyRow?.[fieldId] &&
      gridRows.some(gridRow => gridRow?.[fieldId] && this.getRecordFieldValue(gridRow[fieldId]) !== '')
    );

    return availableKeyFields.length > 0
      ? availableKeyFields
      : this.getCommonHierarchyGridFields(hierarchyRow, gridRows);
  }

  private getCommonHierarchyGridFields(hierarchyRow: any, gridRows: any[]): string[] {
    const ignoredFields = new Set(['ROW_UUID', 'UUID']);
    const hierarchyFieldIds = Object.keys(hierarchyRow || {}).filter(fieldId => !ignoredFields.has(fieldId));

    return hierarchyFieldIds.filter(fieldId =>
      gridRows.some(gridRow => gridRow?.[fieldId] && this.getRecordFieldValue(gridRow[fieldId]) !== '')
    );
  }

  private getRowKey(row: any, keyFields: string[]): string {
    if (!row || keyFields.length === 0) {
      return '';
    }

    const keyParts = keyFields.map(fieldId => this.getRecordFieldValue(row[fieldId]));
    return keyParts.every(value => value !== '') ? keyParts.join('|') : '';
  }

  private getRecordFieldValue(field: RecordField): string {
    if (!field?.vc?.length) {
      return '';
    }
    return field.vc.map(value => value?.c ?? '').join(',');
  }

  /**
   * Builds a single flattened row combining header and grid data
   * @param record MDO record
   * @param rowIndex Current row index within the record's data
   * @param currentSheetFields Fields for current sheet
   * @param allFields All fields metadata
   * @param attachmentList Array to collect attachments
   * @returns Array of cell values for the row
   */
  private async buildFlattenedRow(
    record: MdoRecord,
    rowPlan: FlattenedRowPlan,
    currentSheetFields: Field[],
    allFields: Field[],
    attachmentList: any[],
    childRowIndex: number
  ): Promise<Array<string>> {
    const columnArray: Array<string> = new Array(currentSheetFields.length);
    const hdvs = record.mdoRecordES.hdvs || {};
    const gvs = record.mdoRecordES.gvs || {};

    const objectNumber = record.controlData?.recordNumber || record.mdoRecordES?.id || '';
    const status = record.mdoRecordES?.status || '';

    if (rowPlan.writeHeader) {
      this.writeObjectNumberAndStatus(columnArray, currentSheetFields, objectNumber as string, status);
      await this.writeHeaderFields(columnArray, hdvs, currentSheetFields, objectNumber as string, attachmentList);
      if (record?.mdoRecordES?.descriptions?.length > 0) {
        const attrsByLang = record.mdoRecordES.descriptions?.find((desc: { isDesc: boolean; }) => desc.isDesc)?.attributes || {};
        const langs = Object.keys(attrsByLang);
        let chosenLang = langs.find(l => l.toLowerCase() === this.sharedService.profileLanguage) || langs[0];
        if (chosenLang) {
          const attrObj: any = attrsByLang[chosenLang];
          const classDet = record.mdoRecordES.descriptions?.find((desc: { isDesc: boolean; }) => desc.isDesc) ?? null;
          let classCode = "", classMode = "", classUUID = "", classType = "";
          if(classDet){
            classCode = String(classDet.classCode ?? '');
            classMode = String((classDet as any).classMode ?? '');
            classUUID = String(classDet.classId ?? '');
            classType = String(classDet.classType ?? '');
          }
          const classDisplayValue = classMode && !classCode.includes('--') ? `${classCode}--${classMode}` : classCode;
          const cachedClass: any = this.sharedService.findClassAttributes(classDisplayValue, classType)
            || this.sharedService.classAttributesMap.get(classUUID);
          const classCodeByLanguage = String(cachedClass?.code ?? classCode);
          const classModeByLanguage = String(cachedClass?.mod ?? cachedClass?.classMode ?? classMode);
          const classDisplayByLanguage = classModeByLanguage
            ? `${classCodeByLanguage}--${classModeByLanguage}`
            : classCodeByLanguage;
          console.log("Class display value:", classDisplayByLanguage);
          const attrColumns = this.buildAttrColumns(attrObj, classDisplayByLanguage);
          const classColumnIndex = this.sharedService.staticFieldsObject.find(fld => fld.fieldId === 'CLASSIFICATION_CLASS')?.columnNumber;
          if (classColumnIndex !== undefined) {
            columnArray[classColumnIndex] = classDisplayByLanguage;
            let attrCol = classColumnIndex + 1;
            for (let i = 0; i < attrColumns.length; i++) {
              columnArray[attrCol] = attrColumns[i];
              attrCol++;
            }
          }
        }
      }
    }

    if (Object.keys(rowPlan.gridRowIndexesByGridId).length > 0) {
      await this.writeGridFields(
        columnArray,
        gvs,
        currentSheetFields,
        objectNumber as string,
        rowPlan.gridRowIndexesByGridId,
        attachmentList,
        status
      );
    }

    if (Object.keys(rowPlan.hierarchyRowIndexesByStructureId).length > 0 && this.sharedService.isExtentionProcess) {
      await this.writeHierarchyFields(
        columnArray,
        record.mdoRecordES.hyvs || {},
        currentSheetFields,
        objectNumber as string,
        rowPlan.hierarchyRowIndexesByStructureId,
        attachmentList,
        status
      );
    }

    if (childRowIndex >= 0 && record?.childRecord && record?.childRecord[childRowIndex]) {
      const childHdvs = record.childRecord[childRowIndex].mdoRecordES.hdvs || {};

      const parentRecordNumber = record.controlData.recordNumber;
      const status = record.childRecord[childRowIndex].mdoRecordES?.status || '';

      this.writeObjectNumberAndStatus(columnArray, currentSheetFields, parentRecordNumber as string, status);

      await this.writeHeaderFields(columnArray, childHdvs, currentSheetFields, parentRecordNumber as string, []);
    }
    return columnArray;
  }

  private writeObjectNumberAndStatus(
    columnArray: Array<string>,
    currentSheetFields: Field[],
    objectNumber: string,
    status: string
  ): void {
    const objNumField = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'mdo_number');
    const statusField = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'status');

    if (objNumField) {
      columnArray[objNumField.columnNumber] = objectNumber.toString();
    }
    if (statusField) {
      columnArray[statusField.columnNumber] = status;
    }
  }

  private async writeHeaderFields(
    columnArray: Array<string>,
    hdvs: any,
    currentSheetFields: Field[],
    objectNumber: string,
    attachmentList: any[]
  ): Promise<void> {
    for (const fieldId of Object.keys(hdvs)) {
      const field = hdvs[fieldId];
      const fieldMeta = currentSheetFields.find(f => f.fieldId === fieldId);
      if (fieldMeta && !fieldMeta.metadata[0]?.isGridColumn) {
        const cellValue = await this.formatFieldValue(field, fieldMeta, objectNumber.toString(), attachmentList);
        const column = this.sharedService.staticFieldsObject.find(f => f.fieldId === fieldId)?.columnNumber;
        if (column !== undefined)
          columnArray[column] = cellValue;
      }
    }
  }

  private async writeGridFields(
    columnArray: Array<string>,
    gvs: any,
    currentSheetFields: Field[],
    objectNumber: string,
    gridRowIndexesByGridId: Record<string, number>,
    attachmentList: any[],
    status: string
  ): Promise<void> {
    for (const gridId of Object.keys(gridRowIndexesByGridId)) {
      const grid = gvs[gridId];
      if (grid && grid.rows) {
        // const currentRecordGridRows = grid.rows.filter((row: any) =>
        //   row['ObjectNumber']?.vc?.[0]?.c === objectNumber.toString() ||
        //   row['ObjectNumber']?.vc?.[0]?.c === objectNumber.toString()
        // );
        const rowIndex = gridRowIndexesByGridId[gridId];
        const gridRow = grid.rows[rowIndex];
        if (gridRow) {
          this.writeObjectNumberAndStatus(columnArray, currentSheetFields, objectNumber as string, status);
          for (const fieldId of Object.keys(gridRow)) {
            const field = gridRow[fieldId];
            const fieldMeta = currentSheetFields.find(f => f.fieldId === fieldId);
            if (fieldMeta?.metadata?.[0]?.isGridColumn) {
              const cellValue = await this.formatFieldValue(field, fieldMeta, objectNumber.toString(), attachmentList);
              const column = this.sharedService.staticFieldsObject.find(f => f.fieldId === fieldId)?.columnNumber;
              if (column !== undefined) {
                columnArray[column] = cellValue;
              }
            }
          }
        }
      }
    }
  }

  private async writeHierarchyFields(
    columnArray: Array<string>,
    hyvs: any,
    currentSheetFields: Field[],
    objectNumber: string,
    hierarchyRowIndexesByStructureId: Record<string, number>,
    attachmentList: any[],
    status: string
  ): Promise<void> {
    for (const structureId of Object.keys(hierarchyRowIndexesByStructureId)) {
      const hierarchy = hyvs[structureId];
      if (hierarchy && hierarchy.rows) {
        const rowIndex = hierarchyRowIndexesByStructureId[structureId];
        const hierarchyRow = hierarchy.rows[rowIndex];
        if (hierarchyRow) {
          this.writeObjectNumberAndStatus(columnArray, currentSheetFields, objectNumber as string, status);
          for (const fieldId of Object.keys(hierarchyRow)) {
            const field = hierarchyRow[fieldId];
            const fieldMeta = currentSheetFields.find(f => f.fieldId === fieldId);
            if (fieldMeta) {
              const cellValue = await this.formatFieldValue(field, fieldMeta, objectNumber.toString(), attachmentList);
              const column = this.sharedService.staticFieldsObject.find(f => f.fieldId === fieldId)?.columnNumber;
              if (column !== undefined) {
                columnArray[column] = cellValue;
              }
            }
          }
        }
      }
    }
  }

  /**
   * Formats field value based on field type (reused logic from importDataHelper)
   * @param field Record field
   * @param fieldMeta Field metadata
   * @param objectNumber Object number for attachment tracking
   * @param attachmentList Array to collect attachments
   * @returns Formatted cell value as string
   */
  private async formatFieldValue(
    field: RecordField,
    fieldMeta: Field,
    objectNumber: string,
    attachmentList: any[]
  ): Promise<string> {
    if (!field || !field.vc) {
      return '';
    }
    let cellValue = field.vc.map(m => m.c).join(',') || '';

    // Date field (pickList 52)
    if (fieldMeta?.metadata[0]?.pickList === "52" || fieldMeta?.attributeMetadata?.fieldType === "DATE") {
      const dateValue = this.utilsService.getDateFromMilliseconds(cellValue, DateTime.DATE);
      cellValue = dateValue;
    }

    // DateTime field (pickList 53)
    if (fieldMeta?.metadata[0]?.pickList === "53") {
      const dateValue = this.utilsService.getDateTimeFromMillies(Number(cellValue));
      cellValue = dateValue;
    }

    // Attachment field (pickList 38)
    if (fieldMeta?.metadata[0]?.pickList === "38") {
      const valueCollection = field.vc;
      attachmentList.push(...valueCollection.map(m => ({
        fileName: m.t ? m.t : m.c,
        id: m.c,
        fieldId: field.fId,
        objectNumber
      })));
      cellValue = valueCollection?.map(m => m.t ? m.t : m.c)?.join(CellValueSaparator.attachementSaparator) || '';
    }

    // Transient fields (clear during copy)
    // Note: This would require GlobalVariables.IsCopyReference check if needed
    // if (GlobalVariables.IsCopyReference && fieldMeta?.metadata[0]?.isTransient) {
    //   cellValue = "";
    // }
    const ifExistsInExcelForm = this.sharedService.staticFieldsObject.some(item => item.fieldId === fieldMeta.fieldId);
    if (["1", "4", "18", "37", "30"].includes(fieldMeta.metadata[0].pickList || '') && !this.sharedService.codeFields.includes(fieldMeta.fieldId) && ifExistsInExcelForm && cellValue) {
      cellValue = await this.sharedService.getDropDownCodeOrText({
        field: fieldMeta,
        fieldId: fieldMeta.fieldId,
        cellValue: cellValue,
        datsetId: fieldMeta.metadata?.[0]?.moduleId as string,
        formId: this.sharedService.data.formId
      }, true);
    }

    return cellValue;
  }

  /**
   * Writes batch of rows to Excel sheet (reused pattern from importDataHelper)
   * @param sheetName Target sheet name
   * @param rowsArray Array of rows to write
   * @param startRow Starting row index
   * @param columnCount Total number of columns
   */
  private async bulkWriteToSheet(
    sheetName: string,
    rowsArray: Array<string>[],
    startRow: number,
    columnCount: number
  ): Promise<void> {
    await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
      const sheet: any = ctx.workbook.worksheets.getActiveWorksheet();

      // Compute actual column count from max array length in rowsArray (handles ragged arrays)
      const actualColumnCount = Math.max(columnCount, ...rowsArray.map(row => row?.length ?? 0));

      // Pad shorter rows with blank strings to match max column count
      const paddedRows = rowsArray.map(row => {
        const paddedRow = [...(row || [])];
        while (paddedRow.length < actualColumnCount) {
          paddedRow.push('');
        }
        return paddedRow;
      });


      // Get the range for batch write
      const dataRange = sheet.getRangeByIndexes(
        startRow,
        0,
        paddedRows.length,
        actualColumnCount
      );

      dataRange.load('address');
      await ctx.sync();

      console.log(`Writing to range ${dataRange.address} with ${paddedRows.length} rows and ${actualColumnCount} columns.`);

      // Format all cells as text
      dataRange.numberFormat = [['@']];


      console.log("decimal fields", this.sharedService.decimalFields);
      for (let decimalFieldId of this.sharedService.decimalFields) {
        if (decimalFieldId) {

          try {
            const decimalField = this.sharedService.fieldsMap.get(decimalFieldId) as any;
            console.log("decimalfield", decimalField);
            const fieldMeta = this.sharedService.staticFieldsObject.find(field => field.fieldId == decimalFieldId);
            if (!fieldMeta) continue;
            const decimalFieldCol = fieldMeta.column
            if (decimalFieldCol) {
              const format = "0." + "0".repeat(Number(decimalField?.metadata?.[0]?.decimalValue));
              sheet.getRange(`${decimalFieldCol}:${decimalFieldCol}`).numberFormat = [[format]];
            }
          } catch (nfErr) {
            console.warn(`Failed to set decimal numberFormat for column ${decimalFieldId}:`, nfErr);
          }
        }
      }
      dataRange.format.columnWidth = 150;
      dataRange.load('values');
      await ctx.sync();
      // Write all data at once (bulk operation)
      dataRange.values = paddedRows;

      await ctx.sync();
    }).catch((error: any) => {
      console.error('Error writing to Excel sheet:', error);
      if (error instanceof OfficeExtension.Error) {
        console.error('Debug info:', JSON.stringify(error.debugInfo));
      }
      throw error;
    });
  }

  /**
   * Gets the current used row count in a sheet
   * @param sheetName Sheet name
   * @returns Promise with row count
   */
  async getUsedRows(sheetName: string): Promise<number> {
    return await Excel.run({ delayForCellEdit: true }, async (ctx: any) => {
      const sheet: any = ctx.workbook.worksheets.getItem(sheetName);
      const usedRange = sheet.getUsedRange(true);
      usedRange.load('rowCount');
      await ctx.sync();
      return usedRange.rowCount;
    });
  }

  async parseAndwriteApplyDataRuleResponse(response: MdoRecord[], sheetName?: string) {
    try {
      console.log("Apply Data Rules Response to write in sheet:", response);
      if (!response || response.length == 0) return;
      const allRows: any[] = [];

      for (const record of response) {
        const { hdvs, gvs, hyvs } = record.mdoRecordES;
        const { arbitraryRecordNo } = record.controlData;
        const recordNumber = { vc: [{ c: arbitraryRecordNo }] };
        const ROW_UUID = { vc: [{ c: record.controlData?.excelRow || record.mdoRecordES?.id }] };
        // 1. hdvs (wrap hdvs as a row)
        const hdvsRow: any = { ...hdvs, recordNumber, ROW_UUID };
        allRows.push(hdvsRow);

        // 2. gvs rows
        // 2. gvs rows
        for (const grid of Object.values(gvs || {})) {
          for (const row of grid.rows || []) {
            allRows.push({ ...row, recordNumber });
          }
        }

        // 3. hyvs rows
        for (const hier of Object.values(hyvs || {})) {
          for (const row of hier.rows || []) {
            allRows.push({ ...row, recordNumber });
          }
        }
      };

      await this.writeDefaultsInSheet(allRows, RequestType.Defaults, sheetName);
    } catch (error) {
      console.error('Error parsing and writing Apply Data Rule response:', error);
    }
  }
  async writeDefaultsInSheet(
    recordList: { [FieldId: string]: RecordField }[],
    requestType: RequestType,
    sheetName?: string
  ) {
    console.log("Writing Defaults/Data Rules response in sheet:", recordList);
    await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {

      const workSheet: Excel.Worksheet = sheetName
        ? ctx.workbook.worksheets.getItem(sheetName)
        : ctx.workbook.worksheets.getActiveWorksheet();
      const currentSheetFields = this.sharedService.staticFieldsObject as unknown as Field[];
      const startRow = 2;
      let usedRange = workSheet.getUsedRange();
      usedRange.load('rowCount, columnCount');
      await ctx.sync();
      const TotalRows = usedRange.rowCount;

      const objectNumberColValues = await this.getObjectNumberColValues(workSheet, TotalRows, ctx);

      const { changeLogExistingArray, changeLogNewArray } = this.processRecordList(recordList, currentSheetFields);

      await this.updateExistingRows(workSheet, currentSheetFields, changeLogExistingArray, startRow, TotalRows, ctx);

      await this.insertNewRows(workSheet, currentSheetFields, changeLogNewArray, objectNumberColValues, ctx);
    });
  }

  private async getObjectNumberColValues(workSheet: Excel.Worksheet, TotalRows: number, ctx: any): Promise<string[]> {
    const newColRange = workSheet.getRangeByIndexes(0, 2, TotalRows, 1);
    newColRange.load('text');
    await ctx.sync();
    return newColRange.text.map((r: any) => r[0]);
  }

  private processRecordList(recordList: any[], currentSheetFields: Field[]) {
    let changeLogExistingArray: excelChangeLog[] = [];
    let changeLogNewArray: excelChangeLog[] = [];

    for (let record of recordList) {
      this.processRecordFields(
        record,
        currentSheetFields,
        changeLogExistingArray,
        changeLogNewArray
      );
    }

    return { changeLogExistingArray, changeLogNewArray };
  }

  /**
   * After rows are written to the sheet, apply attribute formatting similar to `createAttributesInSheet`.
   * Scans each written row for attribute name/value pairs and applies bold, wrap, width and mandatory color.
   */
  private async applyAttributeFormatting(startRow: number, rowsChunk: Array<string>[]) {
    if (!this.sharedService.classificationEnabled) return;

    console.log('Applying attribute formatting for rows starting at', startRow, 'with', rowsChunk.length, 'rows');
    if (!rowsChunk || rowsChunk.length === 0) return;
    try {
      await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
        const sheet: any = ctx.workbook.worksheets.getActiveWorksheet();
        const classField = this.sharedService.staticFieldsObject.find((f: any) => f.fieldId === 'CLASSIFICATION_CLASS');
        const classColIndex = classField?.columnNumber;
        console.log('Applying attribute formatting for rows starting at', startRow, 'classColIndex=', classColIndex);
        for (let r = 0; r < rowsChunk.length; r++) {
          const row = rowsChunk[r] || [];
          const excelRow = startRow + r;
          const classCode = (classColIndex != null && row[classColIndex]) ? String(row[classColIndex]).trim() : '';
          console.log(`Row ${excelRow}: classCode='${classCode}'`);
          // Only format attribute columns when a classCode exists and we have metadata for it
          if (!classCode) continue;
          const classType = this.getConfiguredClassificationType();
          const classDet = this.sharedService.findClassAttributes(classCode, classType);
          console.log(`Row ${excelRow}: classDet=`, classDet);
          if (!classDet || !Array.isArray(classDet.charList) || classDet.charList.length === 0) continue;

          // Build a lookup map of attribute names (lowercased) -> attribute metadata
          const attrMap = new Map<string, any>();
          for (const ch of classDet.charList) {
            let name = String(ch.charDesc || '').trim();
            if (ch.labels) {
              name = ch.labels.find((label: any) => label.language == this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang)?.label || name;
            }
            const nameLower = name.toLowerCase();
            if (nameLower) attrMap.set(nameLower, ch);
          }
          console.log(`Row ${excelRow}: attrMap keys=`, Array.from(attrMap.keys()));
          // Attributes are expected to be placed after the CLASSIFICATION_CLASS column
          // They are stored as name,value pairs in adjacent columns. Iterate starting
          // from the column after `classColIndex` and step by 2 to handle pairs.
          if (classColIndex == null) continue;
          const startAttrCol = classColIndex + 1;
          for (let col = startAttrCol; col < row.length - 1; col += 2) {
            const rawName = row[col];
            const rawValue = row[col + 1];
            if (rawName === undefined || rawName === null || String(rawName).trim() === '') continue;
            if (rawValue === undefined) continue;
            const key = String(rawName).trim().toLowerCase();
            const match = attrMap.get(key);
            if (!match) continue; // not an attribute name for this class

            try {
              const nameRange = sheet.getRangeByIndexes(excelRow, col, 1, 1);
              nameRange.format.font.bold = true;
              nameRange.format.wrapText = true;
              nameRange.format.horizontalAlignment = 'Center';
              nameRange.format.verticalAlignment = 'Bottom';
              nameRange.format.columnWidth = 150;
              nameRange.format.fill.color = "#a6c9ec";
              nameRange.format.protection.locked = true;

              const valueRange = sheet.getRangeByIndexes(excelRow, col + 1, 1, 1);
              valueRange.format.wrapText = true;
              valueRange.format.verticalAlignment = 'Bottom';
              valueRange.format.columnWidth = 150;

              if (match && (match.isManatory === true || String(match.isManatory).toLowerCase() === 'true')) {
                nameRange.format.font.color = '#ff0000';
              }
            } catch (cellErr) {
              console.warn(`Failed formatting attribute cell at row ${excelRow} col ${col}:`, cellErr);
            }
          }
        }

        await ctx.sync();
      });
    } catch (err) {
      console.error('applyAttributeFormatting failed:', err);
    }
  }
  private processRecordFields(
    record: any,
    currentSheetFields: Field[],
    changeLogExistingArray: excelChangeLog[],
    changeLogNewArray: excelChangeLog[]
  ) {
    let fieldIdList = Object.keys(record);
    let sheetRowId = Number(record["ROW_UUID"]?.vc[0]?.c);
    let recordNumber = record["recordNumber"]?.vc[0]?.c;

    for (let fieldId of fieldIdList) {
      if (fieldId === "ROW_UUID") continue;

      let field = currentSheetFields.find(m => m.fieldId == fieldId && m.metadata);
      let isNewRecord = fieldId == 'ObjectNumber' && !sheetRowId;
      let recordField = record[fieldId];

      if (recordField?.vc && (recordField?.isCorrected || isNewRecord)) {
        const changeLogArray = sheetRowId ? changeLogExistingArray : changeLogNewArray;
        let columnDataRow = this.buildColumnDataRow(field, recordField, sheetRowId, recordNumber);

        let fieldColumn = changeLogArray.find(m => m.columnId == fieldId);
        if (fieldColumn) {
          fieldColumn.columnRows.push(columnDataRow);
        } else {
          let changeLog: excelChangeLog = new excelChangeLog();
          changeLog.columnId = fieldId;
          changeLog.columnRows.push(columnDataRow);
          changeLogArray.push(changeLog);
        }
      }
    }
  }

  private buildColumnDataRow(
    field: Field | undefined,
    recordField: any,
    sheetRowId: number,
    recordNumber: string
  ): columnDataRow {
    let columnDataRow = {} as columnDataRow;
    columnDataRow.rowid = sheetRowId;
    columnDataRow.Value = recordField?.vc?.[0]?.c;

    if (field?.metadata[0]?.pickList == "52") {
      let dateValue = this.utilsService.getDateFromMilliseconds(recordField?.vc?.[0]?.c, DateTime.DATE);
      columnDataRow.Value = dateValue;
    }
    if (field?.metadata[0]?.pickList == "53") {
      let dateValue = this.utilsService.getDateTimeFromMillies(Number(recordField?.vc?.[0]?.c));
      columnDataRow.Value = dateValue;
    }
    columnDataRow.recordNumber = recordNumber;
    return columnDataRow;
  }

  private async updateExistingRows(
    workSheet: Excel.Worksheet,
    currentSheetFields: any[],
    changeLogExistingArray: excelChangeLog[],
    startRow: number,
    TotalRows: number,
    ctx: any
  ) {
    console.log("Updating existing rows with changes:", changeLogExistingArray);
    if (changeLogExistingArray.length === 0) return;
    for (let changeLog of changeLogExistingArray) {
      let field = currentSheetFields.find(m => m.fieldId == changeLog.columnId);
      if (!field) {
        console.log("Field not availabel in sheet:-", changeLog.columnId);
        continue;
      }
      let existingRows = changeLog.columnRows.filter(m => m.rowid);
      if (existingRows.length > 0) {
        let fieldColumnRange = workSheet.getRangeByIndexes(startRow, field.columnNumber, TotalRows - startRow, 1);
        fieldColumnRange.load('values');
        fieldColumnRange.numberFormat = [["@"]];
        await ctx.sync();
        let fieldColumnRangeValues: string[][] = fieldColumnRange.values;
        for (let row of existingRows) {
          fieldColumnRangeValues[row.rowid - 1] = ["'" + row.Value];
        }
        fieldColumnRange.values = fieldColumnRangeValues;
      }
    }
    await ctx.sync();
  }

  private async insertNewRows(
    workSheet: Excel.Worksheet,
    currentSheetFields: any[],
    changeLogNewArray: excelChangeLog[],
    objectNumberColValues: string[],
    ctx: any
  ) {
    console.log("Inserting new rows with changes:", changeLogNewArray);
    if (changeLogNewArray.length === 0) return;
    for (let changeLog of changeLogNewArray) {
      let field = currentSheetFields.find(m => m.fieldId == changeLog.columnId && m.metadata);
      if (!field) {
        console.log("Field not availabel in sheet:-", changeLog.columnId);
        continue;
      }
      const groupedMap = changeLog.columnRows?.reduce((map, m) => {
        if (m?.recordNumber) {
          if (!map.has(m.recordNumber as string)) {
            map.set(m.recordNumber as string, []);
          }

          map.get(m.recordNumber)?.push([m.Value]);
        }
        return map;
      }, new Map<string, any[][]>());
      for (let [key, newColValues] of groupedMap) {
        if (newColValues.length > 0) {
          let lastIndexofRecord = objectNumberColValues.lastIndexOf(key);
          const insertRowIndex = lastIndexofRecord + 1;

          // Insert a new row by shifting everything below down
          const insertRange = workSheet.getRangeByIndexes(insertRowIndex, 0, 1, 1);
          insertRange.insert(Excel.InsertShiftDirection.down);

          const recordColRange = workSheet.getRangeByIndexes(insertRowIndex, 1, 1, 1);
          recordColRange.numberFormat = [["@"]];
          recordColRange.values = [[key]];

          // Now target the column range in the newly inserted row
          const newColRange = workSheet.getRangeByIndexes(insertRowIndex, field.columnNumber, newColValues.length, 1);

          newColRange.numberFormat = [["@"]];   // set format to text
          newColRange.values = newColValues;    // fill with new values
        }
      }
    }
  }

  async updateModuleReferenceRecordData(data: any) {
    try {
      if (!data) { return }
      await this.utilsService.toggleEvents(false);
      const fieldsData: { [fieldId: string]: RecordField } = data.record;
      const rowIndex = data.rowIndex;
      // guard clauses
      if (!fieldsData) {
        console.warn('updateModuleReferenceRecordData: No records supplied.');
        return;
      }

      // Helper to extract value from record field vc collection
      const getCellValue = (recordField: any): string => {
        if (!recordField?.vc) return '';
        return Array.isArray(recordField.vc) ? recordField.vc.map((m: any) => m.t ? m.t : m.c).join(',') : recordField?.vc?.c;
      };

      await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        const selection = ctx.workbook.getSelectedRange();
        selection.load(['rowIndex']);
        const usedRange = sheet.getUsedRange();
        usedRange.load(['rowCount']);
        await ctx.sync();
        const totalRows = usedRange.rowCount;

        // Load full record number column once (C3..C{totalRows}) to compute last index map
        const recColStartRow = 2; // zero-based index for row 3
        const recColRowCount = Math.max(0, totalRows - recColStartRow);
        console.log(`Loading record number column from row ${recColStartRow + 1} for ${recColRowCount} rows`);
        const recColRange = recColRowCount > 0 ? sheet.getRangeByIndexes(recColStartRow, 2, recColRowCount, 1) : null;
        if (recColRange) recColRange.load('values');
        await ctx.sync();

        const recordNumbers = recColRange ? (recColRange.values as any[]).map(r => r[0]) : [];
        console.log('Record numbers loaded for reference:', recordNumbers);
        // Fixed column positions: Status = col 0, Message = col 1
        const boundMap = this.utilsService.getRecordIndexBoundsMap(recordNumbers);
        const activeRow = rowIndex ?? selection.rowIndex;
        const currentRecordNumber = recordNumbers[activeRow - recColStartRow];
        const recordBounds = boundMap.get(currentRecordNumber) || { first: activeRow, last: activeRow };

        console.log(`Updating recordNumber='${currentRecordNumber}' at rowIndex=${activeRow + 1} with bounds=`, recordBounds);
        const targetFields: any[] = this.sharedService.isExtentionProcess
          ? this.sharedService.staticFieldsObject
          : recordBounds.first == activeRow + 1
            ? this.sharedService.staticFieldsObject
            : this.sharedService.staticFieldsObject;
        const fieldMetaMap: Map<string, any> = new Map(targetFields.map((f: any) => [f.fieldId, f]));
        // Data expected to start at row index 3 (Excel row 4) per earlier convention
        if (activeRow < 2) {
          console.warn('updateModuleReferenceRecordData: Active row is within header area; aborting.');
          return;
        }

        const fieldIds = Array.from(fieldMetaMap.keys());

        // Collect updates to minimize sync calls
        for (const fieldId of fieldIds) {
          const meta = fieldMetaMap.get(fieldId);
          if (!meta) {
            console.warn(`updateModuleReferenceRecordData: No metadata found for fieldId=${fieldId}; skipping.`);
            continue; // Skip fields not present in fieldData
          }

          const recordField = fieldsData[fieldId];
          let cellValue = getCellValue(recordField);
          console.log(`FieldId='${fieldId}' at column=${meta.columnNumber} will be set to value='${cellValue}'`);
          if (!cellValue) { continue; }
          // Basic date / datetime formatting using existing pickList semantics if available in recordField
          // Since we do not have full Field metadata here, infer from value patterns (milliseconds length >= 12)
          if (cellValue && /^\d{12,}$/.test(cellValue)) {
            // Try converting to date (YYYY-MM-DD). Use existing utils if available.
            try {
              const ms = Number(cellValue);
              if (!isNaN(ms)) {
                const dateStr = this.utilsService.getDateFromMilliseconds(ms.toString(), DateTime.DATE);
                if (dateStr) cellValue = dateStr;
              }
            } catch { /* ignore formatting errors */ }
          }

          try {
            const cell = sheet.getCell(activeRow, meta.columnNumber);
            cell.numberFormat = [['@']]; // Force text format expects 2D array
            cell.values = [[cellValue]];
          } catch (err) {
            console.error(`Failed to set cell for fieldId=${fieldId} at column=${meta.columnNumber}`, err);
          }
        }
        await ctx.sync();
      }).catch((error: any) => {
        console.error('Excel error in updateModuleReferenceRecordData:', error);
        if (error instanceof OfficeExtension.Error) {
          console.error('Debug info:', JSON.stringify(error.debugInfo));
        }
      });
    }
    catch (error) {
      console.error('Unexpected error in updateModuleReferenceRecordData:', error);
    }
    finally {
      await this.utilsService.toggleEvents(true);
    }
  }
  async updateModuleReferenceRecordDataBulkRowSafe(dataList: any[], ctx: Excel.RequestContext) {

    console.log("Recevid sap for updating", dataList);
    if (!dataList?.length) return;

    await this.utilsService.toggleEvents(false);

    try {
      const getCellValue = (recordField: any): string => {
        if (!recordField?.vc) return '';
        return Array.isArray(recordField.vc)
          ? recordField.vc.map((m: any) => m.t ?? m.c).join(',')
          : recordField.vc?.c ?? '';
      };


      const sheet = ctx.workbook.worksheets.getActiveWorksheet();

      /* -----------------------------
       * 1. Row bounds
       * -----------------------------*/
      const rowIndexes = dataList.map(d => d.rowIndex);
      const minRow = Math.min(...rowIndexes);
      const maxRow = Math.max(...rowIndexes);
      const rowCount = maxRow - minRow + 1;

      /* -----------------------------
       * 2. Field metadata
       * -----------------------------*/
      const recordNumbers = await this.utilsService.getAllRecordNumbers(ctx)
      const boundMap = this.utilsService.getRecordIndexBoundsMap(recordNumbers);



      /* -----------------------------
       * 3. Group updates by column
       * -----------------------------*/
      const updatesByColumn = new Map<
        number,
        { rowOffset: number; value: string }[]
      >();

      for (const item of dataList) {
        const activeRow = item.rowIndex;
        const currentRecordNumber = recordNumbers[activeRow - 2];
        const recordBounds = boundMap.get(currentRecordNumber) || { first: activeRow, last: activeRow };

        console.log(`Updating recordNumber='${currentRecordNumber}' at rowIndex=${activeRow + 1} with bounds=`, recordBounds);
        const targetFields = this.sharedService.isExtentionProcess
          ? this.sharedService.staticFieldsObject
          : this.sharedService.staticFieldsObject;

        const fieldMetaMap = new Map(
          targetFields.map((f: any) => [f.fieldId, f])
        );
        for (const [fieldId, meta] of fieldMetaMap) {
          const recordField = item?.record[fieldId];
          if (!recordField) continue;

          let value = getCellValue(recordField);
          if (!value) continue;

          if (/^\d{12,}$/.test(value)) {
            const ms = Number(value);
            if (!isNaN(ms)) {
              const dateStr = this.utilsService.getDateFromMilliseconds(
                ms.toString(),
                DateTime.DATE
              );
              if (dateStr) value = dateStr;
            }
          }

          if (!updatesByColumn.has(meta.columnNumber)) {
            updatesByColumn.set(meta.columnNumber, []);
          }

          updatesByColumn.get(meta.columnNumber)!.push({
            rowOffset: item.rowIndex - minRow,
            value
          });
        }
      }

      /* -----------------------------
       * 4. Read → Patch → Write per column
       * -----------------------------*/
      for (const [colIndex, updates] of updatesByColumn) {
        const range = sheet.getRangeByIndexes(
          minRow,
          colIndex,
          rowCount,
          1
        );

        range.load('values');
        await ctx.sync();

        const colValues = range.values as any[][];

        for (const { rowOffset, value } of updates) {
          colValues[rowOffset][0] = value;
        }

        range.numberFormat = Array(rowCount).fill(['@']);
        range.values = colValues;
      }

      await ctx.sync();


    } catch (err) {
      console.error('Row-safe bulk update failed:', err);
    } finally {
      await this.utilsService.toggleEvents(true);
    }
  }



  /**
   * Populates Short Description and Long Description columns in the sheet for each record.
   * Locates description columns by their generated labels. No selected-form
   * field IDs are embedded in the runtime template.
   * Row resolution: uses hdvs['ROW_UUID'] value (assumed Excel 1-based row number). Falls back to skipping if missing.
   * Description resolution: prefers language 'en' if available; otherwise first language block.
   */
  async updateDescription(recordList?: MdoRecordES[], descriptionOnly: boolean = false): Promise<void> {
    try {
      if (!recordList?.length) {
        console.warn('updateDescription: empty record list');
        return;
        //recordList = descriptionRsponse as unknown as MdoRecordES[] || [];
      }

      const shortMeta = this.sharedService.staticFieldsObject.find(f =>
        String(f.label || '').trim().toLowerCase() === 'short description');
      const longMeta = this.sharedService.staticFieldsObject.find(f =>
        String(f.label || '').trim().toLowerCase() === 'long description');
      if (!shortMeta && !longMeta) {
        console.warn('updateDescription: description columns not found in fieldData');
        return;
      }
      else {
        console.log('updateDescription: found description columns:', {
          shortColumn: shortMeta?.columnNumber,
          longColumn: longMeta?.columnNumber
        });
      }

      await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
        try {
          const sheet = ctx.workbook.worksheets.getActiveWorksheet();

          for (const record of recordList) {
            const hdvs = record?.hdvs || {};
            let rowUUIDValue: number | undefined | string = hdvs['ROW_UUID']?.vc?.[0]?.c ?? 2;
            if (this.recordFirstRow.size > 0) {
              rowUUIDValue = this.recordFirstRow.get(record.id.toString());
            }
            const excelRowNumber = Number(rowUUIDValue);
            if (!Number.isFinite(excelRowNumber)) {
              console.log('updateDescription: skipping record with invalid ROW_UUID:', rowUUIDValue);
              continue; // cannot resolve row
            }

            // Convert to zero-based row index. If headers occupy rows 0 and 1, data may start at row 2.
            // ROW_UUID is assumed to be the Excel visible row number; subtract 1 for zero-based index.
            const rowIndex = this.recordFirstRow.size > 0 ? excelRowNumber : excelRowNumber + 1;
            if (rowIndex < 2) { // avoid overwriting header rows
              console.log('updateDescription: skipping record with ROW_UUID in header area:', excelRowNumber);
              continue;
            }

            let descriptions = record?.descriptions || [];
            try {
              descriptions = this.normalizeLongDescription(descriptions, record.id as string);
            }
            catch (err) {
              console.log(err);
            }
            let shortDesc = '';
            let longDesc = '';
            if (descriptions.length) {
              // console.log('updateDescription: processing descriptions for record with ROW_UUID:', excelRowNumber);
              // Aggregate across all classification descriptions; pick preferred language blocks.
              for (const descBlock of descriptions) {
                const attrsByLang = descBlock?.attributes || {};
                // Preferred language order
                const langs = Object.keys(attrsByLang);
                let chosenLang = langs.find(l => l.toLowerCase() === this.sharedService.globalConfig?.UserProfile?.selfServiceUserModel?.lang) || langs[0];
                if (chosenLang) {
                  const attrObj: any = attrsByLang[chosenLang];
                  const sd = (attrObj?.shortDesc || '').toString().trim();
                  const ld = (attrObj?.longDesc || '').toString().trim();
                  if (sd) {
                    shortDesc = shortDesc ? `${shortDesc} | ${sd}` : sd; // concatenate if multiple
                  }
                  if (ld) {
                    longDesc = longDesc ? `${longDesc} | ${ld}` : ld;
                  }
                  if (descriptionOnly) continue;
                }
              }
            }
            else {
              console.log('updateDescription: no descriptions found for record with ROW_UUID:', excelRowNumber);
            }
            let errorMessages = ''
            try {
              console.log("Description errors", (record as any).errorLogs);
              if ((record as any).errorLogs) {
                const errorLogs: [{
                  fieldId: string
                  errMsg: any
                  path: string
                  parentSegment: string
                  rowSno: string
                  errorMsg: string
                }] = (record as any).errorLogs
                for (let error of errorLogs) {
                  errorMessages += `${error.fieldId}:- ${error.errorMsg || ''}, ${error.errMsg || ''} \n`
                }
              }
            }
            catch (err) {
              console.log(err);
            }

            // Write values if metadata found
            try {
              if (shortMeta && shortDesc) {
                const cell = sheet.getCell(rowIndex, shortMeta.columnNumber);
                cell.numberFormat = [['@']];
                cell.values = [[shortDesc.replaceAll(`\\n`, '\n')]]; // handle escaped newlines
                cell.format.fill.color = '#d3d3d3';
                cell.format.protection.locked = true;
                console.log('updateDescription: wrote short description for row', excelRowNumber);
              }
              if (longMeta && longDesc) {
                const cell = sheet.getCell(rowIndex, longMeta.columnNumber);
                cell.numberFormat = [['@']];
                cell.values = [[longDesc.replaceAll(`\\n`, '\n')]]; // handle escaped newlines
                cell.format.fill.color = '#d3d3d3';
                cell.format.protection.locked = true;
                console.log('updateDescription: wrote long description for row', excelRowNumber);
              }
              await ctx.sync();
              if (errorMessages) {
                console.log('updating errors');
                const cell = sheet.getCell(rowIndex, 0);
                cell.numberFormat = [['@']];
                cell.values = [['ERROR']];
                const mCell = sheet.getCell(rowIndex, 1);
                mCell.numberFormat = [['@']];
                mCell.values = [[errorMessages]];
              }
            } catch (e) {
              console.warn('updateDescription: failed writing descriptions for row', excelRowNumber, e);
            }
          }
        } finally {
          await ctx.sync();
        }
      }).catch((error: any) => {
        console.error('Excel error in updateDescription:', error);
        if (error instanceof OfficeExtension.Error) {
          console.error('Debug info:', JSON.stringify(error.debugInfo));
        }
      });
    }
    catch (error) {
      console.error('Unexpected error in updateDescription:', error);
    }
  }
  private normalizeLongDescription(descriptions: ClassificationDescription[], recordNumber: string): ClassificationDescription[] {
    // The description response may include server-only classification metadata
    // that is not rendered into Excel. Merge it before writing cells so later
    // validate/submit requests keep the complete payload for this record.
    const mergedDescriptions = this.sharedService.mergeDescriptions(recordNumber, descriptions);
    for (let descRow of mergedDescriptions) {

      for (const lang of Object.keys(descRow.attributes)) {
        const langBlock = descRow.attributes[lang];

        if (typeof langBlock?.longDesc !== 'string') {
          continue;
        }

        const originalValue = langBlock.longDesc;

        try {
          langBlock.longDesc = originalValue
            .replace(/\\n/g, '\n')
            .replace(/\\t/g, '\t')
            .replace(/\\r/g, '\r');
        } catch (error) {
          langBlock.longDesc = originalValue;
          console.error(
            `Failed to normalize longDesc for language [${lang}]`,
            error
          );
        }
      }
    }
    // Persist the normalised values too; otherwise escaped long descriptions
    // would reappear in the next workbook request after the sheet is updated.
    return this.sharedService.mergeDescriptions(recordNumber, mergedDescriptions);
  }

  buildAttrColumns(attributes: Attributes, classCode: string): (string)[] {
    const columns: string[] = [];
    if (!this.sharedService.classificationEnabled) {
      return columns;
    }

    let classDet = this.sharedService.findClassAttributes(classCode);
    if (!classDet) {
      for (let [key, value] of this.sharedService.classAttributesMap) {
        const displayCode = value?.code && value?.mod ? `${value.code}--${value.mod}` : value?.code;
        if (String(displayCode ?? '').toLowerCase() === classCode.toLowerCase()
          || String(value?.codeLong ?? '').toLowerCase() === classCode.toLowerCase()) {
          classDet = value;
          break;
        }
      }
    }
    if (!classDet || !Array.isArray(classDet.charList) || classDet.charList.length === 0) {
      console.log('No class details found for class code:', classCode, "mapList:", this.sharedService.classAttributesMap);
      return columns;
    };

    // Filter active chars and sort by charOrder
    const sortedChars = classDet.charList
      .filter((c: any) => c && (c.status === true || String(c.status).toLowerCase() === 'true'))
      .slice()
      .sort((a: any, b: any) => Number(a.charOrder || 0) - Number(b.charOrder || 0));

    for (const ch of sortedChars) {
      let name = String(ch.charDesc || '').trim();
      if (ch.labels) {
        name = ch.labels.find((label: any) => label.language == this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang)?.label || name;
      }
      const uuid = ch.uuid;
      // Pull value from provided attributes map; if missing leave empty string
      const attrEntry = attributes?.attrs ? (attributes.attrs as any)[uuid] : undefined;
      const vcValue = (attrEntry?.vc && Array.isArray(attrEntry.vc) && attrEntry.vc[0]) ? (attrEntry.vc[0].c ?? '').toString() : '';
      columns.push(name, vcValue);
    }
    return columns;
  }

  async updateClassAttributes(rowIndex: number, attributeCols: string[], classCode: string) {
    if (!this.sharedService.classificationEnabled) {
      return;
    }

    await Excel.run({ delayForCellEdit: true }, async (ctx: Excel.RequestContext) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRange();
      usedRange.load(["rowCount", "columnCount"]);
      await ctx.sync();

      const classColumnIndex = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'CLASSIFICATION_CLASS')?.columnNumber;
      if (classColumnIndex === undefined) {
        return;
      }
      const cell = sheet.getCell(rowIndex, classColumnIndex);
      cell.numberFormat = [['@']];
      cell.values = [[classCode ?? ""]];
      console.log(attributeCols);
      if (attributeCols.length) {
        let startCol = classColumnIndex + 1;
        let lastCol = startCol + attributeCols.length;

        //update the column width for attributes column
        const col = sheet.getRangeByIndexes(0, startCol, 1, lastCol);
        col.format.columnWidth = 150;
        for (let i = 0; i < attributeCols.length; i++) {
          const val = attributeCols[i];
          const cell = sheet.getCell(rowIndex, startCol + i);
          cell.numberFormat = [['@']];
          cell.values = [[val ?? ""]];
          if (i % 2 === 0) {
            cell.format.font.bold = true;
          }
        }
      }
      await ctx.sync();
    })
  }

  async createAttributesInSheet(attrs: CharacteristicList[], rowIndex: number, context: Excel.RequestContext) {
    if (!this.sharedService.classificationEnabled) {
      return;
    }

    try {
      console.log("Class Attributes received:", attrs);
      if (attrs?.length === 0) {
        return;
      }


      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('id');
      const selection = context.workbook.getSelectedRange();
      selection.load(['rowIndex']);
      await context.sync();

      const activeRow = rowIndex ?? selection.rowIndex;

      // Do not load the complete workbook just to add a few attribute cells.
      // Office.js sync calls cross the Excel host boundary, so full-range reads
      // plus multiple intermediate syncs made the class response appear slow.
      const usedRange = sheet.getUsedRange();
      usedRange.load('columnCount');
      await context.sync();
      const classColumnIndex = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'CLASSIFICATION_CLASS')?.columnNumber;
      console.log("Class column index:", classColumnIndex);
      const recordNumberIndex = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'mdo_number')?.columnNumber;

      let startCol = classColumnIndex;
      if (startCol === undefined || startCol === -1) {
        console.error("CLASSIFICATION_CLASS column not found in sheet headers.");
        return;
      }

      if (activeRow !== undefined && activeRow !== null) {
        if (recordNumberIndex === undefined) {
          console.error('MDO Number column not found in sheet headers.');
          return;
        }
        const recordNumberCell = sheet.getCell(activeRow, recordNumberIndex);
        recordNumberCell.load('values');
        await context.sync();
        if (!String(recordNumberCell.values?.[0]?.[0] ?? '').trim()) return;

        const sortedAttrs = attrs.sort(
          (a: CharacteristicList, b: CharacteristicList) => Number(a.charOrder) - Number(b.charOrder)
        );
        const lastCol = startCol + 1 + sortedAttrs.length * 2;

        let col = classColumnIndex + 1;

        // Attribute labels are created at runtime, so they have no static
        // layout column. Remove the previous cell bindings before rebuilding
        // this row; each generated Value cell is registered below with its
        // exact characteristic metadata for the host dropdown.
        this.sharedService.clearRuntimeClassificationAttributes(sheet.id, activeRow, col);

        // Attribute cells are always to the right of Description Class. Clear
        // that generated area in one range operation; writing below extends the
        // worksheet automatically when this is the first selected class.
        const existingAttributeColumns = usedRange.columnCount - col;
        if (existingAttributeColumns > 0) {
          const attributeRange = sheet.getRangeByIndexes(activeRow, col, 1, existingAttributeColumns);
          attributeRange.clear();
          attributeRange.format.protection.locked = false;
        }

        for (const attribute of sortedAttrs) {
          let attributeName = String(attribute.charDesc || '').trim();
          if (attribute.labels) {
            attributeName = attribute.labels.find((label: any) => label.language == this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang)?.label || attributeName;
          }
          const attrCell = sheet.getCell(activeRow, col);
          attrCell.clear();
          attrCell.format.protection.locked = false;
          console.log("Adding attribute:", attributeName, "at row:", activeRow, "col:", col);
          attrCell.values = [[attributeName]];
          if (attribute.isManatory) {
            attrCell.format.font.color = '#ff0000';
          }
          attrCell.format.fill.color = "#a6c9ec"; // blue background for attributes
          attrCell.format.font.bold = true;
          attrCell.format.wrapText = true;
          attrCell.format.horizontalAlignment = "Center";
          attrCell.format.verticalAlignment = "Bottom";
          attrCell.format.columnWidth = 150;
          attrCell.format.protection.locked = true;

          const valueCell = sheet.getCell(activeRow, col + 1);
          valueCell.clear();
          valueCell.numberFormat = [['@']];
          valueCell.format.protection.locked = false;
          valueCell.values = [[""]];
          valueCell.format.wrapText = true;
          valueCell.format.verticalAlignment = "Bottom";
          valueCell.format.columnWidth = 150;

          // A list-type characteristic must behave exactly like a static
          // class attribute: on selecting its Value cell, the event handler
          // retrieves this metadata and ngx-mdo loads its allowed values.
          this.sharedService.rememberRuntimeClassificationAttribute(
            sheet.id,
            activeRow,
            col + 1,
            attribute
          );

          col += 2;
        }
      }
      await context.sync();

    }
    catch (err) {
      console.error("createAttributesInSheet failed:", err);
    }
  }

  /**
   * Applies Description Class attributes to several rows in one Office.js
   * batch. Dragging/filling a class used to invoke createAttributesInSheet for
   * every row, multiplying host sync calls and making a four-row fill slow.
   */
  async createAttributesInSheetBatch(
    rows: Array<{ rowIndex: number; attrs: CharacteristicList[] }>,
    context: Excel.RequestContext
  ): Promise<void> {
    if (!this.sharedService.classificationEnabled || !rows.length) return;
    try {
      const classColumnIndex = this.sharedService.staticFieldsObject
        .find(field => field.fieldId === 'CLASSIFICATION_CLASS')?.columnNumber;
      const recordNumberIndex = this.sharedService.staticFieldsObject
        .find(field => field.fieldId === 'mdo_number')?.columnNumber;
      if (classColumnIndex === undefined || recordNumberIndex === undefined) return;

      const uniqueRows = Array.from(new Map(rows.map(row => [row.rowIndex, row])).values())
        .filter(row => row.rowIndex >= 2 && row.attrs?.length);
      if (!uniqueRows.length) return;

      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load('id');
      const usedRange = sheet.getUsedRange();
      usedRange.load('columnCount');
      const recordCells = uniqueRows.map(row => sheet.getCell(row.rowIndex, recordNumberIndex));
      recordCells.forEach(cell => cell.load('values'));
      await context.sync();

      const attributeStartColumn = classColumnIndex + 1;
      for (let index = 0; index < uniqueRows.length; index++) {
        const row = uniqueRows[index];
        if (!String(recordCells[index].values?.[0]?.[0] ?? '').trim()) continue;

        const sortedAttrs = [...row.attrs].sort((a, b) => Number(a.charOrder) - Number(b.charOrder));
        this.sharedService.clearRuntimeClassificationAttributes(sheet.id, row.rowIndex, attributeStartColumn);
        const existingAttributeColumns = usedRange.columnCount - attributeStartColumn;
        if (existingAttributeColumns > 0) {
          const staleRange = sheet.getRangeByIndexes(row.rowIndex, attributeStartColumn, 1, existingAttributeColumns);
          staleRange.clear();
          staleRange.format.protection.locked = false;
        }

        let columnIndex = attributeStartColumn;
        for (const attribute of sortedAttrs) {
          let attributeName = String(attribute.charDesc || '').trim();
          if (attribute.labels) {
            attributeName = attribute.labels.find((label: any) =>
              label.language === this.sharedService.globalConfig.UserProfile?.selfServiceUserModel?.lang
            )?.label || attributeName;
          }
          const nameCell = sheet.getCell(row.rowIndex, columnIndex);
          nameCell.clear();
          nameCell.values = [[attributeName]];
          nameCell.format.protection.locked = true;
          nameCell.format.fill.color = '#a6c9ec';
          nameCell.format.font.color = attribute.isManatory ? '#ff0000' : '#000000';
          nameCell.format.font.bold = true;
          nameCell.format.wrapText = true;
          nameCell.format.horizontalAlignment = 'Center';
          nameCell.format.verticalAlignment = 'Bottom';
          nameCell.format.columnWidth = 150;

          const valueCell = sheet.getCell(row.rowIndex, columnIndex + 1);
          valueCell.clear();
          valueCell.numberFormat = [['@']];
          valueCell.values = [['']];
          valueCell.format.protection.locked = false;
          valueCell.format.wrapText = true;
          valueCell.format.verticalAlignment = 'Bottom';
          valueCell.format.columnWidth = 150;
          this.sharedService.rememberRuntimeClassificationAttribute(sheet.id, row.rowIndex, columnIndex + 1, attribute);
          columnIndex += 2;
        }
      }
      await context.sync();
    } catch (error) {
      console.error('createAttributesInSheetBatch failed:', error);
    }
  }

  /**
   * Normal elements retain the historical replacement behaviour. MOCR calls
   * this with preserveExistingSheets=true so independently configured roots
   * are never deleted while their workbook headers are created.
   */
  async createHeader(isHeaderVisible: boolean = true, preserveExistingSheets: boolean = false, sheetName?: string) {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        // Delete all visible sheets
        const worksheets = context.workbook.worksheets;
        const sanitizedName = this.sharedService.data.formDesc.replace(/[\\\/\?\*\[\]:]/g, '');
        const newSheetName = (sheetName || sanitizedName).substring(0, 31);
        console.log("Creating new sheet with name:", newSheetName);
        this.sharedService.currentSheetName = newSheetName;

        worksheets.load('items/name, items/id, items/visibility');
        await context.sync();
        // Create new sheet with name from data.formDesc
        let sheet = null;
        if (worksheets.items.some(ws => ws.name === newSheetName)) {
          console.log(`Sheet with name ${newSheetName} already exists. Reusing it.`);
          const existingSheet = worksheets.getItem(newSheetName);
          sheet = existingSheet;
          await context.sync();
        } else {
          const newSheet = context.workbook.worksheets.add(newSheetName);
          sheet = newSheet;
          await context.sync();
        }

        // Collect sheets to delete (all visible sheets except the new one)
        const sheetsToDelete = preserveExistingSheets ? [] : worksheets.items.filter(ws =>
          ws.visibility === Excel.SheetVisibility.visible && ws.id !== sheet.id
        );

        // Delete sheets one at a time with sync after each deletion
        for (const ws of sheetsToDelete) {
          console.log(`Deleting sheet: ${ws.name}`);
          ws.delete();
          await context.sync();
        }


        // Header rendering must use the immutable generated template order and
        // visual groups. Runtime layout/grid metadata is used for record
        // parsing only; it must never rename, regroup, or reorder columns.
        const fields = [...this.sharedService.staticFieldsObject]
          .filter(f => f.columnNumber >= 0 && f.column !== null)
          .sort((left, right) => (left.columnNumber ?? Number.MAX_SAFE_INTEGER) - (right.columnNumber ?? Number.MAX_SAFE_INTEGER));
        if (isHeaderVisible === true) {
          const fullRange = sheet.getRange("A1:XFD1048576");
          fullRange.format.protection.locked = false;
        }

        const secondRow: string[] = fields.map(f => f.label ?? "");
        const colCount = secondRow.length;

        // getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount)
        const row2Range = sheet.getRangeByIndexes(isHeaderVisible ? 1 : 0, 0, 1, colCount);
        row2Range.values = [secondRow];


        // Format row2
        for (let i = 0; i < fields.length; i++) {
          const f = fields[i];
          const cell = sheet.getCell(isHeaderVisible ? 1 : 0, i);
          const fieldMetadata = this.sharedService.fieldsMap.get(f.fieldId)?.metadata?.[0];
          if (f.fieldId === 'mdo_number' || fieldMetadata?.dataType === 'TEXT') {
            cell.getEntireColumn().numberFormat = [['@']];
          }
          if (fieldMetadata?.dataType === 'DEC') {
            try {
              sheet.getRange(`${f.column}:${f.column}`).numberFormat = [["0.000"]];
            } catch (nfErr) {
              console.warn(`Failed to set decimal numberFormat for column ${f.column}:`, nfErr);
            }
          }
          cell.format.font.set({
            name: "Calibri",
            size: 11,
            bold: f.bold,
            color: f.textColor
          });

          cell.format.fill.color = f.bgColor || "#a6c9ec"; // default light blue
          // The generated workbook owns layout. Template import may supply
          // colours only; alignment and wrapping keep their original defaults.
          cell.format.horizontalAlignment = "Center";
          cell.format.verticalAlignment = "Bottom";
          cell.format.wrapText = true;

          cell.format.borders.getItem("EdgeLeft").style = "Continuous";
          cell.format.borders.getItem("EdgeRight").style = "Continuous";
          cell.format.borders.getItem("EdgeBottom").style = "Continuous";
          sheet.getRange(`${f.column}:${f.column}`).format.columnWidth = f.width;

          // Hide the entire column in Excel if the field is marked as hidden
          if (f.hidden) {
            sheet.getRange(`${f.column}:${f.column}`).columnHidden = true;
          }
        }
        await context.sync();
        if (isHeaderVisible === true) {
          // Recreate only the contiguous merged groups declared by the generated
          // metadata. Template columns without a group intentionally remain blank.
          let start = 0;
          while (start < fields.length) {
            const firstField = fields[start];
            const groupName = firstField.group;
            if (!groupName) {
              start++;
              continue;
            }
            let end = start + 1;
            while (end < fields.length && fields[end].group === groupName) {
              end++;
            }
            const span = end - start;

            const mergeRange = sheet.getCell(0, start).getResizedRange(0, span - 1);

            //no old merge issue
            mergeRange.unmerge();
            if (span > 1) {
              mergeRange.merge();
            }

            // ["Group", "", "", ""]
            const row = [groupName || "", ...new Array(span - 1).fill("")];
            mergeRange.values = [row];

            //format header groups
            mergeRange.format.font.set({
              name: "Calibri",
              size: 11,
              bold: true,
              color: firstField.groupTextColor || "black"
            });
            mergeRange.format.fill.color = firstField.groupBgColor || this.sharedService.getGroupColor(groupName);
            mergeRange.format.borders.getItem("EdgeLeft").style = "Continuous";
            mergeRange.format.borders.getItem("EdgeRight").style = "Continuous";
            mergeRange.format.borders.getItem("EdgeBottom").style = "Continuous";
            mergeRange.format.horizontalAlignment = "Center";
            mergeRange.format.verticalAlignment = "Center";
            mergeRange.format.wrapText = true;
            start = end;
          }
        }

        // MOCR creates several independent worksheets. Keep the user's active
        // sheet unchanged while writing their headers in the background.
        if (!preserveExistingSheets) {
          sheet.activate();
        }
        await context.sync();
      });
    } catch (err) {
      console.error("create header failed:", err);
    }
  }
  /**
   * Apply conditional formatting so that if the Success/Error message contains
   * "<FieldLabel>:" the corresponding field column cells are highlighted automatically.
   * This avoids imperative coloring and lets Excel update highlights as messages change.
   */
  private async applyConditionalFormattingBasedOnMessages() {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        // Find message column letter from fieldData (fieldId 'success_error' preferred)
        const messageMeta = this.sharedService.staticFieldsObject.find(f => f.fieldId === 'success_error')
          || this.sharedService.staticFieldsObject.find(f => (f.label ?? '').toString().trim().toLowerCase() === 'success/error');
        const messageCol = messageMeta?.column || 'B';

        // For each field (except status/message), add a custom CF on that column
        for (const fd of this.sharedService.staticFieldsObject) {
          if (!fd.column || !fd.label) continue;
          if (fd.fieldId === 'status' || fd.fieldId === 'success_error') continue;

          const col = fd.column;
          const label = fd.label.toString();
          const escapedLabel = label.replace(/"/g, '""');

          // Target entire column so it also covers rows added later
          const target = sheet.getRange(`${col}:${col}`);

          // Add a custom formula: highlight if message cell contains "<Label>:"
          // Use INDEX with ROW() to reference the message cell on the same row, across the whole column
          const cf = target.conditionalFormats.add(Excel.ConditionalFormatType.custom);
          cf.custom.rule.formula = `=ISNUMBER(SEARCH(\"${escapedLabel}:\",INDEX($${messageCol}:$${messageCol},ROW())))`;
          cf.custom.format.fill.color = '#ffcccc'; // light red
        }
        await context.sync();
      });
    } catch (e) {
      console.warn('applyConditionalFormattingBasedOnMessages error:', e);
    }
  }
  /**
   * Writes status and extracted error messages into the sheet.
   * Assumptions:
   *  - Second header row (row index 1) contains column labels (created in createHeader).
   *  - Columns whose header text (case-insensitive) match one of:
   *      Status: ['status', 'record status']
   *      Message: ['message', 'messages', 'error message', 'errors']
   *  - controlData.excelRow is 1-based; Excel API uses 0-based rowIndex.
   */
  async updateStatusAndMessages(records: MdoRecord[], requestType?: RequestType, sheetName?: string) {
    if (!records?.length) return;
    try {
      // Clear Status and Message columns before starting a new validation run
      // if (requestType != RequestType.SentForCorrection) {
      //   await this.clearStatusAndMessageColumns();
      // }
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        const sheet = sheetName
          ? context.workbook.worksheets.getItem(sheetName)
          : context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(['rowCount']);
        await context.sync();
        const totalRows = usedRange.rowCount;

        // Load full record number column once (C3..C{totalRows}) to compute last index map
        const recColStartRow = 2; // zero-based index for row 3
        const recColRowCount = Math.max(0, totalRows - recColStartRow);
        const recColRange = recColRowCount > 0 ? sheet.getRangeByIndexes(recColStartRow, 2, recColRowCount, 1) : null;
        const chidlColRange = recColRowCount > 0 ? sheet.getRangeByIndexes(recColStartRow, 3, recColRowCount, 1) : null;
        if (chidlColRange) chidlColRange.load('values');
        if (recColRange) recColRange.load('values');
        await context.sync();

        const recordNumbers = recColRange ? (recColRange.values as any[]).map(r => r[0]) : [];
        const chidlRecordNumbers = chidlColRange ? (chidlColRange.values as any[]).map(r => r[0]) : [];
        // Fixed column positions: Status = col 0, Message = col 1
        const boundMap = this.utilsService.getRecordIndexBoundsMap(recordNumbers);
        const childBoundMap = this.utilsService.getRecordIndexBoundsMap(chidlRecordNumbers);
        console.log('updateStatusAndMessages: boundMap:', boundMap);
        await this.utilsService.toggleEvents(false);

        if (requestType !== RequestType.SentForCorrection) {
          for (let [record, bound] of boundMap) {

            const rowCount = bound.last - bound.first
            if (rowCount === 0) continue;
            const targetRange = sheet.getRangeByIndexes(bound.first, 0, rowCount, 2);

            // Fill all rows in the selected range
            targetRange.values = Array(rowCount).fill(['', '']);
          }
        }

        for (const rec of records) {

          const cd = rec?.controlData;
          const excelRow = cd?.excelRow; // 1-based
          if (!excelRow || Number.isNaN(Number(excelRow))) continue;
          const rowIndex = Number(excelRow) + 1; // convert to 0-based
          let statusVal = cd?.status ?? '';
          if (requestType && requestType == RequestType.SentForCorrection && statusVal == "SUCCESS") {
            statusVal = "Sent for correction";
          }
          const recordNumber = cd?.recordNumber ?? '';
          if (requestType != RequestType.SentForCorrection) {
            this.sharedService.controlDataMap.set(recordNumber, cd);
          }
          let extractedMessges: { parent: string, child: Map<number, string[]> } = cd?.excelValidationResponse ? this.extractErrorMessages(cd.excelValidationResponse) : { parent: '', child: new Map<number, string[]>() };
          let msgVal = extractedMessges?.parent;
          if (requestType == RequestType.SentForCorrection) {
            msgVal = cd.statusMsg || '';
          }
          try {
            // Range covering Status & Message columns for this row
            const targetRange = sheet.getRangeByIndexes(rowIndex, 0, 1, 2);
            targetRange.values = [[statusVal, msgVal]];
            if (recordNumber) {
              await this.writeRecordNumberInsheet(sheet, context, recordNumber, boundMap, rowIndex, cd);
            }

            if (extractedMessges.child.size > 0) {
              for (let [rowIndex, messageList] of extractedMessges.child) {
                const childRowIndex = Number(rowIndex) + 1
                const ctargetRange = sheet.getRangeByIndexes(childRowIndex, 0, 1, 2);
                const childMessage = messageList?.join('\n');
                ctargetRange.values = [[statusVal, childMessage]];
              }
            }
          }
          catch (err) {
            console.error("Error while writing", err);
          }
          // Highlight cells that are in error for this row, clear highlight otherwise.
          try {
            const childRowColMap: Map<number, number[]> = new Map<number, number[]>()
            const errorCols = new Set<number>();
            const validation: any = cd?.excelValidationResponse;
            if (validation && typeof validation === 'object') {
              for (const key of Object.keys(validation)) {
                const arr = validation[key];
                if (!Array.isArray(arr)) continue;
                for (const item of arr) {
                  if (!item || typeof item !== 'object') continue;
                  let colNum: number | null = null;
                  // Prefer fieldId mapping when available
                  if (item.fieldId) {
                    const fld = this.sharedService.staticFieldsObject.find((f: any) => f.fieldId === item.fieldId);
                    if (fld && typeof fld.columnNumber === 'number') colNum = fld.columnNumber;
                  }
                  // Fallback: try matching by fieldDesc/fieldDesc-like text against header labels
                  if (colNum === null) {
                    const desc = (item.fieldDesc ?? item.fieldId ?? '').toString().trim().toLowerCase();
                    if (desc) {
                      const fld2 = this.sharedService.staticFieldsObject.find((f: any) => String(f.label ?? '').trim().toLowerCase() === desc);
                      if (fld2 && typeof fld2.columnNumber === 'number') colNum = fld2.columnNumber;
                    }
                  }
                  if (colNum !== null) {
                    errorCols.add(colNum);
                    try {
                      if (item.childRecord) {
                        const existCols = childRowColMap.get(item.childReferenceRowId);
                        if (existCols && existCols.length > 0) {
                          existCols.push(colNum);
                        }
                        else {
                          childRowColMap.set(item.childReferenceRowId, [colNum]);
                        }
                      }
                    }
                    catch (ex) {
                      console.log(ex);
                    }
                  }
                }
              }
            }

            // Iterate all fields and apply/clear highlight based on presence in errorCols
            for (const fd of this.sharedService.staticFieldsObject) {
              const col = fd?.columnNumber;
              if (col === undefined || col === null || col < 0) continue;
              const cellToCheck = sheet.getCell(rowIndex, col);
              if (errorCols.has(col)) {
                cellToCheck.format.fill.color = '#ffcccc';
                cellToCheck.getEntireRow().format.rowHeight = 30;
              } else if (!this.sharedService.alwaysReadOnly.includes(col) && !this.sharedService.readOnlyColumnsOnChange.includes(col) && !this.sharedService.readonlyFieldCols.includes(col)) {
                cellToCheck.format.fill.clear();
              }
              try {
                for (let [row, cols] of childRowColMap) {
                  for (const fd of this.sharedService.staticFieldsObject) {
                    const col = fd?.columnNumber;
                    if (col === undefined || col === null || col < 0) continue;
                    const cellToCheck = sheet.getCell(row, col);
                    if (cols.includes(col)) {
                      cellToCheck.format.fill.color = '#ffcccc';
                      cellToCheck.getEntireRow().format.rowHeight = 30;
                    } else if (!this.sharedService.alwaysReadOnly.includes(col)) {
                      cellToCheck.format.fill.clear();
                    }
                  }
                }
              }
              catch (ex) {
                console.error(ex);
              }
            }
            await context.sync();
          } catch (errHighlight) {
            console.error('Failed to apply/clear error highlights for row', rowIndex, errHighlight);
          }
        }
        return context.sync();
      });
    } catch (err) {
      console.error('updateStatusAndMessages failed:', err);
    }
  }

  async writeRecordNumberInsheet(sheet: Excel.Worksheet, context: Excel.RequestContext, recordNumber: string, boundMap: any, rowIndex: number, cd: ControlData) {

    let arbitoryNo = cd.arbitraryRecordNo;
    if (!arbitoryNo) {
      const rowRange = sheet.getRangeByIndexes(rowIndex, 2, 1, 1);
      rowRange.load('values');
      await context.sync();
      const values = rowRange.values;
      arbitoryNo = values[0][0];
    }
    console.log('updateStatusAndMessages: writing arbitraryRecordNo for', arbitoryNo, 'at rowIndex', rowIndex);
    console.log('updateStatusAndMessages: boundMap for arbitraryRecordNo:', boundMap.get(arbitoryNo));
    const lastIndex = boundMap.get(arbitoryNo)?.last;
    console.log('updateStatusAndMessages: writing recordNumber for', recordNumber, 'at rowIndex', rowIndex, 'lastIndex', lastIndex);
    const targetRange2 = sheet.getRangeByIndexes(rowIndex, 2, lastIndex ? lastIndex - rowIndex : 1, 1);
    targetRange2.load('values');
    await context.sync();
    let valueArray = targetRange2.values;
    console.log('Current values in recordNumber column:', valueArray);
    for (let i = 0; i < valueArray.length; i++) {
      valueArray[i][0] = recordNumber;
    }

    targetRange2.values = valueArray;
  }
  // Collect all error arrays that follow the same shape (objects with fieldDesc & error)
  extractErrorMessages(validationResponse: any): { parent: string, child: Map<number, string[]> } {
    if (!validationResponse || typeof validationResponse !== 'object') return { parent: '', child: new Map<number, string[]>() }

    // Keys we expect might hold arrays of field error objects
    const possibleErrorArrays = Object.keys(validationResponse);

    const lines = [];
    const childLinesMap: Map<number, string[]> = new Map<number, string[]>();
    for (const key of possibleErrorArrays) {
      const arr = validationResponse[key];
      if (!Array.isArray(arr)) continue;

      for (const item of arr) {
        if (!item || typeof item !== 'object') continue;
        const fieldDesc = (item.fieldDesc ?? item.fieldId ?? '').toString().trim();
        const errorMsg = (item.error ?? '').toString().trim();
        if (fieldDesc && errorMsg) {
          if (item.childRecord) {
            const exisitng = childLinesMap.get(item.childReferenceRowId);//childReferenceRowId
            if (exisitng) {
              exisitng.push(`${fieldDesc}:${errorMsg}`);
            }
            else {
              childLinesMap.set(item.childReferenceRowId, [`${fieldDesc}:${errorMsg}`]);
            }
          }
          else {
            lines.push(`${fieldDesc}:${errorMsg}`);
          }
        }
        else if (errorMsg) {
          if (item.childRecord) {
            const exisitng = childLinesMap.get(item.childReferenceRowId);//childReferenceRowId
            if (exisitng) {
              exisitng.push(`${errorMsg}`);
            }
            else {
              childLinesMap.set(item.childReferenceRowId, [`${errorMsg}`]);
            }
          }
          else {
            lines.push(`${errorMsg}`);
          }

        }
      }
    }

    return { parent: lines.join('\n'), child: childLinesMap }
  }
  // Clears Status (col 0) and Success/Error (col 1) values for all existing data rows
  private async clearStatusAndMessageColumns() {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load('rowCount');
        await context.sync();

        // Data starts after header rows. Assuming headers occupy first 2 rows (0 and 1),
        // data begins at index 2 (0-based) => Excel row 3.
        const dataStartRowIndex = 2; // 0-based
        const totalRows = Math.max(usedRange.rowCount - dataStartRowIndex, 0);
        if (totalRows <= 0) return;

        const target = sheet.getRangeByIndexes(dataStartRowIndex, 0, totalRows, 2);
        const blanks = Array.from({ length: totalRows }, () => ["", ""]);
        target.values = blanks;
        return context.sync();
      });
    } catch (e) {
      console.warn('clearStatusAndMessageColumns failed:', e);
    }
  }
  async setSelectedDropDownValues(data: { field: Field, value: { code: string, text: string } }) {
    try {
      console.log("setSelectedDropDownValues data:", data);
      await Excel.run({ delayForCellEdit: true }, async (context: Excel.RequestContext) => {
        const activeCell = context.workbook.getActiveCell();
        // Check if column description matches any field that should use code
        const shouldUseCode = this.sharedService.codeFields.includes(data.field?.fieldId || '');
        const valueToSet = shouldUseCode ? data.value.code : data.value.text;
        if (valueToSet === undefined || valueToSet === null) {
          console.warn("Value to set is undefined or null, skipping cell update.");
          return;
        }
        activeCell.numberFormat = [['@']];
        activeCell.values = [[valueToSet]];
        activeCell.load('rowIndex');
        await context.sync();
      });
    } catch (error: any) {
      console.error('Error in setSelectedDropDownValues:', error);
      if (error instanceof OfficeExtension.Error) {
        console.error('Debug info:', JSON.stringify(error?.debugInfo));
      }
    }
  }

  async processGeneratedDescription(event: any) {
    if (!this.sharedService.classificationEnabled) {
      return;
    }

    if (event) {
      this.recordFirstRow = new Map<string, number>();
      console.log("Generate Description Response received:", event.data);
      await this.utilsService.toggleEvents(false);
      // Always restore workbook events. A failed description update must not
      // leave the generated element unable to process later sheet changes.
      try {
        await this.updateDescription(event.data, true);
      } finally {
        await this.utilsService.toggleEvents(true);
      }
    }
    else {
      console.log("description timeout event not received")
    }
  }

  /** Uses the active element's generated classification metadata only. */
  private getConfiguredClassificationType(): string {
    const classField = this.sharedService.fieldsMap.get('CLASSIFICATION_CLASS');
    const metadata = classField?.metadata?.[0] as any;
    return String(metadata?.classTypeId ?? metadata?.classType ?? '');
  }

}

export class excelChangeLog {
  columnId?: string;
  columnRows: columnDataRow[] = [];
}

export interface columnDataRow {
  rowid: number;
  Value: string;
  recordNumber: string;
}
