import { Injectable } from '@angular/core';
import { RequestType, MdoRecord, SubmitRequest, daxeXlEvent, DaxeXLBaseEvents, SubmitTriggerRequest } from 'daxe-ui-sdk';
import { ExcelReaderService } from './excel-reader.service';
import { UtilsService } from './utils.service';
import { SharedService } from './shared.service';
import { ExcelDataWriterService } from './excel-data-writer.service';
import { TranslationService } from './translation.service';
import { SheetContext, SheetContextStore } from './sheet-context-store.service';

@Injectable({
  providedIn: 'root'
})
export class DaxeXLService {

  constructor(private readonly excelReader: ExcelReaderService,
    private readonly utilsService: UtilsService,
    private readonly excelWriter: ExcelDataWriterService,
    private readonly sharedService: SharedService,
    private readonly translationService: TranslationService
  ) { }

  async validateRecordsBatch() {
    let totalRecordsToVaidate = 0;
    await this.excelReader.processSheetRecordWise(RequestType.Validation, async ({ totalRecords, batchRecordCount, recordList }: { totalRecords: number; batchRecordCount: number; recordList: MdoRecord[] }) => {
      totalRecordsToVaidate = totalRecords;
      const submitRequest: SubmitRequest = {
        recordList,
        totalRecords,
        recordsInBatch: batchRecordCount,
        processingId: this.sharedService.data.controlData.massId,
        requestType: RequestType.Export
      };
      daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSubmitRecordsRequest, payload: submitRequest });
    });
    const submitRequest: SubmitTriggerRequest = {
      totalRecords: totalRecordsToVaidate,
      processingId: this.sharedService.data.controlData.massId,
      requestType: RequestType.Validation
    };
    daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSubmitRecordsTriggerRequest, payload: submitRequest });
  }

  async validateRecordsMass(requestType: RequestType = RequestType.Validation) {
      this.utilsService.displayStatus(this.translationService.translate('Preparing records for {{requestType}}...', { requestType: String(requestType) }));
      let recordList = await this.excelReader.buildRecordList(requestType) as MdoRecord[] | null;
      if(!recordList?.length){
        this.utilsService.displayStatus(null);
        this.sharedService.isProcessRunning = false;
        return;
      }
      if(recordList.length > 0) {
        await this.mergeDynamicClassificationSheet(recordList);
        this.utilsService.displayStatus(null);
        const submitRequest: SubmitRequest = {
          recordList,
          requestType: requestType
        } as SubmitRequest;
        daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSubmitRecordsRequest, payload: submitRequest });
      }
  }

  /**
   * Joins rows from the generated Dynamic Class Attributes sheet to parent
   * records by Record Number. Each row represents one selected characteristic
   * and its value; imported/description/static classifications are preserved
   * through SharedService.mergeDescriptions.
   */
  private async mergeDynamicClassificationSheet(records: MdoRecord[]): Promise<void> {
    const config = this.sharedService.dynamicClassificationSheet;
    if (!config?.sheetName || !records?.length) return;

    const fields = config.fields || [];
    const columnOf = (fieldId: string) => fields.find((field: any) => field.fieldId === fieldId)?.columnNumber;
    const recordColumn = columnOf('mdo_number');
    const classColumn = columnOf(config.classFieldId);
    const characteristicColumn = columnOf('CHARAC');
    const valueColumn = columnOf('VALUE');
    if (![recordColumn, classColumn, characteristicColumn, valueColumn].every(Number.isFinite)) return;

    const rows = await Excel.run({ delayForCellEdit: true }, async (context) => {
      const sheet = context.workbook.worksheets.getItemOrNullObject(config.sheetName);
      sheet.load('isNullObject');
      await context.sync();
      if (sheet.isNullObject) return [] as any[][];
      const range = sheet.getUsedRangeOrNullObject(true);
      range.load(['isNullObject', 'values']);
      await context.sync();
      return range.isNullObject ? [] as any[][] : range.values as any[][];
    });

    const selected = rows.slice(2)
      .map(row => ({
        recordNumber: String(row[recordColumn] ?? '').trim(),
        classCode: String(row[classColumn] ?? '').trim(),
        characteristic: String(row[characteristicColumn] ?? '').trim(),
        value: row[valueColumn] ?? ''
      }))
      .filter(row => row.recordNumber && row.classCode && row.characteristic);
    if (!selected.length) return;

    const classType = String(config.classType || '');
    const missingClasses = Array.from(new Map(
      selected
        // Use the complete Excel value. It can be `code--mode`, while a
        // Description Class does not always have a CLASS_TYPE.
        .filter(row => !this.sharedService.findClassAttributes(row.classCode, classType))
        .map(row => [`${row.classCode}\u0000${classType}`, { classCode: row.classCode, classType }])
    ).values());
    if (missingClasses.length) await this.sharedService.updateAttributeMap(missingClasses);

    const language = this.sharedService.profileLanguage || this.sharedService.allLanguages?.[0] || 'en';
    const byRecord = new Map<string, any[]>();
    for (const row of selected) {
      const classDetails: any = this.sharedService.findClassAttributes(row.classCode, classType);
      if (!classDetails?.uuid) continue;
      const characteristicCode = row.characteristic.includes('--')
        ? row.characteristic.split('--').pop()?.trim()
        : row.characteristic;
      const characteristic = classDetails.charList?.find((item: any) => item?.charCode === characteristicCode);
      if (!characteristic?.uuid) continue;
      const description = {
        classCode: classDetails.code || row.classCode,
        classType: classDetails.classType || classType,
        classId: classDetails.uuid,
        classMode: String((classDetails as any).mod ?? (classDetails as any).classMode ?? ''),
        isDesc: false,
        attributes: {
          [language]: {
            shortDesc: '',
            longDesc: '',
            attrs: {
              [characteristic.uuid]: { fId: characteristic.charCode, vc: [{ c: row.value }] }
            }
          }
        }
      };
      const list = byRecord.get(row.recordNumber) || [];
      list.push(description);
      byRecord.set(row.recordNumber, list);
    }

    for (const record of records as any[]) {
      const controlData = record?.controlData || {};
      const recordNumber = String(controlData.arbitraryRecordNo || controlData.recordNumber || '').trim();
      const descriptions = byRecord.get(recordNumber);
      if (!descriptions?.length) continue;
      this.sharedService.mergeDescriptions(recordNumber, descriptions);
      record.mdoRecordES = record.mdoRecordES || {};
      record.mdoRecordES.descriptions = this.sharedService.getDescriptions(recordNumber, record.mdoRecordES.descriptions || []);
    }
  }

  /** Reads every MOCR root sheet and emits one combined request. */
  async validateMocrRecordsMass(contextStore: SheetContextStore, requestType: RequestType = RequestType.Validation) {
    this.utilsService.displayStatus(this.translationService.translate('Preparing MOCR records for {{requestType}}...', { requestType: String(requestType) }));
    const records: MdoRecord[] = [];
    for (const sheetName of contextStore.getAllSheetNames()) {
      if (!contextStore.restoreSheetContext(sheetName, this.sharedService)) {
        continue;
      }
      // Read the named worksheet directly. Activating it would visibly flicker
      // through every MOCR root and could disrupt the user's current sheet.
      const rootRecords = await this.excelReader.buildRecordList(requestType, sheetName) as MdoRecord[] | null;
      if (rootRecords?.length) {
        records.push(...rootRecords);
      }
      contextStore.snapshotContext(sheetName, this.sharedService);
    }
    if (!records.length) {
      this.utilsService.displayStatus(null);
      this.sharedService.isProcessRunning = false;
      return;
    }
    this.utilsService.displayStatus(null);
    daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSubmitRecordsRequest, payload: { recordList: records, requestType } as SubmitRequest });
  }

  /**
   * Normal multi-sheet workbooks use one form context, but spread its header,
   * grid and hierarchy values across worksheets. Read every sheet by context,
   * then merge contributions with the same Record Number before calling xFuse.
   */
  async validateWorkbookRecordsMass(contextStore: SheetContextStore, requestType: RequestType = RequestType.Validation) {
    this.utilsService.displayStatus(this.translationService.translate('Preparing workbook records for {{requestType}}...', { requestType: String(requestType) }));
    const records = await this.buildJoinedWorkbookRecords(contextStore, requestType);
    if (!records.length) {
      this.utilsService.displayStatus(null);
      this.sharedService.isProcessRunning = false;
      return;
    }
    this.utilsService.displayStatus(null);
    daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLSubmitRecordsRequest, payload: { recordList: records, requestType } as SubmitRequest });
  }

  /** Applies defaults to the same joined records used for validation/save. */
  async applyWorkbookDataRules(contextStore: SheetContextStore) {
    const records = await this.buildJoinedWorkbookRecords(contextStore, RequestType.Defaults);
    if (!records.length) {
      this.sharedService.isProcessRunning = false;
      return;
    }
    daxeXlEvent.next({ type: DaxeXLBaseEvents.DAXEXLApplyDataRulesRequest, payload: { recordList: records } });
  }

  /**
   * Restores the source-sheet row for each returned record. The backend works
   * with one joined record; Excel writes must still occur on each worksheet
   * that contributed data for that Record Number.
   */
  getWorkbookResponseRecords(contextStore: SheetContextStore, sheetName: string, records: MdoRecord[]): MdoRecord[] {
    const context = contextStore.get(sheetName);
    const rows = context?.recordRows;
    if (!rows?.size) {
      return [];
    }
    return (records || []).flatMap(record => {
      const recordNumber = this.getWorkbookRecordNumber(record);
      const excelRow = recordNumber ? rows.get(recordNumber) : undefined;
      if (!recordNumber || excelRow === undefined) {
        return [];
      }
      return [{
        ...record,
        controlData: { ...(record.controlData || {}), excelRow }
      } as MdoRecord];
    });
  }

  private async buildJoinedWorkbookRecords(contextStore: SheetContextStore, requestType: RequestType): Promise<MdoRecord[]> {
    const joined = new Map<string, MdoRecord>();
    for (const sheetName of contextStore.getAllSheetNames()) {
      const context = contextStore.get(sheetName);
      if (!context || !contextStore.restoreSheetContext(sheetName, this.sharedService)) {
        continue;
      }
      await this.sharedService.updateStatusForMissingRecordNumbers(sheetName);
      const sheetRecords = await this.excelReader.buildRecordList(requestType, sheetName) as MdoRecord[] | null;
      context.recordRows = new Map<string, number>();
      for (const record of sheetRecords || []) {
        const recordNumber = this.getWorkbookRecordNumber(record);
        if (!recordNumber) {
          continue;
        }
        const excelRow = Number((record.controlData as any)?.excelRow);
        if (Number.isFinite(excelRow)) {
          context.recordRows.set(recordNumber, excelRow);
        }
        const current = joined.get(recordNumber);
        joined.set(recordNumber, current
          ? this.mergeWorkbookRecord(current, record, context.sheetRole)
          : this.cloneWorkbookRecord(record));
      }
      contextStore.snapshotContext(sheetName, this.sharedService);
    }
    return [...joined.values()];
  }

  private getWorkbookRecordNumber(record: any): string {
    const controlData = record?.controlData || {};
    return String(controlData.recordNumber || controlData.arbitraryRecordNo || '').trim();
  }

  private cloneWorkbookRecord(record: MdoRecord): MdoRecord {
    const es: any = (record as any).mdoRecordES || {};
    return {
      ...(record as any),
      childRecord: [...((record as any).childRecord || [])],
      controlData: { ...((record as any).controlData || {}) },
      mdoRecordES: {
        ...es,
        hdvs: { ...(es.hdvs || {}) },
        hyvs: { ...(es.hyvs || {}) },
        gvs: { ...(es.gvs || {}) }
      }
    } as MdoRecord;
  }

  private mergeWorkbookRecord(target: MdoRecord, incoming: MdoRecord, role: SheetContext['sheetRole']): MdoRecord {
    const targetEs: any = (target as any).mdoRecordES || {};
    const incomingEs: any = (incoming as any).mdoRecordES || {};
    // Header/record sheets own header values. Grid/hierarchy sheets must not
    // overwrite them with their own blank or technical header values.
    if (role === 'header' || role === 'record') {
      targetEs.hdvs = { ...(targetEs.hdvs || {}), ...(incomingEs.hdvs || {}) };
    }
    targetEs.gvs = this.mergeStructuredValues(targetEs.gvs || {}, incomingEs.gvs || {});
    targetEs.hyvs = this.mergeStructuredValues(targetEs.hyvs || {}, incomingEs.hyvs || {});
    targetEs.descriptions = targetEs.descriptions || incomingEs.descriptions;
    (target as any).mdoRecordES = targetEs;
    return target;
  }

  private mergeStructuredValues(current: Record<string, any>, incoming: Record<string, any>): Record<string, any> {
    const merged = { ...current };
    for (const [key, value] of Object.entries(incoming || {})) {
      if (!merged[key]) {
        merged[key] = value;
        continue;
      }
      const existingRows = Array.isArray(merged[key]?.rows) ? merged[key].rows : [];
      const incomingRows = Array.isArray((value as any)?.rows) ? (value as any).rows : [];
      merged[key] = { ...merged[key], ...(value as any), rows: [...existingRows, ...incomingRows] };
    }
    return merged;
  }

  async highlightCells(validationErrors: any) {
    if (!validationErrors) return;
    const toUse = validationErrors;
    const addresses: string[] = [];
    for (const errorKey of Object.keys(toUse)) {
      const errorData = toUse[errorKey];
      const formKey = Object.keys(errorData)[0];
      const formErrors = errorData[formKey];
      const errs = formErrors?.formsValidationError?.errors as any[] | undefined;
      if (!errs?.length) continue;
      for (const err of errs) {
        const excelRows: string[] = err.excelRowNos || [];
        for (const rowNoStr of excelRows) {
          const rowNo = Number.parseInt(rowNoStr);
          if (!Number.isFinite(rowNo)) continue;
          // Highlight A..B for that row (example kept from earlier logic)
          addresses.push(`A${rowNo}:B${rowNo}`);
        }
      }
    }
    if (!addresses.length) return;
    await Excel.run({ delayForCellEdit: true }, async context => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const ranges = sheet.getRanges(addresses.join(','));
      ranges.format.fill.color = 'red';
      await context.sync();
    });
  }

  // Walk any nested validation response and collect field identifiers
  private extractErrorFieldKeys(validationResponse: any): string[] {
    const results = new Set<string>();
    const visit = (node: any) => {
      if (!node) return;
      if (Array.isArray(node)) {
        for (const it of node) visit(it);
        return;
      }
      if (typeof node === 'object') {
        const maybeId = (node.fieldId ?? node.fieldDesc ?? '').toString().trim();
        const hasError = node.error !== undefined || node.errors !== undefined;
        if (maybeId && hasError) results.add(maybeId);
        // dive deeper
        for (const k of Object.keys(node)) visit(node[k]);
      }
    };
    visit(validationResponse);
    return Array.from(results);
  }

  async generateDescription(aggregatedStartRow: number, aggregatedEndRow: number, context: Excel.RequestContext | null, isFromDaxeXlUI: boolean) {
    if (!this.sharedService.classificationEnabled) {
      return;
    }

    try {
      this.sharedService.isProcessRunning = true;
      await this.utilsService.displayStatus(this.translationService.translate('Generating descriptions...'));
      await this.sharedService.updateStatusForMissingRecordNumbers();
      let recordList: MdoRecord[] = [];

      if (isFromDaxeXlUI) {
        recordList = await this.excelReader.buildRecordList('generateDescriptions' as any) as MdoRecord[];
      } else {
        if (!context) {
          throw new Error('Excel context is required when not calling from DAXE XL UI');
        }
        recordList = await this.excelReader.buildRecordsForRowRange(aggregatedStartRow, aggregatedEndRow, 'generateDescriptions', context);
      }
      console.log('Prepared MDO records for aggregated set:', recordList.length, 'Rows:', aggregatedStartRow, aggregatedEndRow);
      if (recordList.length > 0) {
        const response = await this.sharedService.generateDescriptionsForRecords(recordList);
        await this.excelWriter.processGeneratedDescription(response);
      }
      await this.utilsService.displayStatus(null);
    } catch (error_) {
      console.error('Failed to prepare/emit MDO records for aggregated rows:', error_);
    } finally {
      this.sharedService.isProcessRunning = false;
      await this.utilsService.displayStatus(null);
    }
  }


  async applyDataRules(contextStore?: SheetContextStore) {
    if (!contextStore) {
      await this.sharedService.updateStatusForMissingRecordNumbers();
      const recordList = await this.excelReader.buildRecordList(RequestType.Defaults) as MdoRecord[];
      if(recordList.length === 0){
        this.sharedService.isProcessRunning = false;
        return;
      }
      daxeXlEvent.next({
        type: DaxeXLBaseEvents.DAXEXLApplyDataRulesRequest, payload: { recordList }
      });
      return;
    }

    const recordList: MdoRecord[] = [];
    for (const sheetName of contextStore.getAllSheetNames()) {
      if (!contextStore.restoreSheetContext(sheetName, this.sharedService)) {
        continue;
      }
      await this.sharedService.updateStatusForMissingRecordNumbers(sheetName);
      const rootRecords = await this.excelReader.buildRecordList(RequestType.Defaults, sheetName) as MdoRecord[] | null;
      if (rootRecords?.length) {
        recordList.push(...rootRecords);
      }
      contextStore.snapshotContext(sheetName, this.sharedService);
    }
    if(recordList.length === 0){
      this.sharedService.isProcessRunning = false;
      return;
    }
    daxeXlEvent.next({
      type: DaxeXLBaseEvents.DAXEXLApplyDataRulesRequest, payload: { recordList }
    });
  }

  async sendForCorrection() {
    const recordList = await this.excelReader.PrepareRecordsToSendForCorrection();
    daxeXlEvent.next({
      type: DaxeXLBaseEvents.DAXELSendRecordsForCorrectionRequest, payload: { recordList, step: 1 }
    })
  }
}
