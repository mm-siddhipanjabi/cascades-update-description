import { Injectable } from '@angular/core';
import { Field, ClassAttributes, ClassificationDescription, DaxeXLBaseEvents, DaxeXlEventType, daxeXlEvent, MdoRecord, ControlData, MdoRecordES, LayoutDetail } from 'daxe-ui-sdk';
import { groupBgColors } from '../daxe-element/daxe-xl/daxe-header-data';
import { catchError, filter, firstValueFrom, of, Subject, take, timeout } from 'rxjs';
import { TranslationService } from './translation.service';
import { Md5 } from 'ts-md5';

@Injectable({
  providedIn: 'root'
})
export class SharedService {
  fieldsMap: Map<string, Field> = new Map<string, Field>();
  staticFieldsObject: any[] = [{}];
  extentionHeaderFields: string[] = [];
  globalConfig: any;
  classAttributesMap: Map<string, ClassAttributes> = new Map<string, ClassAttributes>();
  /**
   * Characteristics created after a user selects a Description Class do not
   * exist in the form layout. Keep their metadata by sheet/cell so selection
   * of a list-type value can still use ngx-mdo's normal characteristic
   * dropdown rather than falling back to a plain Excel cell.
   */
  private readonly runtimeClassificationAttributes = new Map<string, any>();
  /**
   * Excel can deliver change notifications for an add-in write after
   * `runtime.enableEvents` is enabled again. Keep the exact generated range
   * briefly so the worksheet handler can discard those delayed notifications
   * without treating them as a user edit.
   */
  /** One generated worksheet for a real form CLASS_TYPE field, if configured. */
  dynamicClassificationSheet?: { sheetName: string; classFieldId: string; classType?: string; fields: any[] };
  private readonly dynamicClassAttributesByRow = new Map<string, any[]>();
  classDetails!: Map<string, ClassificationDescription>;
  language: string = "en";
  activeCellRow: number | null = 0;
  event: Subject<DaxeXlEventType> = daxeXlEvent;
  private _controlDataMap: Map<string, ControlData> = new Map<string, ControlData>();
  records: MdoRecord[] = [];
  recordsImported: boolean = false;
  /**
   * Complete classification state per logical record. The visible worksheet
   * contains only the fields configured for this element; this store preserves
   * imported/generated class metadata and attributes that are not visible in
   * the template so validate and submit never send a partial descriptions
   * object back to Core CRUD.
   */
  private readonly descriptionsByRecord = new Map<string, ClassificationDescription[]>();
  /**
   * Legacy one-description lookup used by the existing writer. New code must
   * update it through `mergeDescriptions` so it remains a view of the complete
   * per-record state rather than the source of truth.
   */
  generatedDescriptions: Map<string, ClassificationDescription | undefined> = new Map<string, ClassificationDescription | undefined>();
  currentSheetName: string = "";
  isExtentionProcess: boolean = false;
  /** Set only when the selected flow is an XFUSE MOCR workbook. */
  isMocrRequest: boolean = false;
  isProcessRunning: boolean = false;
  updatedDescHash: string | undefined;
  HDVSStringHash: string | undefined;

  codeFields: string[] = [];
  controlData = {
    moduleId: "",
    massId: "",
    layoutId: '',
    tenantId: "",
    userId: "",
    validationRequired: true
  }

  data = {
    controlData: this.controlData,
    flowId: '',
    formId: '',
    formDesc: '',
    stepId: '',
    schemaId: '',
    childFormId: '',
    classificationEnabled: false,
    businessRules: [] as any[],
    baseStructure: [] as any[],
    layoutDetails: [] as LayoutDetail[],
    uniqueStructures: [] as string[],
  }

  constructor(private readonly translationService: TranslationService) { }

  rememberRuntimeClassificationAttribute(sheetId: string, rowIndex: number, columnIndex: number, attribute: any): void {
    if (!sheetId || rowIndex < 0 || columnIndex < 0 || !attribute?.uuid) return;
    this.runtimeClassificationAttributes.set(`${sheetId}:${rowIndex}:${columnIndex}`, attribute);
  }

  getRuntimeClassificationAttribute(sheetId: string, rowIndex: number, columnIndex: number): any | undefined {
    return this.runtimeClassificationAttributes.get(`${sheetId}:${rowIndex}:${columnIndex}`);
  }

  clearRuntimeClassificationAttributes(sheetId: string, rowIndex: number, fromColumn = 0): void {
    const prefix = `${sheetId}:${rowIndex}:`;
    for (const key of this.runtimeClassificationAttributes.keys()) {
      if (!key.startsWith(prefix)) continue;
      const column = Number(key.slice(prefix.length));
      if (Number.isFinite(column) && column >= fromColumn) {
        this.runtimeClassificationAttributes.delete(key);
      }
    }
  }

  rememberDynamicClassAttributes(sheetId: string, rowIndex: number, attributes: any[]): void {
    if (!sheetId || rowIndex < 0) return;
    this.dynamicClassAttributesByRow.set(`${sheetId}:${rowIndex}`, attributes || []);
  }

  getDynamicClassAttributes(sheetId: string, rowIndex: number): any[] {
    return this.dynamicClassAttributesByRow.get(`${sheetId}:${rowIndex}`) || [];
  }


  /** Returns a defensive copy of the full classification payload for a row. */
  getDescriptions(recordKey: string, importedDescriptions: ClassificationDescription[] = []): ClassificationDescription[] {
    this.checkContextChange();
    const key = this.resolveDescriptionRecordKey(recordKey);
    if (!key) return this.cloneDescriptions(importedDescriptions);
    if (!this.descriptionsByRecord.has(key)) {
      this.descriptionsByRecord.set(key, this.cloneDescriptions(importedDescriptions));
    }
    return this.cloneDescriptions(this.descriptionsByRecord.get(key) || []);
  }

  /**
   * Merges data returned by description/class services with imported state.
   * A class is identified by its durable ID where available; class code/type is
   * retained as a fallback for legacy payloads. Attribute maps are merged per
   * language so server-only fields survive later workbook requests.
  */
  mergeDescriptions(recordKey: string, descriptions: ClassificationDescription[] | ClassificationDescription | undefined): ClassificationDescription[] {
    this.checkContextChange();
    const key = this.resolveDescriptionRecordKey(recordKey);
    if (!key || !descriptions) return this.getDescriptions(key);

    const incoming = Array.isArray(descriptions) ? descriptions : [descriptions];
    const current = this.getDescriptions(key);
    for (const description of incoming.filter(Boolean)) {
      const index = current.findIndex(item => this.sameDescriptionClass(item, description));
      if (index < 0) {
        current.push(this.cloneDescription(description));
      } else {
        current[index] = this.mergeDescription(current[index], description);
      }
    }
    const mergedDescriptions = this.cloneDescriptions(current);
    this.descriptionsByRecord.set(key, mergedDescriptions);
    const generated = current.find(description => description?.isDesc);
    if (generated) this.generatedDescriptions.set(key, this.cloneDescription(generated));
    // Imported records already contain the full description object and are
    // persisted for refresh. Keep that source of truth updated instead of
    // writing a second, independent description cache to local storage.
    this.updateImportedRecordDescriptions(key, mergedDescriptions);
    return this.cloneDescriptions(mergedDescriptions);
  }

  /** Adds the configured class selected in Excel without dropping other classes. */
  mergeWorkbookClass(recordKey: string, description?: ClassificationDescription): ClassificationDescription[] {
    return description ? this.mergeDescriptions(recordKey, description) : this.getDescriptions(recordKey);
  }

  /** Stops a new description request from using an obsolete generated result. */
  clearGeneratedDescription(recordKey: string): void {
    this.checkContextChange();
    this.generatedDescriptions.delete(this.resolveDescriptionRecordKey(recordKey));
  }

  /**
   * A blank Description Class cell is an explicit user removal. Keep static
   * classifications, but remove the generated/imported description class and
   * its attributes so neither validation nor Generate Description can reuse
   * stale values from memory or the persisted imported record.
   */
  clearDescriptionClass(recordKey: string): void {
    this.checkContextChange();
    const key = this.resolveDescriptionRecordKey(recordKey);
    if (!key) return;
    const remaining = this.getDescriptions(key).filter(description => !description?.isDesc);
    this.descriptionsByRecord.set(key, this.cloneDescriptions(remaining));
    this.generatedDescriptions.delete(key);
    this.updateImportedRecordDescriptions(key, remaining);
  }

  /**
   * Writer responses can identify a row as `${recordNumber}_${rowIndex}` while
   * reader requests use the logical record number. Reuse the existing logical
   * key so the response enriches the same description state instead of creating
   * a second, disconnected entry.
   */
  private resolveDescriptionRecordKey(recordKey: string): string {
    const rawKey = String(recordKey || '');
    if (!rawKey || this.descriptionsByRecord.has(rawKey)) return rawKey;
    return Array.from(this.descriptionsByRecord.keys())
      .find(existingKey => rawKey.startsWith(`${existingKey}_`)) || rawKey;
  }

  private sameDescriptionClass(left: ClassificationDescription, right: ClassificationDescription): boolean {
    if (left?.classId && right?.classId) return left.classId === right.classId;
    // Description fields do not guarantee CLASS_TYPE. A class modifier is part
    // of the business identity: ABRASIVE and ABRASIVE--BELT are different
    // classes even when both records have an empty classType.
    return this.getDescriptionClassIdentity(left) === this.getDescriptionClassIdentity(right);
  }

  private getDescriptionClassIdentity(description: ClassificationDescription | undefined): string {
    const rawCode = String(description?.classCode ?? '').trim();
    const separator = rawCode.indexOf('--');
    const classCode = separator >= 0 ? rawCode.slice(0, separator).trim() : rawCode;
    // Older imports can carry the mode only in classCode (`code--mode`), while
    // newer payloads use the separate classMode property. Treat both equally.
    const classMode = String(description?.classMode ?? (separator >= 0 ? rawCode.slice(separator + 2) : '')).trim();
    return `${classCode.toLocaleLowerCase()}\u0000${classMode.toLocaleLowerCase()}`;
  }

  private mergeDescription(current: ClassificationDescription, incoming: ClassificationDescription): ClassificationDescription {
    const merged: any = { ...this.cloneDescription(current), ...this.cloneDescription(incoming), attributes: {} };
    const languages = new Set([...Object.keys(current?.attributes || {}), ...Object.keys(incoming?.attributes || {})]);
    for (const language of languages) {
      const currentLanguage: any = current?.attributes?.[language] || {};
      const incomingLanguage: any = incoming?.attributes?.[language] || {};
      merged.attributes[language] = {
        ...currentLanguage,
        ...incomingLanguage,
        attrs: { ...(currentLanguage.attrs || {}), ...(incomingLanguage.attrs || {}) }
      };
    }
    return merged as ClassificationDescription;
  }

  private cloneDescriptions(descriptions: ClassificationDescription[] = []): ClassificationDescription[] {
    return descriptions.filter(Boolean).map(description => this.cloneDescription(description));
  }

  private cloneDescription(description: ClassificationDescription): ClassificationDescription {
    return JSON.parse(JSON.stringify(description)) as ClassificationDescription;
  }

  /** Hydrates the fast lookup map from the descriptions already persisted with imported records. */
  private rebuildDescriptionsFromImportedRecords(): void {
    this.descriptionsByRecord.clear();
    this.generatedDescriptions.clear();
    for (const record of this.records || []) {
      const recordKey = String((record as any)?.controlData?.recordNumber ?? (record as any)?.mdoRecordES?.id ?? '');
      const descriptions = (record as any)?.mdoRecordES?.descriptions;
      if (!recordKey || !Array.isArray(descriptions)) continue;
      const cloned = this.cloneDescriptions(descriptions);
      this.descriptionsByRecord.set(recordKey, cloned);
      const generated = cloned.find(description => description?.isDesc);
      if (generated) this.generatedDescriptions.set(recordKey, this.cloneDescription(generated));
    }
  }

  private updateImportedRecordDescriptions(recordKey: string, descriptions: ClassificationDescription[]): void {
    const record = (this.records || []).find((item: any) => {
      const key = String(item?.controlData?.recordNumber ?? item?.mdoRecordES?.id ?? '');
      return key === recordKey;
    });
    if (!record?.mdoRecordES) return;
    record.mdoRecordES.descriptions = this.cloneDescriptions(descriptions);
    this.persistImportedRecords();
  }

  // Load persisted records (if any) for the current step/mass context
  private readonly storagePrefix = 'DAXE_records_';
  private readonly controlDataStoragePrefix = 'DAXE_controlData_';
  private lastStoredStepId: string | null = null;
  private lastStoredMassId: string | null = null;
  private lastStoredModuleId: string | null = null;
  private lastStoredFormId: string | null = null;

  private getStorageKey(): string {
    const stepId = String(((this.data as any)?.stepId) ?? '');
    const massId = String(((this.data as any)?.controlData?.massId) ?? '');
    return `${this.storagePrefix}${stepId}_${massId}`;
  }

  private getControlDataStorageKey(): string {
    const stepId = String(((this.data as any)?.stepId) ?? '');
    const massId = String(((this.data as any)?.controlData?.massId) ?? '');
    return `${this.controlDataStoragePrefix}${stepId}_${massId}`;
  }

  private loadRecordsFromStorage(): void {
    try {
      const key = this.getStorageKey();
      if (!key) return;
      const raw = localStorage.getItem(key);
      if (!raw) return;
      const parsed = JSON.parse(raw) as MdoRecord[];
      if (Array.isArray(parsed)) {
        this.records = parsed;
        this.recordsImported = true;
        this.rebuildDescriptionsFromImportedRecords();
      }
    } catch (e) {
      console.warn('Failed to load persisted records from storage:', e);
    }
  }

  private loadControlDataFromStorage(): void {
    try {
      const key = this.getControlDataStorageKey();
      if (!key) return;
      const raw = localStorage.getItem(key);
      if (!raw) return;
      const parsed = JSON.parse(raw) as Record<string, ControlData>;
      if (parsed && typeof parsed === 'object') {
        this._controlDataMap.clear();
        for (const [mapKey, value] of Object.entries(parsed)) {
          this._controlDataMap.set(mapKey, value);
        }
      }
    } catch (e) {
      console.warn('Failed to load control data from storage:', e);
    }
  }

  private persistControlDataToStorage(): void {
    try {
      const key = this.getControlDataStorageKey();
      if (!key) return;
      const mapData: Record<string, ControlData> = {};
      for (const [mapKey, value] of this._controlDataMap) {
        mapData[mapKey] = value;
      }
      localStorage.setItem(key, JSON.stringify(mapData));
    } catch (e) {
      console.warn('Failed to persist control data to storage:', e);
    }
  }

  private checkContextChange(): void {
    const currentStepId = String(((this.data as any)?.stepId) ?? '');
    const currentMassId = String(((this.data as any)?.controlData?.massId) ?? '');
    const currentModuleId = String(((this.data as any)?.controlData?.moduleId) ?? '');
    const currentFormId = String(((this.data as any)?.formId) ?? '');

    if (this.lastStoredStepId !== currentStepId || this.lastStoredMassId !== currentMassId
      || this.lastStoredModuleId !== currentModuleId || this.lastStoredFormId !== currentFormId) {
      this.lastStoredStepId = currentStepId;
      this.lastStoredMassId = currentMassId;
      this.lastStoredModuleId = currentModuleId;
      this.lastStoredFormId = currentFormId;
      this.loadControlDataFromStorage();
      this.loadRecordsFromStorage();
    }
  }

  private cleanupOldRecords(): void {
    try {
      const currentKey = this.getStorageKey();
      const currentControlKey = this.getControlDataStorageKey();
      const keysToRemove: string[] = [];

      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (!k) continue;
        // Remove old record keys and control data keys that don't match current context
        const isOldRecordKey = k.startsWith(this.storagePrefix) && k !== currentKey;
        const isOldControlKey = k.startsWith(this.controlDataStoragePrefix) && k !== currentControlKey;
        // Description data is now stored inside DAXE_records_ only. Remove
        // obsolete independent caches created by older generated projects.
        const isLegacyDescriptionsKey = k.startsWith('DAXE_descriptions_');
        if (isOldRecordKey || isOldControlKey || isLegacyDescriptionsKey) {
          keysToRemove.push(k);
        }
      }

      for (const k of keysToRemove) {
        localStorage.removeItem(k);
        console.log(`Cleaned up old storage key: ${k}`);
      }
    } catch (e) {
      console.warn('Failed to cleanup old records:', e);
    }
  }

  private cleanupOtherStoredContexts(): void {
    this.cleanupOldRecords();
  }

  set attributesMap(data: { classCode: string, attribiute: ClassAttributes }) {
    if (!this.classAttributesMap.get(data.classCode)) {
      this.classAttributesMap.set(data.classCode, data.attribiute);
    }
  }

  /**
   * Cache a class returned from the picker/characteristics API under all
   * durable identifiers used by the generated workbook. A Description Class
   * column is template-only and does not always have a fixed class type in
   * form metadata, so the picker response is the authoritative source.
   */
  cacheClassAttributes(classAttributes: ClassAttributes | undefined): void {
    if (!classAttributes) return;
    const classData: any = classAttributes;
    const classType = String(classData.classType ?? '');
    if (classData.code) this.classAttributesMap.set(`${classData.code}${classType}`, classAttributes);
    if (classData.codeLong) this.classAttributesMap.set(`${classData.codeLong}${classType}`, classAttributes);
    // Class dropdowns render a modifier as `code--mod`; retain that display
    // key too so pasted and dropdown-selected classes resolve identically.
    if (classData.code && classData.mod) {
      this.classAttributesMap.set(`${classData.code}--${classData.mod}${classType}`, classAttributes);
    }
    if (classData.uuid) this.classAttributesMap.set(String(classData.uuid), classAttributes);
  }

  /**
   * Resolves the exact configured type when present. For a synthetic
   * Description Class column, use the picker-cached class only when its code
   * is unambiguous; this prevents a static class with another type from being
   * selected accidentally.
   */
  findClassAttributes(classCode: unknown, classType?: unknown): ClassAttributes | undefined {
    const code = String(classCode ?? '').trim();
    const type = String(classType ?? '').trim();
    if (!code) return undefined;
    // A template Description Class is synthetic: it has no fixed CLASS_TYPE,
    // while the host caches the chosen class under its exact Excel value.
    // Resolve that direct key before trying type-based/ambiguous matching so
    // Generate Description receives the same class shown in the worksheet.
    const direct = this.classAttributesMap.get(code);
    if (direct) return direct;
    if (type) {
      const exact = this.classAttributesMap.get(`${code}${type}`);
      if (exact) return exact;
    }
    const matches = Array.from(new Set(this.classAttributesMap.values())).filter((item: any) => {
      const classCode = String(item?.code ?? '').trim().toLowerCase();
      const classCodeLong = String(item?.codeLong ?? '').trim().toLowerCase();
      const displayCode = item?.code && item?.mod ? `${item.code}--${item.mod}`.toLowerCase() : '';
      const requestedCode = code.toLowerCase();
      return classCode === requestedCode || classCodeLong === requestedCode || displayCode === requestedCode;
    });
    return matches.length === 1 ? matches[0] : undefined;
  }

  /**
   * Get all control data entries. Checks context changes and loads from storage if needed.
   */
  getAllControlData(): Map<string, ControlData> {
    this.checkContextChange();
    return this._controlDataMap;
  }

  /**
   * Clear all control data and storage.
   */
  clearControlData(): void {
    this._controlDataMap.clear();
    try {
      localStorage.removeItem(this.getControlDataStorageKey());
    } catch (e) {
      console.warn('Failed to clear control data from storage:', e);
    }
  }

  /**
   * Getter for controlDataMap that handles context changes and persistence.
   * Use like: this.sharedService.controlDataMap.get(key) or this.sharedService.controlDataMap.set(key, value)
   * Automatically detects context changes and loads from storage.
   */
  get controlDataMap(): Map<string, ControlData> {
    this.checkContextChange();
    return this._controlDataMap;
  }

  /**
   * Setter for controlDataMap to allow direct assignment and automatic persistence.
   */
  set controlDataMap(value: Map<string, ControlData>) {
    this.checkContextChange();
    this._controlDataMap = value;
    this.persistControlDataToStorage();
  }

  parseRow(address: string): number | null {
    const cell = address.includes("!") ? address.split("!")[1] : address;
    const match = /^([A-Z]+)(\d+)$/.exec(cell);
    if (!match) {
      return null;
    }
    const rowStr = match[2];
    return Number(rowStr);
  }
  getGroupColor(groupName: string) {
    if (!groupName) {
      return "#ffffff";
    }
    const key = groupName.trim();
    return groupBgColors[key] || "#99ffff";
  }
  async updateAttributeMapBasedOnClassId(uniqueClassIds: string[]): Promise<void> {
    for (const classId of uniqueClassIds) {

      if (classId == null || classId.trim() === '') {
        console.log('Skipping empty classId');
        continue;
      }
      let response: any = {};
      await new Promise<void>((resolve) => {
        this.event.pipe(
          filter(evt => evt.type === DaxeXLBaseEvents.XFUSEClassAttributesResponse),
          timeout(5000),
          take(1),
          catchError(() => {
            console.error('ClassAttributes response error for classId:', classId);
            return of(null);
          }) // Prevent unhandled errors
        ).subscribe((evt) => {
          response = evt;
          resolve();
        });
        this.event.next({
          type: DaxeXLBaseEvents.DAXEXLClassAttributesRequest,
          payload: { classId, rowIndex: 0 }
        });

      });

      console.log('Received XFUSEClassAttributesResponse:', response);
      // If no response (timeout or error), exit silently or throw?
      if (!response?.data) return;
      const respData = response.data as { response: ClassAttributes, rowIndex: number };
      const classAttr = respData.response;
      if (classAttr) {
        this.cacheClassAttributes(classAttr);
      }
    }
  }
  async updateAttributeMap(uniqueClassCodes: {classCode: string, classType: string}[]): Promise<void> {
    let response: any = {};
    await new Promise<void>((resolve) => {
      this.event.pipe(
        filter(evt => evt.type === DaxeXLBaseEvents.XFUSEClassDetailsResponse),
        timeout(10000),
        take(1),
        catchError(err => {
          console.error('Dropdown response error:', err);
          return of(null);
        })
      ).subscribe((evt) => {
        response = evt;
        resolve();
      });

      this.event.next({
        // Opt into the generated-element class contract. Older elements omit
        // this marker, so ngx-mdo preserves their original lookup behavior.
        type: DaxeXLBaseEvents.DAXEXLClassDetailsRequest, payload: {
          moduleId: this.data.controlData.moduleId,
          classCodeList: uniqueClassCodes,
          classificationV2: true
        } as any
      });
    });

    if (!response?.data) return;

    // Update map
    const newMap = response.data as Map<string, ClassAttributes>;
    for (const [key, value] of newMap) {
      // Preserve the host's exact request key (including `code--mod`) in
      // addition to the API's class fields. Some class-attribute responses do
      // not contain `mod`, but the worksheet value does.
      this.classAttributesMap.set(String(key), value);
      this.cacheClassAttributes(value);
    }
    console.log("Attribute map updated:", this.classAttributesMap);
  }

  private readonly codeCache = new Map<string, { code: string; text: string }>(); // cache for dropdown codes


  async getDropDownCodeOrText(data: { field: Field, fieldId: string, cellValue: string, formId: string, datsetId: string }, getText: boolean = false): Promise<string> {

    const cacheKey = JSON.stringify({ fieldId: data.fieldId, cellValue: data.cellValue, callType: getText ? 'text' : 'code' });

    // 1️⃣ If request already exists in cache → return from cache
    if (this.codeCache.has(cacheKey)) {
      const value = this.codeCache.get(cacheKey)?.[getText ? 'text' : 'code'] || '';
      console.log(`getDropDownCodeOrText: cache hit for fieldId=${data.fieldId}, cellValue=${data.cellValue},cacheValue='${value}'`);
      return value;
    }
    try {
      console.log(`getDropDownCodeOrText: cache(${cacheKey}) miss for fieldId=${data.fieldId}, cellValue=${data.cellValue}; firing request`);

      let response: any = {};
      await new Promise<void>((resolve) => {
        this.event.pipe(
          filter(evt => evt.type === DaxeXLBaseEvents.XFUSEDropdownCodeResponse),
          timeout(15000),
          take(1),
          catchError(err => {
            console.error('Dropdown response error:', err);
            return of(null);
          })
        ).subscribe((evt) => {
          response = evt;
          resolve();
        });
        // 2️⃣ Fire request event
        this.event.next({ type: DaxeXLBaseEvents.DAXEXLDropdownCodeRequest, payload: { ...data } });
      });


      // If no response, cache empty string
      if (!response?.data) {
        console.warn(`getDropDownCodeOrText: no response received for fieldId=${data.fieldId}`);
        return data.cellValue;
      }

      const key = getText ? 'code' : 'text';
      const item = (response as any)?.data?.content?.find((item: any) => item[key] === data.cellValue);
      console.log("item received for fieldId:", data.fieldId);
      console.log("response:", item);
      const value = getText ? item?.text : item?.code;

      console.log(`getDropDownCodeOrText: received value='${value}' for fieldId=${data.fieldId}, cellValue=${data.cellValue}`);

      if (value) {
        // 4️⃣ Cache the result
        this.codeCache.set(cacheKey, item);

        const cacheKeySecond = JSON.stringify({ fieldId: data.fieldId, cellValue: data.cellValue, callType: getText ? 'code' : 'text' });
        if (!this.codeCache.has(cacheKeySecond)) {
          this.codeCache.set(cacheKeySecond, item);
        }
      }
      return value ?? '';
    } catch (e) {
      console.error('Error in getDropDownCodeOrText:', e);
      return data.cellValue;
    }
  }

  /**
   * Generic function to send a request event and wait for a response event
   * @param requestType - The event type to send
   * @param responseType - The event type to listen for
   * @param payload - The payload to send with the request
   * @param filterCondition - Optional additional filter condition for the response (receives the event)
   * @param timeoutMs - Timeout in milliseconds (default: 5000)
   * @param errorMessage - Custom error message for logging (default: 'Request response error')
   * @returns Promise resolving to the response event data
   */
  async sendEventAndWaitForResponse<T = any>(
    requestType: DaxeXLBaseEvents,
    responseType: DaxeXLBaseEvents,
    payload: any,
    filterCondition?: (evt: any) => boolean,
    timeoutMs: number = 5000,
    errorMessage: string = 'Request response error'
  ): Promise<T | null> {
    const response$ = daxeXlEvent.pipe(
      filter(evt => {
        const typeMatches = evt.type === responseType;
        return typeMatches && (filterCondition ? filterCondition(evt) : true);
      }),
      timeout(timeoutMs),
      take(1),
      catchError(err => {
        console.error(`${errorMessage}:`, err);
        return of(null);
      })
    );

    daxeXlEvent.next({ type: requestType, payload } as DaxeXlEventType);

    const evt: any = await firstValueFrom(response$);
    return evt?.data ?? null;
  }

  async generateDescriptionsForRecords(recordList: MdoRecord[]): Promise<ClassificationDescription[]> {
    if (!this.classificationEnabled) {
      return [];
    }

    const mdoRecordESList = recordList.map(r => r.mdoRecordES);
    const descriptionData = mdoRecordESList.filter(item =>
      Array.isArray(item.descriptions) &&
      item.descriptions.some(desc => desc)
    );
    let response: any = {};
    await new Promise<void>((resolve) => {
      daxeXlEvent.pipe(
        filter(evt => evt.type === DaxeXLBaseEvents.XFUSEGenerateDescriptionResponse),
        timeout(20000),
        take(1),
        catchError(err => {
          console.error('Generate description response error:', err);
          return of(null)
        })
      ).subscribe(async (evt) => {
        response = evt;

        try {
        if (response?.data) {
          const UpdatedDesc = descriptionData.flatMap(({ descriptions }: any) =>
            descriptions.map(({ classCode, attributes }: any) => ({ classCode, attributes }))
          );

          const UpdatedDescString = JSON.stringify(UpdatedDesc);
          this.updatedDescHash = Md5.hashStr(UpdatedDescString);
        }
        } catch (e) {
          console.error('Error computing description hashes:', e);
        }
        resolve();
      });

      daxeXlEvent.next({
        type: DaxeXLBaseEvents.DAXEXLGenerateDescriptionRequest, payload: {
          descriptionData,
          moduleId: this.data.controlData.moduleId,
          formRuleId: this.data.businessRules.find(m => m.ruleType == "BR_CLASSIFICATION")?.ruleId || "",
          schemaId: "",
        }
      });

    });

    return (response as any) as ClassificationDescription[] || [];
  }

  collectAllUniqueClassCodes(mdoRecord: MdoRecord[]) {
    const classCodeSet = new Set<string>();
    for (const record of mdoRecord) {
      const descriptions = record.mdoRecordES.descriptions ?? [];
      for (const desc of descriptions) {
        classCodeSet.add(desc.classCode);
      }
    }
    return Array.from(classCodeSet);
  }

  saveRecordsImported(records: MdoRecord[]) {
    this.records = records;
    this.rebuildDescriptionsFromImportedRecords();
    this.persistImportedRecords();
  }

  private persistImportedRecords(): void {
    try {
      const key = this.getStorageKey();
      if (key) {
        // remove other contexts so stored data stays consistent with current step/mass
        this.cleanupOtherStoredContexts();
        localStorage.setItem(key, JSON.stringify(this.records));
        this.recordsImported = true;
      }
    } catch (e) {
      console.warn('Failed to persist imported records to storage:', e);
    }
  }

  // Attempt to restore persisted records and control data when service initializes
  // (call at end of constructor)
  public initPersistence() {
    this.loadRecordsFromStorage();
    this.loadControlDataFromStorage();
    this.lastStoredStepId = String(((this.data as any)?.stepId) ?? '');
    this.lastStoredMassId = String(((this.data as any)?.controlData?.massId) ?? '');
    this.lastStoredModuleId = String(((this.data as any)?.controlData?.moduleId) ?? '');
    this.lastStoredFormId = String(((this.data as any)?.formId) ?? '');
    // Cleanup old storage from other contexts on initialization
    this.cleanupOldRecords();
  }

  get headerColumns() {
    return this.staticFieldsObject
      .filter(f => !f.hidden && f.columnNumber >= 0)
      .map(f => f.columnNumber);
  }

  /** Clears stale missing-number messages on the supplied sheet without activating it. */
  async updateStatusForMissingRecordNumbers(sheetName?: string){
    try {
      await Excel.run(async (context) => {
      const sheet = sheetName
        ? context.workbook.worksheets.getItem(sheetName)
        : context.workbook.worksheets.getActiveWorksheet();
      const startRow = 2;

      const usedRange = sheet.getUsedRange();
      usedRange.load('rowCount');
      await context.sync();

      const totalRows = usedRange.rowCount;
      const rowCount = Math.max(0, totalRows - startRow);

      const range = sheet.getRangeByIndexes(startRow, 0, rowCount, 2);

      range.load('values');
      await context.sync();

      if (rowCount === 0) {
        return;
      }

      const excelValues = range.values as string[][];
      const missingRecordNumberMessage = this.translationService.translate('RecordNumberMissing')?.toLowerCase()?.trim();

      for(const row of excelValues){
        const statusMessage = row[1];
         if (statusMessage?.toLowerCase()?.trim() === missingRecordNumberMessage) {
            row[1] = '';
            row[0] = '';
          }
      }
      
      range.values = excelValues;
      await context.sync();
    })
    }
    catch(error) {
      console.error("Error clearing columns: ", error);
    }
  }

  get readOnlyColumnsOnChange(): number[] {
    return this.staticFieldsObject
      .filter(f => f.fieldId === 'mdo_number' || f.hidden)
      .map(f => f.columnNumber);
  }

  get rowInsertReadonly(): number[] {
    return this.staticFieldsObject
      .filter(f => f.fieldId === 'mdo_number')
      .map(f => f.columnNumber);
  }

  get alwaysReadOnly(): number[] {
    return [];
  }

  get decimalFields() {
    return Array.from(this.fieldsMap.values()).filter(f => f.metadata?.[0]?.dataType == 'DEC')?.map((f: any) => f.fieldId);
  }

  get classificationEnabled() {
    return Boolean((this.data as any)?.classificationEnabled);
  }

  get readonlyFieldCols(): number[] {
    let readonlyCols: number[] = [];
    if (this.isExtentionProcess) {
      readonlyCols = Array.from(this.fieldsMap.values()).filter(f => f.isReadOnly)?.map((f: any) => f.columnNumber);
      readonlyCols.push(2);
    }
    return readonlyCols;
  }

  get profileLanguage() {
    return this.globalConfig?.UserProfile?.selfServiceUserModel?.lang
  }

  get allLanguages() {
    return this.globalConfig?.UserProfile?.lang || ['en']
  }

  createLock = () => {
    const queue: any[] = [];
    let active = false;
    return (fn: any) => {
      let deferredResolve: any;
      let deferredReject: any;
      const deferred = new Promise((resolve, reject) => {
        deferredResolve = resolve;
        deferredReject = reject;
      });
      const exec = async () => {
        await fn().then(deferredResolve, deferredReject);
        if (queue.length > 0) {
          queue.shift()();
        } else {
          active = false;
        }
      };
      if (active) {
        queue.push(exec);
      } else {
        active = true;
        exec();
      }
      return deferred;
    };
  };
  lock = this.createLock();
}
