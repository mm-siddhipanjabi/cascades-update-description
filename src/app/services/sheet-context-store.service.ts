import { Injectable } from '@angular/core';
import { SharedService } from './shared.service';

export interface SheetContext {
  sheetName: string;
  /** A normal workbook sheet's contribution to one merged form record. */
  sheetRole?: 'header' | 'grid' | 'hierarchy' | 'record';
  /** Local row position by Record Number, used to return responses to this sheet. */
  recordRows?: Map<string, number>;
  flowId: string;
  stepId: string;
  formId: string;
  childFormId: string;
  schemaId: string;
  moduleId: string;
  layoutId: string;
  formDesc: string;
  massId: string;
  userId: string;
  tenantId: string;
  mocrId: string;
  fieldsMap: Map<string, any>;
  staticFieldsObject: any[];
  codeFields: string[];
  extensionHeaderFields: string[];
  businessRules: any[];
  baseStructure: any[];
}

@Injectable({ providedIn: 'root' })
export class SheetContextStore {
  private readonly contexts = new Map<string, SheetContext>();

  set(sheetName: string, ctx: SheetContext): void {
    this.contexts.set(sheetName, ctx);
  }

  getAllSheetNames(): string[] {
    return [...this.contexts.keys()];
  }

  get(sheetName: string): SheetContext | undefined {
    return this.contexts.get(sheetName);
  }

  /** Applies the shared request control data to every independent MOCR root. */
  updateControlData(controlData: Partial<Pick<SheetContext, 'massId' | 'userId' | 'tenantId' | 'mocrId'>>): void {
    for (const context of this.contexts.values()) {
      Object.assign(context, controlData);
    }
  }

  /**
   * Restores SharedService to the named sheet's context. This changes only
   * in-memory request metadata; it does not activate an Excel worksheet.
   * All existing services (ExcelEventHandler, ExcelReader, ExcelWriter, DaxeXLService)
   * read from `data` and `sharedService` — so they transparently pick up the right flow.
   */
  restoreSheetContext(sheetName: string, sharedService: SharedService): boolean {
    const ctx = this.contexts.get(sheetName);
    if (!ctx) return false;

    sharedService.data.flowId          = ctx.flowId;
    sharedService.data.stepId          = ctx.stepId;
    sharedService.data.formId          = ctx.formId;
    sharedService.data.formDesc        = ctx.formDesc;
    sharedService.data.childFormId     = ctx.childFormId;
    sharedService.data.schemaId        = ctx.schemaId;
    sharedService.data.businessRules   = ctx.businessRules;
    sharedService.data.baseStructure   = ctx.baseStructure;

    sharedService.data.controlData.moduleId  = ctx.moduleId;
    sharedService.data.controlData.layoutId  = ctx.layoutId;
    sharedService.data.controlData.massId    = ctx.massId;
    sharedService.data.controlData.userId    = ctx.userId;
    sharedService.data.controlData.tenantId  = ctx.tenantId;
    (sharedService.data.controlData as any).mocrId = ctx.mocrId;

    // Restore SharedService state
    sharedService.fieldsMap          = ctx.fieldsMap;
    sharedService.staticFieldsObject = ctx.staticFieldsObject;
    sharedService.codeFields         = ctx.codeFields;
    sharedService.extentionHeaderFields = ctx.extensionHeaderFields;

    return true;
  }

  /**
   * Saves the current SharedService context snapshot.
   * Call this before switching away so any runtime updates (businessRules, controlData) are preserved.
   */
  snapshotContext(sheetName: string, sharedService: SharedService): void {
    const ctx = this.contexts.get(sheetName);
    if (!ctx) return;

    ctx.fieldsMap          = sharedService.fieldsMap;
    ctx.staticFieldsObject = sharedService.staticFieldsObject;
    ctx.codeFields         = [...sharedService.codeFields];
    ctx.extensionHeaderFields = [...sharedService.extentionHeaderFields];
    ctx.businessRules      = sharedService.data.businessRules;
    ctx.massId             = sharedService.data.controlData.massId;
    ctx.userId             = sharedService.data.controlData.userId;
    ctx.tenantId           = sharedService.data.controlData.tenantId;
    ctx.mocrId             = (sharedService.data.controlData as any).mocrId || '';
  }
}
