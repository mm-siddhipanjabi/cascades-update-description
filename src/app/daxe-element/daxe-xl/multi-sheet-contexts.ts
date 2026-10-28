/**
 * Empty for the standard single-sheet element. Workbook-layout generation
 * replaces this with sheet contexts that are joined by Record Number.
 */
export const multiSheetContexts: ReadonlyArray<{
  sheetName: string;
  role: 'header' | 'grid' | 'hierarchy' | 'record';
  recordNumberFieldLabel: string;
  moduleId: string;
  layoutId: string;
  formDesc: string;
  stepId: string;
  fieldData: any[];
  codeFields: string[];
}> = [];
