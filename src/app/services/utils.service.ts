import { Injectable, NgZone } from '@angular/core';
import moment from 'moment-timezone/moment-timezone-utils';

@Injectable({
  providedIn: 'root'
})
export class UtilsService {

  globalConfig: any;

  constructor(private zone: NgZone) { }
  /**
 * Gets the current timezone of the user.
 * @returns The user's timezone.
 */
  getTimeZone() {
    return this.globalConfig.UserProfile.selfServiceUserModel.timezone;
  }

  /**
   * Gets the date format from user preferences.
   * @returns The date format string.
   */
  getDateformat() {
    return this.globalConfig.UserProfile.selfServiceUserModel.dFormat || "DD-MM-YYYY";
  }

  /**
    * Gets the time format from user preferences.
    * @returns The time format string.
    */
  getTimeformat() {
    return this.globalConfig.UserProfile.selfServiceUserModel.tFormat == '12' ? 'hh:mm A' : 'HH:mm';
  }
  getDateFromMilliseconds(dateValue: any, pickList: string = "") {
    if (dateValue) {
      const dateFormats = this.globalConfig.UserProfile?.selfServiceUserModel?.dFormat ? this.globalConfig.UserProfile?.selfServiceUserModel?.dFormat : 'DD-MM-YYYY';
      const timeformat = this.globalConfig.UserProfile?.selfServiceUserModel?.tFormat ? this.globalConfig.UserProfile?.selfServiceUserModel?.tFormat : '24';
      let timeFormat = 'HH:mm:ss';
      if (timeformat === '12') {
        timeFormat = 'hh:mm:ss a'
      }

      let dateNumber = Number(dateValue);
      if (dateNumber) {
        let utc = moment(dateNumber).utc();
        const timezone = this.globalConfig.UserProfile?.selfServiceUserModel?.timezone ? this.globalConfig.UserProfile.selfServiceUserModel.timezone : (Intl.DateTimeFormat() as any).resolvedOptions().timeZone;
        let dateformated = '';
        if (pickList == "53")
          dateformated = moment(utc).tz(timezone).format(dateFormats);
        else if (pickList == "52")
          dateformated = moment(utc).format(dateFormats);
        return dateformated;
      }
      else
        return "";
    } else {
      return dateValue;
    }
  }
  getDateTimeFromMillies(millies: number, hours: number = 0, minutes: number = 0): string {
    if (!millies) { return ''; }
    let userTimezoneDate = moment(millies).tz(this.getTimeZone()).add(hours, 'hours').add(minutes, 'minutes');
    let dateValue = userTimezoneDate.format(this.getDateformat());
    let timeValue = userTimezoneDate.format(this.getTimeformat());
    return `'${dateValue} (${timeValue})`;
  }
  async toggleEvents(EnableEvents: boolean) {
    await Excel.run({ delayForCellEdit: true }, async (context: { workbook: { worksheets: { getActiveWorksheet: () => any; }; }; runtime: { load: (arg0: string) => void; enableEvents: boolean; }; sync: () => any; application: { suspendScreenUpdatingUntilNextSync: () => void; suspendApiCalculationUntilNextSync: () => void; }; }) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      context.runtime.load("enableEvents");
      sheet.protection.load("protected");
      await context.sync();
      context.runtime.enableEvents = EnableEvents;
      if (EnableEvents) {
        if (!sheet.protection.protected) {
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
        }
        console.log("Events are currently on.");
      } else {
        if (sheet.protection.protected) {
          sheet.protection.unprotect();
        }
        console.log("Events are currently off.");
        context.application.suspendScreenUpdatingUntilNextSync();
        context.application.suspendApiCalculationUntilNextSync();
      }
      await context.sync();
    }).catch(() => {
      console.error("Error toggling events");
    });
  }

  /**
   * Temporarily suppresses Office events without changing sheet protection.
   * MOCR header creation writes several non-active worksheets, so using the
   * normal toggleEvents method would otherwise protect the user's active sheet.
   */
  async toggleRuntimeEvents(enableEvents: boolean) {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: any) => {
        context.runtime.enableEvents = enableEvents;
        if (!enableEvents) {
          context.application.suspendScreenUpdatingUntilNextSync();
          context.application.suspendApiCalculationUntilNextSync();
        }
        await context.sync();
      });
    } catch (error) {
      console.error('Error toggling runtime events', error);
    }
  }

  /** Protects a named worksheet directly, without making it the active sheet. */
  async protectWorksheet(sheetName: string) {
    try {
      await Excel.run({ delayForCellEdit: true }, async (context: any) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        sheet.protection.load('protected');
        await context.sync();
        if (!sheet.protection.protected) {
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
        }
        await context.sync();
      });
    } catch (error) {
      console.error(`Error protecting worksheet ${sheetName}`, error);
    }
  }

  /**
   * Display a message at the bottom of the task pane.
   */
  async displayStatus(message: string | null, clearAfterMs?: number): Promise<void> {
    try {

      this.zone.run(async () => {
        // Create or update a notification element at the bottom of task pane
        let notifContainer = document.getElementById('status-notification-container');
        if (!notifContainer) {
          notifContainer = document.createElement('div');
          notifContainer.id = 'status-notification-container';
          notifContainer.style.cssText = `
          position: fixed;
          bottom: 0;
          left: 0;
          right: 0;
          background-color: #0078d4;
          color: white;
          padding: 12px 16px;
          font-size: 14px;
          box-shadow: 0 -2px 8px rgba(0,0,0,0.15);
          z-index: 10000;
          display: none;
          animation: slideUp 0.3s ease-in-out;
        `;

          // Add CSS animation
          if (!document.getElementById('status-notification-style')) {
            const style = document.createElement('style');
            style.id = 'status-notification-style';
            style.textContent = `
            @keyframes slideUp {
              from { transform: translateY(100%); opacity: 0; }
              to { transform: translateY(0); opacity: 1; }
            }
            @keyframes slideDown {
              from { transform: translateY(0); opacity: 1; }
              to { transform: translateY(100%); opacity: 0; }
            }
          `;
            document.head.appendChild(style);
          }

          document.body.appendChild(notifContainer);
        }

        if (message) {
          notifContainer.textContent = message;
          notifContainer.style.display = 'block';
          notifContainer.style.animation = 'slideUp 0.3s ease-in-out';
          console.log(`[Status] ${message}`);
        } else {
          notifContainer.style.animation = 'slideDown 0.3s ease-in-out';
          setTimeout(() => {
            if (notifContainer) {
              notifContainer.style.display = 'none';
            }
          }, 300);
        }

        if (message && clearAfterMs && clearAfterMs > 0) {
          await new Promise(resolve => {
            setTimeout(() => {
              if (notifContainer) {
                notifContainer.style.animation = 'slideDown 0.3s ease-in-out';
                setTimeout(() => {
                  if (notifContainer) {
                    notifContainer.style.display = 'none';
                  }
                }, 300);
              }
              console.log(`[Status] (cleared)`);
              resolve(undefined);
            }, clearAfterMs);
          });
        }
      });
    } catch (err) {
      console.warn('displayStatus failed:', err);
    }
  }
  /**
   * Builds a map of record number -> { first: rowIndex, last: rowIndex }
   * Row indexes are 1-based Excel row numbers (matching the existing +3 offset logic).
   * Assumes `recordNumbers[0]` corresponds to Excel row 3.
   */
  public getRecordIndexBoundsMap(recordNumbers: string[]): Map<string, { first: number; last: number }> {
    const bounds = new Map<string, { first: number; last: number }>();
    for (let idx = 0; idx < recordNumbers.length; idx++) {
      const recNo = recordNumbers[idx];
      if (!recNo) {
        console.log(`Skipping empty record number at index ${idx}`, recNo);
        continue;
      }
      const excelRow = idx + 3; // convert zero-based array index (starting at row 3) to Excel row number
      const existing = bounds.get(recNo);
      if (!existing) {
        bounds.set(recNo as string, { first: excelRow, last: excelRow });
      } else {
        existing.last = excelRow;
      }
    }
    return bounds;
  }

  async getAllRecordNumbers(ctx: Excel.RequestContext): Promise<string[]> {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load(['rowCount']);
    await ctx.sync();
    const totalRows = usedRange.rowCount;
    // Load full record number column once (C3..C{totalRows}) to compute last index map
    const recColStartRow = 2; // zero-based index for row 3
    const recColRowCount = Math.max(0, totalRows - recColStartRow);
    const recColRange = recColRowCount > 0 ? sheet.getRangeByIndexes(recColStartRow, 2, recColRowCount, 1) : null;
    if (recColRange) recColRange.load('values');
    await ctx.sync();

    const recordNumbers = recColRange ? (recColRange.values as any[]).map(r => r[0]) : [];
    return recordNumbers;
  }

  /**
 * To get milliseconds from date.
 * @param dateStr Date as string
 * @returns Milliseconds
 */
  convertDateToMilliseconds(dateStr: string) {

    if (dateStr) {
      let dateFormat = this.globalConfig.UserProfile?.selfServiceUserModel?.dFormat?.toUpperCase() || "DD-MM-YYYY";
      let milliseconds = moment(moment.utc(dateStr, dateFormat, true)).valueOf();
      return String(milliseconds);
    }
    else { return null; }
  }

  async isMDOConfigurationExist(): Promise<boolean> {
    let isMDOConfigurationExist: boolean = false;
    await Excel.run({ delayForCellEdit: true }, async (context: { workbook: { worksheets: any; }; sync: () => any; }) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/name");
      await context.sync();
      isMDOConfigurationExist = worksheets.items.map((sheet: { name: any; }) => sheet.name).includes("MDOConfiguration");
    });
    return isMDOConfigurationExist;
  }


  async displayDialog(
    title: string,
    message: string,
    options?: { okText?: string }
  ): Promise<void> {
    try {
      return new Promise<void>((resolve) => {
        this.zone.run(() => {

          // Prevent duplicate dialogs
          let existingDialog = document.getElementById('custom-dialog-overlay');
          if (existingDialog) {
            existingDialog.remove();
          }

          // Overlay
          const overlay = document.createElement('div');
          overlay.id = 'custom-dialog-overlay';
          overlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.4);
            z-index: 10001;
            display: flex;
            align-items: center;
            justify-content: center;
          `;

          // Dialog box
          const dialog = document.createElement('div');
          dialog.style.cssText = `
            background: #ffffff;
            width: 420px;
            max-width: 90%;
            border-radius: 8px;
            box-shadow: 0 8px 24px rgba(0,0,0,0.2);
            font-family: Arial, sans-serif;
            animation: dialogFadeIn 0.25s ease-out;
          `;

          dialog.innerHTML = `
            <div style="padding: 16px 20px; border-bottom: 1px solid #eee; font-size: 16px; font-weight: 600;">
              ${title}
            </div>
            <div style="padding: 20px; font-size: 14px; color: #333;">
              ${message}
            </div>
            <div style="padding: 12px 20px; text-align: right; border-top: 1px solid #eee;">
              <button id="dialog-ok-btn"
                style="
                  background: #0078d4;
                  color: white;
                  border: none;
                  padding: 6px 14px;
                  border-radius: 4px;
                  cursor: pointer;
                  font-size: 13px;">
                ${options?.okText ?? 'OK'}
              </button>
            </div>
          `;

          // Animation CSS (once)
          if (!document.getElementById('custom-dialog-style')) {
            const style = document.createElement('style');
            style.id = 'custom-dialog-style';
            style.textContent = `
              @keyframes dialogFadeIn {
                from { transform: scale(0.95); opacity: 0; }
                to { transform: scale(1); opacity: 1; }
              }
            `;
            document.head.appendChild(style);
          }

          overlay.appendChild(dialog);
          document.body.appendChild(overlay);

          // OK button handler
          dialog.querySelector('#dialog-ok-btn')?.addEventListener('click', () => {
            overlay.remove();
            resolve();
          });
        });
      });
    } catch (err) {
      console.warn('displayDialog failed:', err);
    }
  }

  public getUniqueStructureList(layoutDetails: any[]): { structure: string, sheetName: string }[] {
    const structureMap = new Map<string, { structure: string, sheetName: string }>();

    layoutDetails?.forEach(layoutDetail => {
      layoutDetail?.fields?.forEach((field: any) => {
        const metadata = field?.metadata?.[0];
        const structure = field?.metadata?.[0]?.grid ? field.fieldId : metadata?.structureId?.toString();

        if (!structure || structure === 'undefined' || structureMap.has(structure)) {
          return;
        }

        structureMap.set(structure, {
          structure,
          sheetName: metadata?.sheetName || metadata?.structDesc || layoutDetail?.tabText || ''
        });
      });
    });

    return Array.from(structureMap.values());
  }

}

export enum DateTime {
  DATE = "52",
  DATETIME = "53",
  TIME = "54"
}
export enum CellValueSaparator {
  classSeparator = ';',
  attachementSaparator = '|',
  normalDropdownSaparator = ','
}
