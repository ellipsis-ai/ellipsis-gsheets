import { Client, EllipsisObjectWithEnvVars } from './client';
import { JWT } from 'google-auth-library';
import { google, sheets_v4 } from 'googleapis';
import { sheets } from 'googleapis/build/src/apis/sheets';

type SheetRow = Array<any>

interface SheetInfo {
  id: number | null
  name: string | null
  data?: Array<SheetRow> | null
}

export class Sheet {
  readonly sheets: sheets_v4.Sheets;
  readonly client: JWT;
  authorized: boolean;

  constructor(
    readonly ellipsis: EllipsisObjectWithEnvVars,
    readonly spreadsheetId: string,
    readonly overrideServiceAccountEmail?: string | null,
    readonly overridePrivateKey?: string | null
  ) {
    this.spreadsheetId = spreadsheetId;
    this.client = Client(ellipsis, overrideServiceAccountEmail, overridePrivateKey);
    this.sheets = google.sheets({
      version: "v4",
      auth: this.client
    });
    this.authorized = false;
  }

  private checkAuthAnd<T>(next: () => Promise<T>): Promise<T> {
    if (this.authorized) {
      return next();
    } else {
      return this.client.authorize().then(() => {
        this.authorized = true;
        return next();
      });
    }
  }

  append(range: string, rows: Array<SheetRow>): Promise<number | null> {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: rows
        }
      }).then((response) => response.data.updates && response.data.updates.updatedCells || null);
    });
  }

  update(range: string, rows: Array<SheetRow>): Promise<number | null> {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          range: range,
          values: rows,
        }
      }).then((response) => response.data.updatedCells || null);
    });
  }

  get(range: string): Promise<Array<SheetRow>> {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueRenderOption: 'FORMATTED_VALUE'
      }).then((response) => response.data.values || []);
    });
  }

  getAllSheets(options?: { includeData?: boolean }): Promise<Array<SheetInfo>> {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.get({
        spreadsheetId: this.spreadsheetId,
        ranges: [],
        includeGridData: options && options.includeData || false
      }).then((response) => {
        if (!response.data.sheets) {
          return [];
        } else {
          return response.data.sheets.map((sheet) => {
            const firstGrid = sheet.data && sheet.data[0] || null;
            return {
              id: sheet.properties && typeof sheet.properties.sheetId === "number" ? sheet.properties.sheetId : null,
              name: sheet.properties && typeof sheet.properties.title === "string" ? sheet.properties.title : null,
              data: firstGrid && firstGrid.rowData ? firstGrid.rowData.map((cellData) => {
                return cellData.values ? cellData.values.map((cellValue) => cellValue.formattedValue || null) : [];
              }) : null
            };
          });
        }
      });
    });
  }

  createSheet(name: string): Promise<SheetInfo> {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.batchUpdate({
        spreadsheetId: this.spreadsheetId,
        requestBody: {
          requests: [{
            addSheet: {
              properties: {
                title: name,
                gridProperties: {
                  frozenRowCount: 1
                }
              }
            }
          }]
        }
      }).then((response) => {
        const newSheetResponse = response.data.replies && response.data.replies[0] || null;
        const properties = newSheetResponse && newSheetResponse.addSheet && newSheetResponse.addSheet.properties || null;
        return {
          id: properties && properties.sheetId || null,
          name: properties && properties.title || null
        };
      });
    });
  }

}
