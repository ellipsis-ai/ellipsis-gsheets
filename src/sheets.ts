import { Client, EllipsisObjectWithEnvVars } from './client';
import { JWT } from 'google-auth-library';
import { google, sheets_v4 } from 'googleapis';
import { sheets } from 'googleapis/build/src/apis/sheets';

interface SheetOptions {
  spreadsheetId: string
  range: string
  client: JWT
}

type CellValue = string | number
type SheetRow = Array<CellValue>

type WriteSheetOptions = SheetOptions & {
  rows: Array<SheetRow>
}

interface WriteOptions {
  range: string
  rows: Array<SheetRow>
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

  append(range: string, rows: Array<SheetRow>) {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.append({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: rows
        },
        auth: this.client
      }).then((response) => response.data.updates ? response.data.updates.updatedCells : null);
    });
  }

  update(range: string, rows: Array<SheetRow>) {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.update({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueInputOption: "USER_ENTERED",
        requestBody: {
          range: range,
          values: rows,
        },
        auth: this.client
      }).then((response) => response.data.updatedCells);
    });
  }

  get(range: string) {
    return this.checkAuthAnd(() => {
      return this.sheets.spreadsheets.values.get({
        spreadsheetId: this.spreadsheetId,
        range: range,
        valueRenderOption: 'FORMATTED_VALUE'
      }).then((response) => response.data.values);
    });
  }
}

export default Sheet;
