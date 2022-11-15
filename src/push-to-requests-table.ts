/* global Logger, SpreadsheetApp, GoogleAppsScript */
/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars */
function getRandomString(): string {
    const fromRandom = Math.random().toString(32).replace('0.', '');
    const fromTime = Date.now().toString(32);

    return `${fromRandom}${fromTime}`.toLowerCase();
}

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';

// first row with data
const dataRowBegin = 3;
const managerColumnBegin = 'A';
const managerColumnEnd = 'I';
const requestsColumnBegin = 'J';
const requestsColumnEnd = 'T';
const tableIdColumnName = 'AY';
const firstColumnName = 'A';
const rowIdColumnName = 'AZ'; // should be last column
// const requestDataRange = 'E3:H';

type ConnectTableConfigType = Readonly<{
    requestsTableId: string;
}>;

class PushToRequestsTable {
    private static requestsTableId = '';

    static initialize(connectTableConfig: ConnectTableConfigType) {
        PushToRequestsTable.requestsTableId = connectTableConfig.requestsTableId;
        PushToRequestsTable.makeUiMenu();
        PushToRequestsTable.updateRowsId();

        // protect row id
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        sheet.getRange(rowIdColumnName).protect().setDescription('Protected range.');
    }

    static updateRowsId() {
        const range = PushToRequestsTable.getAllDataRange();
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row.pop() || '').trim();
            const rowNumber: number = index + dataRowBegin;
            const hasRowValue = row.join('').trim().length > 0;
            const cellIdRange = sheet.getRange(`${rowIdColumnName}${rowNumber.toString(10)}`);

            if (hasRowValue && columnId) {
                return;
            }

            if (!hasRowValue && !columnId) {
                return;
            }

            if (hasRowValue && !columnId) {
                cellIdRange.setValue(getRandomString());
                return;
            }

            if (!hasRowValue && columnId) {
                cellIdRange.setValue('');
            }
        });
    }

    static getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            `${firstColumnName}${dataRowBegin.toString(10)}:${rowIdColumnName}`
        );

        return range;
    }

    static pushDataToRequestTable() {
        const requestsSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(requestsTableId);
        const sheet: GoogleAppsScript.Spreadsheet.Sheet | null = requestsSheet.getSheetByName('requests');

        if (!sheet) {
            const appUI = SpreadsheetApp.getUi();

            appUI.alert('!pushDataToRequestTable');
            return;
        }

        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            `${firstColumnName}${dataRowBegin.toString(10)}:${rowIdColumnName}`
        );

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row.pop() || '').trim();
            const rowNumber: number = index + dataRowBegin;
            const hasRowValue = row.join('').trim().length > 0;
            const cellIdRange = sheet.getRange(`${rowIdColumnName}${rowNumber.toString(10)}`);

            if (hasRowValue && columnId) {
                return;
            }

            if (!hasRowValue && !columnId) {
                return;
            }

            if (hasRowValue && !columnId) {
                cellIdRange.setValue('1');
                return;
            }

            if (!hasRowValue && columnId) {
                cellIdRange.setValue('2');
            }
        });

        /*
                const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
                const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();
                const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(managerDataRange);

                range.getValues().forEach((row: unknown) => {
                    Logger.log(JSON.stringify(row));
                });

                Logger.log(sheet.getSheetId());
                Logger.log(sheet.getSheetName());

                range.sort({ascending: true, column: 2});
        */

        Logger.log('////////////////');
        // const appUI = SpreadsheetApp.getUi();
        // appUI.alert("pushDataToRequestTable");
    }

    static makeUiMenu() {
        const appUI = SpreadsheetApp.getUi();
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data');

        menu.addItem('push to requests table', 'PushToRequestsTable.pushDataToRequestTable');
        menu.addToUi();
    }
}

// will call on document open
// eslint-disable-next-line no-unused-vars, @typescript-eslint/no-unused-vars
function onOpen() {
    PushToRequestsTable.initialize({requestsTableId});
}

function onEdit() {
    PushToRequestsTable.updateRowsId();
}
