/* global Logger, SpreadsheetApp, GoogleAppsScript */

/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars */

// TODO: fix it, make for any number
function columnStringToNumber(columnString: string): number {
    const list: Array<string> = 'abcdefghijklmnopqrstuvwxyz'.toUpperCase().split('');

    if (columnString === '') {
        throw new Error('columnStringToNumber: can not read empty string');
    }

    if (columnString.length === 1) {
        return list.indexOf(columnString) + 1;
    }

    if (columnString.length === 2) {
        return (list.indexOf(columnString[0]) + 1) * 26 + columnStringToNumber(columnString[1]);
    }

}


function getRandomString(): string {
    const fromRandom = Math.random().toString(32).replace('0.', '');
    const fromTime = Date.now().toString(32);

    return `${fromRandom}${fromTime}`.toLowerCase();
}

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';

// first row with data
const dataRowBegin = 3;
const managerColumnBeginString = 'A';
const managerColumnBeginNumber = 1;
const managerColumnEndString = 'I';
const managerColumnEndNumber = 9;
const requestsColumnBeginString = 'J';
const requestsColumnBeginNumber = 10;
const requestsColumnEndString = 'T';
const requestsColumnEndNumber = 21;
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
        const requestsSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(requestsTableId);
        const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null = requestsSpreadsheet.getSheetByName('requests');

        if (!requestsSheet) {
            const appUI = SpreadsheetApp.getUi();

            appUI.alert('!pushDataToRequestTable');
            return;
        }

        const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(
            `${firstColumnName}${dataRowBegin.toString(10)}:${rowIdColumnName}`
        );

        const managerRange: GoogleAppsScript.Spreadsheet.Range = PushToRequestsTable.getAllDataRange();

        managerRange.getValues()
            .forEach((managerRow: Array<unknown>, managerIndex: number) => {
                const managerColumnId = String(managerRow.pop() || '').trim();
                const managerRowNumber: number = managerIndex + dataRowBegin;

                let isRowUpdated = false;
                // try to find needed row in requests table
                requestsRange.getValues().forEach((requestsRow: Array<unknown>, requestsRowIndex: number) => {
                    const requestsColumnId = String(requestsRow.pop() || '').trim();
                    if (isRowUpdated) {
                        return;
                    }

                    if (requestsColumnId === managerColumnId) {
                        isRowUpdated = true;

                        managerRow.forEach((managerCellValue: unknown, managerCellValueIndex: number) => {
                            const requestsCellIdRange = requestsSheet
                                .getRange(`${rowIdColumnName}${requestsRowIndex.toString(10)}`);

                            requestsCellIdRange.setValue('1');
                        })
                    }


                });


            });

        // eslint-disable-next-line complexity
        requestsRange.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row.pop() || '').trim();
            const rowNumber: number = index + dataRowBegin;
            const hasRowValue = row.join('').trim().length > 0;
            const cellIdRange = requestsSheet.getRange(`${rowIdColumnName}${rowNumber.toString(10)}`);

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
