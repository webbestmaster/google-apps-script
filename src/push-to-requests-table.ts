/* global Logger, SpreadsheetApp, GoogleAppsScript */

/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars */

function columnStringToNumber(columnStringRaw: string): number {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const alphabetLength: number = alphabet.length;
    const list: Array<string> = [...alphabet.toUpperCase()];
    const columnString = columnStringRaw.trim().toUpperCase();

    return [...columnString].reverse().reduce<number>((sum: number, char: string, index: number) => {
        return sum + (list.indexOf(char) + 1) * Math.pow(alphabetLength, index);
    }, 0);
}

// eslint-disable-next-line max-statements, complexity
function columnNumberToString(columnNumber: number): string {
    const alphabet = 'abcdefghijklmnopqrstuvwxyz';
    const alphabetLength: number = alphabet.length;
    const list: Array<string> = [...alphabet.toUpperCase()];
    const minColumnNumber = 1;
    const maxColumnNumber = alphabetLength * (alphabetLength + 1);
    const appUI = SpreadsheetApp.getUi();

    if (Math.round(columnNumber) !== columnNumber) {
        const errorMessage = `The column number is not integer. Column number is ${columnNumber}`;

        appUI.alert(errorMessage);

        throw new Error(errorMessage);
    }

    if (columnNumber < minColumnNumber) {
        // eslint-disable-next-line max-len
        const errorMessage = `The column number is too small. Column number is ${columnNumber}, but min is ${minColumnNumber}.`;

        appUI.alert(errorMessage);

        throw new Error(errorMessage);
    }

    if (columnNumber > maxColumnNumber) {
        // eslint-disable-next-line max-len
        const errorMessage = `The column number is too big. Column number is ${columnNumber}, but max is ${maxColumnNumber}.`;

        appUI.alert(errorMessage);

        throw new Error(errorMessage);
    }

    if (columnNumber === maxColumnNumber) {
        return list[alphabetLength - 1] + list[alphabetLength - 1];
    }

    if (columnNumber <= alphabetLength) {
        return list[columnNumber - 1];
    }

    return list[Math.floor(columnNumber / alphabetLength) - 1] + list[(columnNumber % alphabetLength) - 1];
}

/*
console.log(columnNumberToString(1));
console.log(columnStringToNumber('a'));

console.log(columnNumberToString(5));
console.log(columnStringToNumber('e'));

console.log(columnNumberToString(57));
console.log(columnStringToNumber('be'));

console.log(columnNumberToString(26));
console.log(columnStringToNumber('z'));

console.log(columnNumberToString(27));
console.log(columnStringToNumber('aa'));

console.log(columnNumberToString(700));
console.log(columnStringToNumber('zx'));

console.log(columnNumberToString(701));
console.log(columnStringToNumber('zy'));

console.log(columnNumberToString(702));
console.log(columnStringToNumber('zz'));
*/

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
const managerColumnEndString = 'I';
const requestsColumnBeginString = 'J';
const requestsColumnEndString = 'T';
const tableIdColumnName = 'AY';
const firstColumnName = 'A';
const lastColumnName = 'AZ';
const rowIdColumnName = 'AZ';
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
        /*
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        sheet.getRange(rowIdColumnName).protect().setDescription('Protected range.');
*/
    }

    static updateRowsId() {
        const range = PushToRequestsTable.getAllDataRange();
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row[columnStringToNumber(rowIdColumnName) - 1] || '').trim();
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
            `${firstColumnName}${dataRowBegin.toString(10)}:${lastColumnName}`
        );

        return range;
    }

    // eslint-disable-next-line sonarjs/cognitive-complexity
    static pushDataToRequestTable() {
        const requestsSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(requestsTableId);
        const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null = requestsSpreadsheet.getSheetByName('requests');

        if (!requestsSheet) {
            const appUI = SpreadsheetApp.getUi();

            appUI.alert('!pushDataToRequestTable');
            return;
        }

        const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(
            `${firstColumnName}${dataRowBegin.toString(10)}:${lastColumnName}`
        );

        const managerRange: GoogleAppsScript.Spreadsheet.Range = PushToRequestsTable.getAllDataRange();

        managerRange.getValues().forEach((managerRow: Array<unknown>, managerIndex: number) => {
            const managerColumnId = String(managerRow[columnStringToNumber(rowIdColumnName) - 1] || '').trim();
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
                        const requestColumnName = columnNumberToString(managerCellValueIndex + 1);
                        const requestRowNumber = requestsRowIndex + dataRowBegin;

                        const requestsCellIdRange = requestsSheet.getRange(
                            `${requestColumnName}${requestRowNumber.toString(10)}`
                        );

                        requestsCellIdRange.setValue(managerCellValue);
                    });
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
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to requests table');

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
