/* global Logger, SpreadsheetApp, GoogleAppsScript */

/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars */

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';
const requestsSheetName = 'requests';

// first row with data
const dataRowBegin = 3;
const firstColumnName = 'A';
const lastColumnName = 'AZ';

const managerColumnBeginString = 'A';
const managerColumnEndString = 'I';

const requestsColumnBeginString = 'J';
const requestsColumnEndString = 'T';

const commonColumnBeginString = 'J';
const commonColumnEndString = 'T';

const tableIdColumnName = 'AY';
const rowIdColumnName = 'AZ';
// const requestDataRange = 'E3:H';

const util = {
    // eslint-disable-next-line max-statements, complexity
    columnNumberToString(columnNumber: number): string {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const alphabetLength: number = alphabet.length;
        const list: Array<string> = [...alphabet.toUpperCase()];
        const lastLetter = list[alphabetLength - 1];
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
            return lastLetter + lastLetter;
        }

        if (columnNumber <= alphabetLength) {
            return lastLetter;
        }

        return list[Math.floor(columnNumber / alphabetLength) - 1] + list[(columnNumber % alphabetLength) - 1];
    },
    // eslint-disable-next-line max-statements, complexity
    columnStringToNumber(columnStringRaw: string): number {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const alphabetLength: number = alphabet.length;
        const list: Array<string> = [...alphabet.toUpperCase()];
        const columnString = columnStringRaw.trim().toUpperCase();

        return [...columnString].reverse().reduce<number>((sum: number, char: string, index: number) => {
            return sum + (list.indexOf(char) + 1) * Math.pow(alphabetLength, index);
        }, 0);
    },
    getRandomString(): string {
        const fromRandom = Math.random().toString(32).replace('0.', '');
        const fromTime = Date.now().toString(32);

        return `${fromRandom}${fromTime}`.toLowerCase();
    },
};

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

const pushToRequestsTable = {
    getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            `${firstColumnName}${dataRowBegin}:${lastColumnName}`
        );

        return range;
    },

    getAllRequestsDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const requestsSheet = pushToRequestsTable.getRequestsSheet();

        const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(
            `${firstColumnName}${dataRowBegin}:${lastColumnName}`
        );

        return requestsRange;
    },

    getRequestsRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range {
        const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null = pushToRequestsTable.getRequestsSheet();

        const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(a1Notation);

        return requestsRange;
    },

    getRequestsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const requestsSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(requestsTableId);
        const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
            requestsSpreadsheet.getSheetByName(requestsSheetName);

        const appUI: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();

        if (!requestsSheet) {
            const errorMessage = '[getRequestsSheet]: Can not requests table and/or requests sheet.';

            appUI.alert(errorMessage);

            throw new Error(errorMessage);
        }

        return requestsSheet;
    },

    initialize() {
        pushToRequestsTable.makeUiMenu();
        pushToRequestsTable.updateRowsId();

        // protect row id
        /*
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        sheet.getRange(rowIdColumnName).protect().setDescription('Protected range.');
*/
    },

    makeUiMenu() {
        const appUI = SpreadsheetApp.getUi();
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to requests table');

        menu.addItem('push to requests table', 'pushToRequestsTable.pushDataToRequestTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToRequestTable() {
        const managerRange: GoogleAppsScript.Spreadsheet.Range = pushToRequestsTable.getAllDataRange();

        managerRange.getValues().forEach((managerRow: Array<unknown>) => {
            const managerColumnId = String(managerRow[util.columnStringToNumber(rowIdColumnName) - 1] || '').trim();

            if (!managerColumnId) {
                return;
            }

            const requestsRange: GoogleAppsScript.Spreadsheet.Range = pushToRequestsTable.getAllRequestsDataRange();
            const requestsSheet = pushToRequestsTable.getRequestsSheet();
            const startColumnNumber = util.columnStringToNumber(managerColumnBeginString);
            const endColumnNumber = util.columnStringToNumber(managerColumnEndString);

            const requestsRangeRowIndex: number = requestsRange
                .getValues()
                .findIndex((requestsRow: Array<unknown>): boolean => {
                    const requestsColumnId = String(
                        requestsRow[util.columnStringToNumber(rowIdColumnName) - 1] || ''
                    ).trim();

                    return requestsColumnId === managerColumnId;
                });

            const requestsRangeRowNumber: number =
                requestsRangeRowIndex === -1 ? requestsSheet.getLastRow() + 1 : requestsRangeRowIndex + dataRowBegin;

            managerRow.forEach((managerRowData: unknown, managerColumnIndex: number) => {
                const currentColumnNumber = managerColumnIndex + 1;

                if (currentColumnNumber < startColumnNumber && currentColumnNumber > endColumnNumber) {
                    return;
                }

                requestsSheet.getRange(requestsRangeRowNumber, currentColumnNumber).setValue(managerRowData);
            });
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
    },

    updateRowsId() {
        const range = pushToRequestsTable.getAllDataRange();
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row[util.columnStringToNumber(rowIdColumnName) - 1] || '').trim();
            const rowNumber: number = index + dataRowBegin;
            const hasRowValue = row.join('').trim().length > 0;
            const cellIdRange = sheet.getRange(`${rowIdColumnName}${rowNumber}`);
            // const cellTableIdRange = sheet.getRange(`${tableIdColumnName}${rowNumber}`);

            if (hasRowValue && columnId) {
                return;
            }

            if (!hasRowValue && !columnId) {
                return;
            }

            if (hasRowValue && !columnId) {
                cellIdRange.setValue(util.getRandomString());
                return;
            }

            if (!hasRowValue && columnId) {
                cellIdRange.setValue('');
            }
        });
    },
};

// will call on document open
function onOpen() {
    pushToRequestsTable.initialize();
}

function onEdit() {
    pushToRequestsTable.updateRowsId();
}
