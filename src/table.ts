/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars, @typescript-eslint/no-use-before-define */

/* global Logger, SpreadsheetApp, GoogleAppsScript */

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';
const requestsSheetName = 'Requests';

const managerTable1Id = '11ZNH5S8DuZUobQU6sw-_Svx5vVt62I9CmxSSF_eGFDM';
const managerTable2Id = '1C3pU0hsaGZnsztX72ZND6WRsHBA5EyY5ozEilL2CrOA';
const managerTableIdList: Set<string> = new Set([managerTable1Id, managerTable2Id]);

// ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
const managerColumnList: Set<string> = new Set(['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']);
const requestsColumnList: Set<string> = new Set(['J', 'K', 'L', 'M', 'N', 'O']);
const commonColumnList: Set<string> = new Set(['P', 'Q', 'R', 'S', 'T', 'U', 'V']);
const rowIdColumnName = 'AZ';
const skipRowColumnName = 'BA';
const removeRowColumnName = 'BB';

// first row with data
const dataRowBegin = 3;
const firstColumnName = 'A';
const [lastColumnName] = [rowIdColumnName, skipRowColumnName, removeRowColumnName].sort().reverse();

Logger.log(lastColumnName);

const appUI: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();

const util = {
    // eslint-disable-next-line max-statements, complexity
    columnNumberToString(columnNumber: number): string {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const alphabetLength: number = alphabet.length;
        const list: Array<string> = [...alphabet.toUpperCase()];
        const lastLetter = list[alphabetLength - 1];
        const minColumnNumber = 1;
        const maxColumnNumber = alphabetLength * (alphabetLength + 1);

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
    columnStringToNumber(columnStringRaw: string): number {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const alphabetLength: number = alphabet.length;
        const list: Array<string> = [...alphabet.toUpperCase()];
        const columnString = columnStringRaw.trim().toUpperCase();

        return [...columnString].reverse().reduce<number>((sum: number, char: string, index: number) => {
            return sum + (list.indexOf(char) + 1) * Math.pow(alphabetLength, index);
        }, 0);
    },
    getIsManagerSpreadsheet(): boolean {
        const spreadsheetId = SpreadsheetApp.getActive().getId();

        return managerTableIdList.has(spreadsheetId);
    },
    getIsRequestsSpreadsheet(): boolean {
        const spreadsheetId = SpreadsheetApp.getActive().getId();

        return spreadsheetId === requestsTableId;
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

const mainTable = {
    updateRowsId() {
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();

        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            `${firstColumnName}${dataRowBegin}:${lastColumnName}`
        );

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, index: number) => {
            const columnId = String(row[util.columnStringToNumber(rowIdColumnName) - 1] || '').trim();
            const rowNumber: number = index + dataRowBegin;
            const hasRowValue = row.join('').replace(columnId, '').trim() !== '';
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

const managerTable = {
    getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(
            `${firstColumnName}${dataRowBegin}:${lastColumnName}`
        );

        return range;
    },

    makeUiMenuForManager() {
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to requests table');

        menu.addItem('push to requests table', 'managerTable.pushDataToRequestTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToRequestTable() {
        const managerRange: GoogleAppsScript.Spreadsheet.Range = managerTable.getAllDataRange();

        managerRange.getValues().forEach((managerRow: Array<unknown>) => {
            const managerColumnId = String(managerRow[util.columnStringToNumber(rowIdColumnName) - 1] || '').trim();

            if (!managerColumnId) {
                return;
            }

            const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsTable.getAllRequestsDataRange();
            const requestsSheet = requestsTable.getRequestsSheet();

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
                const currentColumnString = util.columnNumberToString(currentColumnNumber);
                const isInManagerColumnRange = managerColumnList.has(currentColumnString);
                const isInCommonColumnRange = commonColumnList.has(currentColumnString);
                const isRowIdColumn = currentColumnString === rowIdColumnName;

                Logger.log(managerColumnIndex + '-' + managerRowData);

                if (isInManagerColumnRange || isInCommonColumnRange || isRowIdColumn) {
                    requestsSheet.getRange(requestsRangeRowNumber, currentColumnNumber).setValue(managerRowData);
                }
            });
        });

        SpreadsheetApp.flush();
    },
};

const requestsTable = {
    getAllRequestsDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const requestsSheet = requestsTable.getRequestsSheet();

        const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(
            `${firstColumnName}${dataRowBegin}:${lastColumnName}`
        );

        return requestsRange;
    },

    getRequestsSheet(): GoogleAppsScript.Spreadsheet.Sheet {
        const requestsSpreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(requestsTableId);
        const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null =
            requestsSpreadsheet.getSheetByName(requestsSheetName);

        if (!requestsSheet) {
            const errorMessage = '[getRequestsSheet]: Can not requests table and/or requests sheet.';

            appUI.alert(errorMessage);
            throw new Error(errorMessage);
        }

        return requestsSheet;
    },

    makeUiMenuForRequests() {
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to managers table');

        menu.addItem('push to managers table', 'requestsTable.pushDataToManagerTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToManagerTable() {
        appUI.alert('pushDataToManagerTable');
    },

    /*
        getRequestsRange(a1Notation: string): GoogleAppsScript.Spreadsheet.Range {
            const requestsSheet: GoogleAppsScript.Spreadsheet.Sheet | null = requestsTable.getRequestsSheet();

            const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsSheet.getRange(a1Notation);

            return requestsRange;
        },
    */
};

// will call on document open
function onOpen() {
    if (util.getIsManagerSpreadsheet()) {
        managerTable.makeUiMenuForManager();
    }

    if (util.getIsRequestsSpreadsheet()) {
        requestsTable.makeUiMenuForRequests();
    }

    mainTable.updateRowsId();
}

function onEdit() {
    mainTable.updateRowsId();
}
