/* eslint-disable @typescript-eslint/no-unused-vars, no-unused-vars, @typescript-eslint/no-use-before-define, sort-keys */

/* global Logger, SpreadsheetApp, GoogleAppsScript */

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';
const requestsSheetName = 'Requests';

const managerTable1Id = '11ZNH5S8DuZUobQU6sw-_Svx5vVt62I9CmxSSF_eGFDM';
const managerTable2Id = '1C3pU0hsaGZnsztX72ZND6WRsHBA5EyY5ozEilL2CrOA';
const managerTableIdList: Set<string> = new Set([managerTable1Id, managerTable2Id]);

enum RowActionNameEnum {
    remove = 'remove',
    skip = 'skip',
    updateOrAdd = 'update/add',
}

// ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
const managerColumnList: Array<string> = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'];
const requestsColumnList: Array<string> = ['J', 'K', 'L', 'M', 'N', 'O'];
const commonColumnList: Array<string> = ['P', 'Q', 'R', 'S', 'T', 'U', 'V'];
const rowIdColumnName = 'AZ';
const rowActionColumnName = 'BA';

// first row with data
const dataRowBegin = 3;
const firstColumnName = 'A';
const [lastColumnName] = [rowIdColumnName, rowActionColumnName].sort().reverse();

Logger.log(lastColumnName);

const appUI: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();

const util = {
    // eslint-disable-next-line max-statements, complexity
    /*
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
            return list[columnNumber - 1];
        }

        return list[Math.floor(columnNumber / alphabetLength) - 1] + list[(columnNumber % alphabetLength) - 1];
    },
*/
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
    getIsSkipRow(row: Array<unknown>): boolean {
        const cellRawValue = util
            .stringify(row[util.columnStringToNumber(rowActionColumnName) - 1])
            .trim()
            .toLowerCase();

        return cellRawValue === RowActionNameEnum.skip;
    },
    getIsRemoveRow(row: Array<unknown>): boolean {
        const cellRawValue = util
            .stringify(row[util.columnStringToNumber(rowActionColumnName) - 1])
            .trim()
            .toLowerCase();

        return cellRawValue === RowActionNameEnum.remove;
    },
    getIsUpdateOrAdd(row: Array<unknown>): boolean {
        const cellRawValue = util
            .stringify(row[util.columnStringToNumber(rowActionColumnName) - 1])
            .trim()
            .toLowerCase();

        return cellRawValue === RowActionNameEnum.updateOrAdd;
    },
    // eslint-disable-next-line max-statements, complexity
    stringify(value: unknown): string {
        if (typeof value === 'string') {
            return value.trim();
        }

        if (typeof value === 'boolean') {
            return value.toString().toLowerCase();
        }

        if (typeof value === 'number') {
            return value.toString(10);
        }

        if (!value) {
            return '';
        }

        return String(value);
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
        range.getValues().forEach((row: Array<unknown>, rowIndex: number) => {
            const rwoId = util.stringify(row[util.columnStringToNumber(rowIdColumnName) - 1]);
            const rowAction = util.stringify(row[util.columnStringToNumber(rowActionColumnName) - 1]);
            const rowNumber: number = rowIndex + dataRowBegin;
            const hasRowValue = row.join('').replace(rwoId, '').replace(rowAction, '').trim().length > 0;
            const cellIdRange = sheet.getRange(`${rowIdColumnName}${rowNumber}`);

            if (hasRowValue && rwoId) {
                return;
            }

            if (!hasRowValue && !rwoId) {
                return;
            }

            if (hasRowValue && !rwoId) {
                cellIdRange.setValue(util.getRandomString());
                return;
            }

            if (!hasRowValue && rwoId) {
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

        menu.addItem('Push to Requests Table', 'managerTable.pushDataToRequestTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToRequestTable() {
        // remove rows in requests table
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[util.columnStringToNumber(rowIdColumnName) - 1]);

                if (!managerRowId) {
                    return;
                }

                const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsTable.getAllRequestsDataRange();
                const requestsSheet = requestsTable.getRequestsSheet();

                const requestsRangeRowIndex: number = requestsRange
                    .getValues()
                    .findIndex((requestsRow: Array<unknown>): boolean => {
                        const requestsRowId = util.stringify(
                            requestsRow[util.columnStringToNumber(rowIdColumnName) - 1]
                        );

                        return requestsRowId === managerRowId;
                    });

                if (requestsRangeRowIndex === -1) {
                    return;
                }

                requestsSheet.deleteRow(requestsRangeRowIndex + dataRowBegin);
            });

        // remove rows in manager table
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[util.columnStringToNumber(rowIdColumnName) - 1]);

                if (!managerRowId) {
                    return;
                }

                const managerRemoveRange: GoogleAppsScript.Spreadsheet.Range = managerTable.getAllDataRange();
                const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActive().getActiveSheet();

                const removeManagerRangeRowIndex: number = managerRemoveRange
                    .getValues()
                    .findIndex((removeManagerRow: Array<unknown>): boolean => {
                        const removeManagerRowId = util.stringify(
                            removeManagerRow[util.columnStringToNumber(rowIdColumnName) - 1]
                        );

                        return removeManagerRowId === managerRowId;
                    });

                if (removeManagerRangeRowIndex === -1) {
                    return;
                }

                sheet.deleteRow(removeManagerRangeRowIndex + dataRowBegin);
            });

        // update rows
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsUpdateOrAdd(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[util.columnStringToNumber(rowIdColumnName) - 1]);

                if (!managerRowId) {
                    return;
                }

                const requestsRange: GoogleAppsScript.Spreadsheet.Range = requestsTable.getAllRequestsDataRange();
                const requestsSheet = requestsTable.getRequestsSheet();

                const requestsRangeRowIndex: number = requestsRange
                    .getValues()
                    .findIndex((requestsRow: Array<unknown>): boolean => {
                        const requestsRowId = util.stringify(
                            requestsRow[util.columnStringToNumber(rowIdColumnName) - 1]
                        );

                        return requestsRowId === managerRowId;
                    });

                const requestsRangeRowNumber: number =
                    requestsRangeRowIndex === -1
                        ? requestsSheet.getLastRow() + 1
                        : requestsRangeRowIndex + dataRowBegin;

                managerRow.forEach((managerRowData: unknown, managerColumnIndex: number) => {
                    const currentColumnNumber = managerColumnIndex + 1;
                    const isInManagerColumnRange = managerColumnList
                        .map(util.columnStringToNumber)
                        .includes(currentColumnNumber);
                    const isInCommonColumnRange = commonColumnList
                        .map(util.columnStringToNumber)
                        .includes(currentColumnNumber);
                    const isRowIdColumn = currentColumnNumber === util.columnStringToNumber(rowIdColumnName);

                    if (isInManagerColumnRange || isInCommonColumnRange || isRowIdColumn) {
                        requestsSheet.getRange(requestsRangeRowNumber, currentColumnNumber).setValue(managerRowData);
                    }
                });
            });
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
