/* eslint-disable
 @typescript-eslint/no-unused-vars,
 no-unused-vars,
 @typescript-eslint/no-use-before-define,
 sort-keys,
 unicorn/prefer-set-has,
 no-plusplus,
 sonarjs/no-duplicate-string
*/

/* global Logger, SpreadsheetApp, GoogleAppsScript */

// main constants
const requestsTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';
const requestsSheetName = 'Requests';

const managerTable1Id = '11ZNH5S8DuZUobQU6sw-_Svx5vVt62I9CmxSSF_eGFDM';
const managerTable2Id = '1C3pU0hsaGZnsztX72ZND6WRsHBA5EyY5ozEilL2CrOA';
const managerTableIdList: Array<string> = [managerTable1Id, managerTable2Id];

enum RowActionNameEnum {
    remove = 'remove',
    skip = 'skip',
    updateOrAdd = 'update/add',
}

type RowDataType = {
    rowDataList: Array<unknown>;
    rowId: string;
    rowNumber: number;
    sheet: GoogleAppsScript.Spreadsheet.Sheet | null;
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null;
};

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

const allDataRange = `${firstColumnName}${dataRowBegin}:${lastColumnName}`;

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
    showProgress(title: string, partOfProgress: number) {
        SpreadsheetApp.getActiveSpreadsheet().toast('', `${title}${(partOfProgress * 100).toFixed(2)}`, -1);
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

        return managerTableIdList.includes(spreadsheetId);
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
        const cellRawValue = util.stringify(row[rowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.skip;
    },
    getIsRemoveRow(row: Array<unknown>): boolean {
        const cellRawValue = util.stringify(row[rowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.remove;
    },
    getIsUpdateOrAdd(row: Array<unknown>): boolean {
        const cellRawValue = util.stringify(row[rowActionColumnIndex]).trim().toLowerCase();

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

const rowIdColumnIndex = util.columnStringToNumber(rowIdColumnName) - 1;
const rowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
const managerColumnNumberList: Array<number> = managerColumnList.map<number>(util.columnStringToNumber);
const requestsColumnNumberList: Array<number> = requestsColumnList.map<number>(util.columnStringToNumber);
const commonColumnNumberList: Array<number> = commonColumnList.map<number>(util.columnStringToNumber);

const mainTable = {
    updateRowsId() {
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();

        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(allDataRange);

        // eslint-disable-next-line complexity
        range.getValues().forEach((row: Array<unknown>, rowIndex: number) => {
            const rwoId = util.stringify(row[rowIdColumnIndex]);
            const rowAction = util.stringify(row[rowActionColumnIndex]);
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
    getRowDataById(searchingRowId: string, tableIdList: Array<string>): RowDataType {
        let rowData: RowDataType = {
            rowId: '',
            rowNumber: 0,
            spreadsheet: null,
            rowDataList: [],
            sheet: null,
        };

        tableIdList.some((spreadsheetId: string): boolean => {
            const spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(spreadsheetId);
            const sheetList: Array<GoogleAppsScript.Spreadsheet.Sheet> = spreadsheet.getSheets();

            return sheetList.some((sheet: GoogleAppsScript.Spreadsheet.Sheet): boolean => {
                const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(allDataRange);

                return range.getValues().some((row: Array<unknown>, rowIndex: number): boolean => {
                    const rowId = util.stringify(row[rowIdColumnIndex]);

                    if (rowId !== searchingRowId) {
                        return false;
                    }

                    rowData = {
                        rowId,
                        rowNumber: rowIndex + dataRowBegin,
                        spreadsheet,
                        rowDataList: [...row],
                        sheet,
                    };

                    return true;
                });
            });
        });

        return rowData;
    },
};

const managerTable = {
    getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();

        return sheet.getRange(allDataRange);
    },

    makeUiMenuForManager() {
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to requests table');

        menu.addItem('Push to Requests Table', 'managerTable.pushDataToRequestTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToRequestTable() {
        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 0/3', 'Syncing...', -1);

        // remove rows in requests table
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[rowIdColumnIndex]);

                if (!managerRowId) {
                    return;
                }

                const requestsRowData = mainTable.getRowDataById(managerRowId, [requestsTableId]);

                requestsRowData.sheet?.deleteRow(requestsRowData.rowNumber);
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 1/3', 'Syncing...', -1);

        // remove rows in manager table
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[rowIdColumnIndex]);

                if (!managerRowId) {
                    return;
                }

                const managerRowData = mainTable.getRowDataById(managerRowId, [SpreadsheetApp.getActive().getId()]);

                managerRowData.sheet?.deleteRow(managerRowData.rowNumber);
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 2/3', 'Syncing...', -1);

        // update rows
        managerTable
            .getAllDataRange()
            .getValues()
            .forEach((managerRow: Array<unknown>) => {
                if (!util.getIsUpdateOrAdd(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[rowIdColumnIndex]);

                if (!managerRowId) {
                    return;
                }

                const requestsSheet = requestsTable.getRequestsSheet();
                const requestsRowData = mainTable.getRowDataById(managerRowId, [requestsTableId]);
                const requestsRangeRowNumber: number = requestsRowData.sheet
                    ? requestsRowData.rowNumber
                    : requestsSheet.getLastRow() + 1;

                managerRow.forEach((managerRowData: unknown, managerColumnIndex: number) => {
                    const currentColumnNumber = managerColumnIndex + 1;
                    const isInManagerColumnRange = managerColumnNumberList.includes(currentColumnNumber);
                    const isInCommonColumnRange = commonColumnNumberList.includes(currentColumnNumber);
                    const isRowIdColumn = managerColumnIndex === rowIdColumnIndex;

                    if (isInManagerColumnRange || isInCommonColumnRange || isRowIdColumn) {
                        requestsSheet.getRange(requestsRangeRowNumber, currentColumnNumber).setValue(managerRowData);
                    }
                });
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 3/3', 'Synced!', 2);
    },
};

const requestsTable = {
    getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const requestsSheet = requestsTable.getRequestsSheet();

        return requestsSheet.getRange(allDataRange);
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
        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 0/3', 'Syncing...', -1);

        // remove rows in manager table
        requestsTable
            .getAllDataRange()
            .getValues()
            .forEach((requestsRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(requestsRow)) {
                    return;
                }

                const requestsRowId = util.stringify(requestsRow[rowIdColumnIndex]);

                if (!requestsRowId) {
                    return;
                }

                const managerRowData = mainTable.getRowDataById(requestsRowId, managerTableIdList);

                managerRowData.sheet?.deleteRow(managerRowData.rowNumber);
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 1/3', 'Syncing...', -1);

        // remove rows in requests table
        requestsTable
            .getAllDataRange()
            .getValues()
            .forEach((requestsRow: Array<unknown>) => {
                if (!util.getIsRemoveRow(requestsRow)) {
                    return;
                }

                const requestsRowId = util.stringify(requestsRow[rowIdColumnIndex]);

                if (!requestsRowId) {
                    return;
                }

                const requestsRowData = mainTable.getRowDataById(requestsRowId, [requestsTableId]);

                requestsRowData.sheet?.deleteRow(requestsRowData.rowNumber);
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 2/3', 'Syncing...', -1);

        // update rows
        requestsTable
            .getAllDataRange()
            .getValues()
            .forEach((requestsRow: Array<unknown>) => {
                const requestsRowId = util.stringify(requestsRow[rowIdColumnIndex]);

                if (!requestsRowId) {
                    return;
                }

                const managerRowData = mainTable.getRowDataById(requestsRowId, managerTableIdList);
                const {sheet: managerSheet, rowNumber: managerRowNumber} = managerRowData;

                if (!managerSheet) {
                    appUI.alert(`Can not find row with id ${requestsRowId}`);
                    return;
                }

                requestsRow.forEach((requestsRowData: unknown, requestColumnIndex: number) => {
                    const currentColumnNumber = requestColumnIndex + 1;
                    const isInRequestsColumnRange = requestsColumnNumberList.includes(currentColumnNumber);
                    const isInCommonColumnRange = commonColumnNumberList.includes(currentColumnNumber);

                    if (isInRequestsColumnRange || isInCommonColumnRange) {
                        managerSheet.getRange(managerRowNumber, currentColumnNumber).setValue(requestsRowData);
                    }
                });
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 3/3', 'Synced!', 2);
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
