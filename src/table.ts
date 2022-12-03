/* eslint-disable
 @typescript-eslint/no-unused-vars,
 no-unused-vars,
 @typescript-eslint/no-use-before-define,
 sort-keys,
 unicorn/prefer-set-has,
 no-plusplus,
 sonarjs/no-duplicate-string
*/

/* global Logger, SpreadsheetApp, GoogleAppsScript, UrlFetchApp */

// main constants
const requestsTableId = '1qmpdaS_EWJJd-C8ntLhhQ-AqrCVuI4ahBFnP0gx9ZT8';
const requestsSheetName = 'Общая статистика';
const requestsRequiredColumnName = ['R', 'S', 'T', 'V'];

const managerTable1Id = '1mD4Fxu1r4_lVvZ4h5mZ1SorsKW1dsNBkkQoOZn0P1o0';
const managerRequiredColumnName = ['B', 'C', 'D', 'H', 'I', 'M', 'N', 'O', 'Q'];
const managerTableIdList: Array<string> = [managerTable1Id];
const bgColorSynced = '#00D100';
const bgColorChanged = '#ECF87F';
const bgColorDefault = '#FFFFFF';

const enum RowActionNameEnum {
    remove = 'remove',
    skip = 'skip',
    update = 'update',
    updateOrAdd = 'update/add',
}

type RowDataType = {
    rowDataList: Array<unknown>;
    rowId: string;
    rowNumber: number;
    sheet: GoogleAppsScript.Spreadsheet.Sheet | null;
    spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null;
};

type NotificationMessageType = {
    columnName: string;
    managersName: string;
    newCellContent: string;
    oldCellContent: string;
    rowNumber: number;
    sheetName: string;
};

type RowRangesType = {
    end: number;
    start: number;
};

// ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
const managerColumnList: Array<string> = ['A', 'B', 'C', 'E', 'G', 'H', 'I', 'L', 'N', 'O', 'P', 'Q'];
const requestsColumnList: Array<string> = ['R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y'];
const commonColumnList: Array<string> = ['D', 'F', 'J', 'K', 'M'];
const rowIdColumnName = 'BF';
const rowActionColumnName = 'A';
const nonUpdatableColumnNameList: Array<string> = [rowActionColumnName];

// first row with data
const dataRowBegin = 3;
const firstColumnName = 'A';
const lastColumnName = rowIdColumnName;

const allDataRange = `${firstColumnName}${dataRowBegin}:${lastColumnName}`;

Logger.log(lastColumnName);

const appUI: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();

const util = {
    columnStringToNumber(columnStringRaw: string): number {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const alphabetLength: number = alphabet.length;
        const list: Array<string> = [...alphabet.toUpperCase()];
        const columnString = columnStringRaw.trim().toUpperCase();

        return [...columnString].reverse().reduce<number>((sum: number, char: string, index: number) => {
            return sum + (list.indexOf(char) + 1) * Math.pow(alphabetLength, index);
        }, 0);
    },
    columnNumberToString(columnNumber: number): string {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const list: Array<string> = [...alphabet.toUpperCase()];

        return list[columnNumber - 1] || '!';
    },
    getSpreadSheetUrl(tableId: string): string {
        return SpreadsheetApp.openById(tableId).getUrl();
    },
    getSpreadSheetName(tableId: string): string {
        return SpreadsheetApp.openById(tableId).getName();
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
    getIsUpdate(row: Array<unknown>): boolean {
        const cellRawValue = util.stringify(row[rowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.update;
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
    getIsSyncedCell(rowNumber: number, columnNumber: number): boolean {
        const cell = SpreadsheetApp.getActiveSheet().getRange(rowNumber, columnNumber);
        const cellBackgroundColor = cell.getBackground();

        return cellBackgroundColor.toLowerCase() === bgColorSynced.toLowerCase();
    },
    // return invalid rows/cells
    getInvalidRequiredCellList(requiredColumnList: Array<string>): Array<string> {
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(allDataRange);
        const errorCellList: Array<string> = [];
        const requiredColumnNumberList: Array<number> = requiredColumnList.map<number>(util.columnStringToNumber);

        range.getValues().forEach((row: Array<unknown>, rowIndex: number) => {
            // check for this row should be updated
            if (!util.getIsUpdateOrAdd(row) && !util.getIsUpdate(row)) {
                return;
            }

            const rowNumber: number = rowIndex + dataRowBegin;

            row.forEach((cellValue: unknown, columnIndex: number) => {
                const columnNumber = columnIndex + 1;

                if (!requiredColumnNumberList.includes(columnNumber)) {
                    return;
                }

                if (util.stringify(cellValue) === '') {
                    errorCellList.push(`${util.columnNumberToString(columnNumber)}${rowNumber}`);
                }
            });
        });

        return errorCellList;
    },
    sendNotificationOnUpdate(messageContent: NotificationMessageType, sheetUrl: string): void {
        const {sheetName, managersName, rowNumber, columnName, oldCellContent, newCellContent} = messageContent;

        const message = `${sheetName}: Внесены изменения: ${
            managersName || 'Не указано'
        }, ${rowNumber} строка, ${columnName}:  ${oldCellContent} - ${newCellContent}`;

        UrlFetchApp.fetch(
            'https://workspace1.rocket.chat/hooks/6388be8fc6750ea9174eb2ed/L9QLKoFg3t9kEaSweTiXQA2A8dikKcfbpauXN9PstZj6fMzo',
            {
                method: 'get',
                payload: {
                    text: message,
                    title: 'SpreadSheet',
                    title_link: sheetUrl,
                    link_description: 'Ссылка на таблицу',
                },
            }
        );
    },
    sendNotificationOnAdd(range: RowRangesType, sheetUrl: string): void {
        const message = `Добавлена новая информация: строки ${range.start}-${range.end}`;

        UrlFetchApp.fetch(
            'https://workspace1.rocket.chat/hooks/6388be8fc6750ea9174eb2ed/L9QLKoFg3t9kEaSweTiXQA2A8dikKcfbpauXN9PstZj6fMzo',
            {
                method: 'get',
                payload: {
                    text: message,
                    title: 'SpreadSheet',
                    title_link: sheetUrl,
                    link_description: 'Ссылка на таблицу',
                },
            }
        );
    },
};

const rowIdColumnIndex = util.columnStringToNumber(rowIdColumnName) - 1;
const rowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
const managerColumnNumberList: Array<number> = managerColumnList.map<number>(util.columnStringToNumber);
const requestsColumnNumberList: Array<number> = requestsColumnList.map<number>(util.columnStringToNumber);
const commonColumnNumberList: Array<number> = commonColumnList.map<number>(util.columnStringToNumber);
const nonUpdatableColumnNumberList: Array<number> = nonUpdatableColumnNameList.map<number>(util.columnStringToNumber);

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
        const invalidRequiredCellList = util.getInvalidRequiredCellList(managerRequiredColumnName);

        if (invalidRequiredCellList.length > 0) {
            appUI.alert(`You should set follow cells: ${invalidRequiredCellList.join(', ')}.`);
            return;
        }

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

        const addedRows: Array<number> = [];

        // update rows

        managerTable
            .getAllDataRange()
            .getValues()
            // eslint-disable-next-line complexity
            .forEach((managerRow: Array<unknown>, managerRowIndex: number) => {
                if (!util.getIsUpdateOrAdd(managerRow)) {
                    return;
                }

                const managerRowId = util.stringify(managerRow[rowIdColumnIndex]);

                if (!managerRowId) {
                    return;
                }

                const requestsSheet = requestsTable.getRequestsSheet();
                const requestsRowData = mainTable.getRowDataById(managerRowId, [requestsTableId]);
                const hasToMakeNewLine = !requestsRowData.sheet;
                const requestsRangeRowNumber: number = hasToMakeNewLine
                    ? requestsSheet.getLastRow() + 1
                    : requestsRowData.rowNumber;
                const requestsRangeRowBgColor: string = hasToMakeNewLine ? bgColorDefault : bgColorSynced;

                if (hasToMakeNewLine) {
                    addedRows.push(managerRowIndex);
                }

                // eslint-disable-next-line complexity
                managerRow.forEach((managerRowData: unknown, managerColumnIndex: number) => {
                    const currentColumnNumber = managerColumnIndex + 1;
                    const managerRowNumber = managerRowIndex + dataRowBegin;
                    const isInManagerColumnRange = managerColumnNumberList.includes(currentColumnNumber);
                    const isInCommonColumnRange = commonColumnNumberList.includes(currentColumnNumber);
                    const isRowIdColumn = managerColumnIndex === rowIdColumnIndex;
                    const isNonUpdatableColumn = nonUpdatableColumnNumberList.includes(currentColumnNumber);

                    if (isNonUpdatableColumn || util.getIsSyncedCell(managerRowNumber, currentColumnNumber)) {
                        return;
                    }

                    if (isInManagerColumnRange || isInCommonColumnRange || isRowIdColumn) {
                        const oldCellValue = requestsSheet.getRange(requestsRangeRowNumber, currentColumnNumber);

                        if (!hasToMakeNewLine && managerRowData) {
                            const managerNameColumn = 2;
                            const sheetUrl = util.getSpreadSheetUrl(requestsTableId);
                            const managerName: string = requestsSheet
                                .getRange(requestsRangeRowNumber, managerNameColumn)
                                .getValue();
                            const columnTitle: string = requestsSheet
                                .getRange(managerNameColumn, currentColumnNumber)
                                .getValue();

                            util.sendNotificationOnUpdate(
                                {
                                    sheetName: util.getSpreadSheetName(requestsTableId),
                                    managersName: managerName,
                                    rowNumber: requestsRangeRowNumber,
                                    oldCellContent: oldCellValue.getValue() || "'Пустое поле'",
                                    newCellContent: util.stringify(managerRowData),
                                    columnName: columnTitle,
                                },
                                sheetUrl
                            );
                        }

                        oldCellValue.setValue(managerRowData).setBackground(requestsRangeRowBgColor);
                        SpreadsheetApp.getActiveSheet()
                            .getRange(managerRowNumber, currentColumnNumber)
                            .setBackground(bgColorSynced);
                    }
                });
            });

        if (addedRows.length > 0) {
            util.sendNotificationOnAdd(
                {
                    start: addedRows[0],
                    end: addedRows[addedRows.length - 1],
                },
                util.getSpreadSheetUrl(requestsTableId)
            );
        }

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
        const invalidRequiredCellList = util.getInvalidRequiredCellList(requestsRequiredColumnName);

        if (invalidRequiredCellList.length > 0) {
            appUI.alert(`You should set follow cells: ${invalidRequiredCellList.join(', ')}.`);
            return;
        }

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
            .forEach((requestsRow: Array<unknown>, requestsRowIndex: number) => {
                if (!util.getIsUpdate(requestsRow)) {
                    return;
                }

                const requestsRowId = util.stringify(requestsRow[rowIdColumnIndex]);

                if (!requestsRowId) {
                    return;
                }

                const managerRowData = mainTable.getRowDataById(requestsRowId, managerTableIdList);
                const {
                    sheet: managerSheet,
                    rowNumber: managerRowNumber,
                    spreadsheet: managerSpreadSheet,
                } = managerRowData;

                if (!managerSheet) {
                    appUI.alert(`Can not find row with id ${requestsRowId}`);
                    return;
                }

                const requestsRangeRowNumber: number = requestsRowIndex + dataRowBegin;

                // eslint-disable-next-line complexity
                requestsRow.forEach((requestsRowData: unknown, requestColumnIndex: number) => {
                    const currentColumnNumber = requestColumnIndex + 1;
                    const isInRequestsColumnRange = requestsColumnNumberList.includes(currentColumnNumber);
                    const isInCommonColumnRange = commonColumnNumberList.includes(currentColumnNumber);
                    const isNonUpdatableColumn = nonUpdatableColumnNumberList.includes(currentColumnNumber);

                    if (isNonUpdatableColumn) {
                        return;
                    }

                    if (util.getIsSyncedCell(requestsRangeRowNumber, currentColumnNumber)) {
                        return;
                    }

                    if (isInRequestsColumnRange || isInCommonColumnRange) {
                        const managerNameColumn = 2;
                        const oldCellValue = managerSheet.getRange(requestsRangeRowNumber, currentColumnNumber);
                        const managerSpreadSheetId = util.stringify(managerSpreadSheet?.getId());
                        const sheetUrl = util.getSpreadSheetUrl(managerSpreadSheetId);
                        const managerName: string = managerSheet
                            .getRange(requestsRangeRowNumber, managerNameColumn)
                            .getValue();
                        const columnTitle: string = managerSheet
                            .getRange(managerNameColumn, currentColumnNumber)
                            .getValue();

                        util.sendNotificationOnUpdate(
                            {
                                sheetName: util.getSpreadSheetName(managerSpreadSheetId),
                                managersName: managerName,
                                rowNumber: requestsRangeRowNumber,
                                oldCellContent: oldCellValue.getValue() || "'Пустое поле'",
                                newCellContent: util.stringify(requestsRowData),
                                columnName: columnTitle,
                            },
                            sheetUrl
                        );

                        oldCellValue.setValue(requestsRowData).setBackground(bgColorSynced);
                        SpreadsheetApp.getActiveSheet()
                            .getRange(requestsRangeRowNumber, currentColumnNumber)
                            .setBackground(bgColorSynced);
                    }
                });
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 3/3', 'Synced!', 2);
    },
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

type OnEditEventType = {range: GoogleAppsScript.Spreadsheet.Range; value: unknown};

function onEdit(changeData: OnEditEventType) {
    changeData.range.setBackground(bgColorChanged);
    SpreadsheetApp.getActiveSheet()
        .getRange(`${rowActionColumnName}${dataRowBegin}:${rowActionColumnName}`)
        .setBackground(bgColorDefault);
    mainTable.updateRowsId();
}
