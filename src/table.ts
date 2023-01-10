/* global Logger, SpreadsheetApp, GoogleAppsScript, UrlFetchApp, Utilities */

// ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

// main constants
const requestsTableId = '1qmpdaS_EWJJd-C8ntLhhQ-AqrCVuI4ahBFnP0gx9ZT8';
const requestsColumnList: Array<string> = ['G', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB'];
const requestsRequiredColumnName: Array<string> = ['K', 'T', 'U', 'V', 'W', 'X'];
const requestsSheetName = 'Общая статистика';

const managerTable1Id = '1mD4Fxu1r4_lVvZ4h5mZ1SorsKW1dsNBkkQoOZn0P1o0';
const managerColumnList: Array<string> = ['A', 'B', 'C', 'D', 'E', 'H', 'I', 'J', 'L', 'N', 'P', 'Q', 'R', 'S'];
const managerRequiredColumnName: Array<string> = ['B', 'C', 'D', 'I', 'J', 'O', 'P', 'Q', 'S'];
const managerTableIdList: Array<string> = [managerTable1Id];

const commonColumnList: Array<string> = ['F', 'K', 'O']; //
const rowIdColumnName = 'BF';
const rowActionColumnName = 'A';
const nonUpdatableColumnNameList: Array<string> = [rowActionColumnName];

// first row with data
const dataRowBegin = 3;
const firstColumnName = 'A';
const lastColumnName = rowIdColumnName;

const bgColorSynced = '#00D100';
const bgColorChanged = '#ECF87F';
const bgColorDefault = '#FFFFFF';

const moscowGmt = 'GMT+3';

const allDataRange = `${firstColumnName}${dataRowBegin}:${lastColumnName}`;

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

type UpdatedCellType = {
    columnName: string;
    newCellContent: string;
    oldCellContent: string;
};

type NotificationMessageType = {
    managerName: string;
    rowNumber: number;
    sheetName: string;
    updatedCells: Array<UpdatedCellType>;
};

type RowRangesType = {
    end: number;
    start: number;
};

Logger.log(lastColumnName);

const appUI: GoogleAppsScript.Base.Ui = SpreadsheetApp.getUi();

const util = {
    columnNumberToString(columnNumber: number): string {
        const alphabet = 'abcdefghijklmnopqrstuvwxyz';
        const list: Array<string> = [...alphabet.toUpperCase()];

        if (columnNumber < 1) {
            appUI.alert('[columnNumberToString] value should be 1 or more');
            throw new Error('[columnNumberToString] value should be 1 or more');
        }

        if (columnNumber >= 702) {
            appUI.alert('[columnNumberToString] value should be less then 702');
            throw new Error('[columnNumberToString] value should be less then 702');
        }

        if (columnNumber > 26) {
            const firstNumber = Math.floor(columnNumber / 26);

            return list[firstNumber - 1] + util.columnNumberToString(columnNumber % 26);
        }

        return list[columnNumber - 1];
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
    getInvalidRequiredCellList(requiredColumnList: Array<string>): Array<string> {
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(allDataRange);
        const errorCellList: Array<string> = [];
        // eslint-disable-next-line unicorn/prefer-set-has
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
    getIsManagerSpreadsheet(): boolean {
        const spreadsheetId = SpreadsheetApp.getActive().getId();

        return managerTableIdList.includes(spreadsheetId);
    },
    getIsRemoveRow(row: Array<unknown>): boolean {
        const utilRowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
        const cellRawValue = util.stringify(row[utilRowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.remove;
    },
    getIsRequestsSpreadsheet(): boolean {
        const spreadsheetId = SpreadsheetApp.getActive().getId();

        return spreadsheetId === requestsTableId;
    },
    getIsSkipRow(row: Array<unknown>): boolean {
        const utilRowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
        const cellRawValue = util.stringify(row[utilRowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.skip;
    },
    getIsSyncedCell(rowNumber: number, columnNumber: number): boolean {
        const cell = SpreadsheetApp.getActiveSheet().getRange(rowNumber, columnNumber);
        const cellBackgroundColor = cell.getBackground();

        return cellBackgroundColor.toLowerCase() === bgColorSynced.toLowerCase();
    },
    getIsUpdate(row: Array<unknown>): boolean {
        const utilRowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
        const cellRawValue = util.stringify(row[utilRowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.update;
    },
    getIsUpdateOrAdd(row: Array<unknown>): boolean {
        const utilRowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
        const cellRawValue = util.stringify(row[utilRowActionColumnIndex]).trim().toLowerCase();

        return cellRawValue === RowActionNameEnum.updateOrAdd;
    },
    getRandomString(): string {
        const fromRandom = Math.random().toString(32).replace('0.', '');
        const fromTime = Date.now().toString(32);

        return `${fromRandom}${fromTime}`.toLowerCase();
    },
    getSpreadSheetName(tableId: string): string {
        return SpreadsheetApp.openById(tableId).getName();
    },
    getSpreadSheetUrl(tableId: string): string {
        return SpreadsheetApp.openById(tableId).getUrl();
    },
    // return invalid rows/cells
    sendNotification(notification: Record<'message' | 'sheetUrl', string>): void {
        UrlFetchApp.fetch(
            'https://chat.sigirgroup.com/hooks/63a42d41681796c1e50bd542/xxy6a8PxCCEEwom7HPcoQuJhHp953CQkDvQaMBeZideT7sJF',
            {
                method: 'get',
                payload: {
                    // eslint-disable-next-line id-match, camelcase
                    link_description: 'Ссылка на таблицу',
                    text: notification.message,
                    title: 'SpreadSheet',
                    // eslint-disable-next-line id-match, camelcase
                    title_link: notification.sheetUrl,
                },
            }
        );
    },
    sendNotificationOnAdd(range: RowRangesType, sheetUrl: string): void {
        const message = `Добавлена новая информация: строки ${range.start}-${range.end}`;

        util.sendNotification({message, sheetUrl});
    },
    sendNotificationOnUpdate(updates: Array<NotificationMessageType>, sheetUrl: string): void {
        const message = updates
            .map(
                ({sheetName, managerName, rowNumber, updatedCells}, index: number) =>
                    `${index + 1}) ${sheetName}: Внесены изменения: ${
                        managerName || 'Не указано'
                    }, ${rowNumber} строка, ` +
                    updatedCells
                        .filter(cells => cells.newCellContent && cells.oldCellContent)
                        .map(
                            ({columnName, newCellContent, oldCellContent}) =>
                                `${columnName}:  ${oldCellContent} - ${newCellContent}`
                        )
                        .join('; ')
            )
            .join('\n');

        util.sendNotification({message, sheetUrl});
    },
    sendNotificationOndDelete(
        sheetInfo: {sheetName: string; sheetUrl: string},
        deletedCells: Array<string>,
        rowNumber: number
    ): void {
        const message = `${sheetInfo.sheetName}: **УДАЛЕНО**: ${rowNumber} строка: ${deletedCells.join('; ')}`;

        util.sendNotification({message, sheetUrl: sheetInfo.sheetUrl});
    },
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

const rowIdColumnIndex = util.columnStringToNumber(rowIdColumnName) - 1;
const rowActionColumnIndex = util.columnStringToNumber(rowActionColumnName) - 1;
// eslint-disable-next-line unicorn/prefer-set-has
const managerColumnNumberList: Array<number> = managerColumnList.map<number>(util.columnStringToNumber);
// eslint-disable-next-line unicorn/prefer-set-has
const requestsColumnNumberList: Array<number> = requestsColumnList.map<number>(util.columnStringToNumber);
// eslint-disable-next-line unicorn/prefer-set-has
const commonColumnNumberList: Array<number> = commonColumnList.map<number>(util.columnStringToNumber);
// eslint-disable-next-line unicorn/prefer-set-has
const nonUpdatableColumnNumberList: Array<number> = nonUpdatableColumnNameList.map<number>(util.columnStringToNumber);

const mainTable = {
    getRowDataById(searchingRowId: string, tableIdList: Array<string>): RowDataType {
        let rowData: RowDataType = {
            rowDataList: [],
            rowId: '',
            rowNumber: 0,
            sheet: null,
            spreadsheet: null,
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
                        rowDataList: [...row],
                        rowId,
                        rowNumber: rowIndex + dataRowBegin,
                        sheet,
                        spreadsheet,
                    };

                    return true;
                });
            });
        });

        return rowData;
    },
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
        const syncingText = 'Syncing...';
        const invalidRequiredCellList = util.getInvalidRequiredCellList(requestsRequiredColumnName);

        if (invalidRequiredCellList.length > 0) {
            appUI.alert(`You should set follow cells: ${invalidRequiredCellList.join(', ')}.`);
            return;
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 0/3', syncingText, -1);

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

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 1/3', syncingText, -1);

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

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 2/3', syncingText, -1);

        const updatedRows: Array<NotificationMessageType> = [];
        let spreadSheetId = '';
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
                    spreadsheet: managerSpreadSheet,
                    rowNumber: managerRowNumber,
                } = managerRowData;

                if (!managerSheet) {
                    appUI.alert(`Can not find row with id ${requestsRowId}`);
                    return;
                }

                const requestsRangeRowNumber: number = requestsRowIndex + dataRowBegin;

                const managerNameColumn = 2;
                const managerSpreadSheetId = util.stringify(managerSpreadSheet?.getId());

                spreadSheetId = managerSpreadSheetId;

                const updatedRow: NotificationMessageType = {
                    managerName: managerSheet.getRange(requestsRangeRowNumber, managerNameColumn).getValue(),
                    rowNumber: requestsRangeRowNumber,
                    sheetName: util.getSpreadSheetName(managerSpreadSheetId),
                    updatedCells: [],
                };

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
                        const oldCellValue = managerSheet.getRange(managerRowNumber, currentColumnNumber);

                        updatedRow.updatedCells.push({
                            columnName: managerSheet.getRange(managerNameColumn, currentColumnNumber).getValue(),
                            newCellContent: util.stringify(requestsRowData),
                            oldCellContent: oldCellValue.getValue() || '\'Пустое поле\'',
                        });

                        oldCellValue.setValue(requestsRowData).setBackground(bgColorSynced);
                        SpreadsheetApp.getActiveSheet()
                            .getRange(requestsRangeRowNumber, currentColumnNumber)
                            .setBackground(bgColorSynced);
                    }
                });
                if (updatedRow.updatedCells.length > 0) {
                    updatedRows.push(updatedRow);
                }
            });

        if (updatedRows.length > 0) {
            util.sendNotificationOnUpdate(updatedRows, util.getSpreadSheetUrl(spreadSheetId));
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 3/3', 'Synced!', 2);
    },
};

const managerTable = {
    getAllDataRange(): GoogleAppsScript.Spreadsheet.Range {
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = SpreadsheetApp.getActiveSheet();

        return sheet.getRange(allDataRange);
    },

    mainDateFormat: 'dd-MM-yyyy',

    makeUiMenuForManager() {
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data to requests table');

        menu.addItem('Push to Requests Table', 'managerTable.pushDataToRequestTable');
        menu.addToUi();
    },

    // eslint-disable-next-line sonarjs/cognitive-complexity
    pushDataToRequestTable() {
        const syncingText = 'Syncing...';
        const invalidRequiredCellList = util.getInvalidRequiredCellList(managerRequiredColumnName);

        if (invalidRequiredCellList.length > 0) {
            appUI.alert(`You should set follow cells: ${invalidRequiredCellList.join(', ')}.`);
            return;
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 0/3', syncingText, -1);

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

                const {sheet, rowNumber} = mainTable.getRowDataById(managerRowId, [requestsTableId]);
                const deletedCells: Array<string> = [];

                managerRow.forEach((managerRowData: unknown, managerColumnIndex: number) => {
                    const currentColumnNumber = managerColumnIndex + 1;
                    const cell = sheet?.getRange(rowNumber, currentColumnNumber);

                    const columnLetter = util.columnNumberToString(currentColumnNumber);
                    const deletedColumns = ['D', 'M', 'N', 'S', 'V'];
                    let cellValue: Date | string = cell?.getValue();

                    if (cellValue instanceof Date) {
                        cellValue = Utilities.formatDate(cellValue, moscowGmt, managerTable.mainDateFormat);
                    }

                    if (deletedColumns.includes(columnLetter)) {
                        deletedCells.push(cellValue || 'Не указано');
                    }
                });

                util.sendNotificationOndDelete(
                    {
                        sheetName: util.getSpreadSheetName(requestsTableId),
                        sheetUrl: util.getSpreadSheetUrl(requestsTableId),
                    },
                    deletedCells,
                    rowNumber
                );

                sheet?.deleteRow(rowNumber);
            });

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 1/3', syncingText, -1);

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

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 2/3', syncingText, -1);

        const addedRows: Array<number> = [];
        const updatedRows: Array<NotificationMessageType> = [];

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

                const managerNameColumn = 2;

                const updatedRow: NotificationMessageType = {
                    managerName: requestsSheet.getRange(requestsRangeRowNumber, managerNameColumn).getValue(),
                    rowNumber: requestsRangeRowNumber,
                    sheetName: util.getSpreadSheetName(requestsTableId),
                    updatedCells: [],
                };

                if (hasToMakeNewLine) {
                    addedRows.push(requestsRangeRowNumber);
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

                        let oldValue: Date | string = oldCellValue.getValue();
                        let newValue = util.stringify(managerRowData);

                        if (oldValue instanceof Date) {
                            oldValue = Utilities.formatDate(oldValue, moscowGmt, managerTable.mainDateFormat);
                        }

                        if (managerRowData instanceof Date) {
                            newValue = Utilities.formatDate(managerRowData, moscowGmt, managerTable.mainDateFormat);
                        }

                        if (!hasToMakeNewLine && managerRowData) {
                            updatedRow.updatedCells.push({
                                columnName: requestsSheet.getRange(managerNameColumn, currentColumnNumber).getValue(),
                                newCellContent: newValue,
                                oldCellContent: oldValue,
                            });
                        }

                        oldCellValue.setValue(managerRowData).setBackground(requestsRangeRowBgColor);
                        SpreadsheetApp.getActiveSheet()
                            .getRange(managerRowNumber, currentColumnNumber)
                            .setBackground(bgColorSynced);
                    }
                });

                if (updatedRow.updatedCells.length > 0) {
                    updatedRows.push(updatedRow);
                }
            });

        if (addedRows.length > 0) {
            util.sendNotificationOnAdd(
                {
                    end: addedRows[addedRows.length - 1],
                    start: addedRows[0],
                },
                util.getSpreadSheetUrl(requestsTableId)
            );
        }

        if (updatedRows.length > 0) {
            util.sendNotificationOnUpdate(updatedRows, util.getSpreadSheetUrl(requestsTableId));
        }

        SpreadsheetApp.getActiveSpreadsheet().toast('Done: 3/3', 'Synced!', 2);
    },
};

// will call on document open
// eslint-disable-next-line @typescript-eslint/no-unused-vars, no-unused-vars,
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

// eslint-disable-next-line @typescript-eslint/no-unused-vars, no-unused-vars,
function onEdit(changeData: OnEditEventType) {
    changeData.range.setBackground(bgColorChanged);
    SpreadsheetApp.getActiveSheet()
        .getRange(`${rowActionColumnName}${dataRowBegin}:${rowActionColumnName}`)
        .setBackground(bgColorDefault);
    mainTable.updateRowsId();
}
