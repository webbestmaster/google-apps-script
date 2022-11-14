/* global Logger, SpreadsheetApp, GoogleAppsScript */

const requestTableId = '1E5BIjJ6cpFsSl9fVcBvt9x8Bk7-uhqrz64jFbmqg5GI';
const managerTableId1 = '11ZNH5S8DuZUobQU6sw-_Svx5vVt62I9CmxSSF_eGFDM';
const managerTableId2 = '1C3pU0hsaGZnsztX72ZND6WRsHBA5EyY5ozEilL2CrOA';
const managerDataRange = 'A3:D';
// const requestDataRange = 'E3:H';

type ConnectTableConfigType = Readonly<{
    managerTableIdList: Array<string>;
    requestTableId: string;
}>;

// eslint-disable-next-line unicorn/no-static-only-class
class ConnectTable {
    static managerTableIdList: Array<string> = [];
    static requestTableId = '';

    static initialize(connectTableConfig: ConnectTableConfigType) {
        ConnectTable.requestTableId = connectTableConfig.requestTableId;
        ConnectTable.managerTableIdList = connectTableConfig.managerTableIdList;
    }

    static pushDataToRequestTable() {
        const spreadsheetApp: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActive();
        const sheet: GoogleAppsScript.Spreadsheet.Sheet = spreadsheetApp.getActiveSheet();
        const range: GoogleAppsScript.Spreadsheet.Range = sheet.getRange(managerDataRange);

        range.getValues().forEach((row: unknown) => {
            Logger.log(JSON.stringify(row));
        });

        Logger.log(sheet.getSheetId());
        Logger.log(sheet.getSheetName());

        range.sort({ascending: true, column: 2});

        const anotherSheet: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.openById(managerTableId1);
        Logger.log(anotherSheet.getSheetName());
        // const appUI = SpreadsheetApp.getUi();
        // appUI.alert("pushDataToRequestTable");
    }

    static pushDataToManagerTable() {
        // const appUI = SpreadsheetApp.getUi();
        // appUI.alert("pushDataToManagerTable");
    }

    static makeUiMenu() {
        const appUI = SpreadsheetApp.getUi();
        const menu: GoogleAppsScript.Base.Menu = appUI.createMenu('Push data');

        menu.addItem('push Data To Request Table', 'ConnectTable.pushDataToRequestTable');
        menu.addItem('push Data To Manager Table', 'ConnectTable.pushDataToManagerTable');
        menu.addToUi();
    }
}

ConnectTable.initialize({
    managerTableIdList: [managerTableId1, managerTableId2],
    requestTableId,
});

// will call on document open
// eslint-disable-next-line no-unused-vars, @typescript-eslint/no-unused-vars
function onOpen() {
    ConnectTable.makeUiMenu();
    ConnectTable.pushDataToRequestTable();

    Logger.log('will call on document open!!!');
}
