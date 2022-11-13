/* global Logger, SpreadsheetApp, GoogleAppsScript */

export function greeter(person: string): string {
    return `Hello, ${person}!!!`;
}

const user = 'Grant';

Logger.log(greeter(user));

const appUI = SpreadsheetApp.getUi();
const response: GoogleAppsScript.Base.PromptResponse = appUI.prompt(
    'Getting to know you',
    'May I know your name?',
    appUI.ButtonSet.YES_NO
);

if (response.getSelectedButton() === appUI.Button.YES) {
    Logger.log("The user's name is %s.", response.getResponseText());
} else if (response.getSelectedButton() === appUI.Button.NO) {
    Logger.log("The user didn't want to provide a name.");
} else {
    Logger.log("The user clicked the close button in the dialog's title bar.");
}

// not tested - set cell's value
SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("a1").setValue("w   ");

// not tested - add UI
const appUI2 = SpreadsheetApp.getUi();
const menu: GoogleAppsScript.Base.Menu = appUI2.createMenu("My menu")
function onButtonClick() {
    Logger.log("onButtonClick!!!")
}

appUI2.alert("appUI2.alert");

menu.addItem("My button", "onButtonClick");
menu.addToUi();

// will call on document open
function onOpen() {
    Logger.log("will call on document open!!!")
}

class Person {
    firstName: string;
    lastName: string;
    static age: number = 0;

    constructor(firstName: string, lastName: string, age: number) {
        this.firstName = firstName;
        this.lastName = lastName;
        Person.age = age;
    }
}

const newP = new Person("dd", "ssd", 11);
console.log(Person.age)
