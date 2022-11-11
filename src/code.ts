/* global Logger, SpreadsheetApp, GoogleAppsScript */

export function greeter(person: string): string {
    return `Hello, ${person}!`;
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
