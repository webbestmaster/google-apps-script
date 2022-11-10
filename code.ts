const greeter = (person: string) => {
    return `Hello, ${person}!`;
}

const user = 'Grant';
Logger.log(greeter(user));

const ui = SpreadsheetApp.getUi();
const response = ui.prompt('Getting to know you', 'May I know your name?', ui.ButtonSet.YES_NO);

if (response.getSelectedButton() == ui.Button.YES) {
    Logger.log('The user\'s name is %s.', response.getResponseText());
} else if (response.getSelectedButton() == ui.Button.NO) {
    Logger.log('The user didn\'t want to provide a name.');
} else {
    Logger.log('The user clicked the close button in the dialog\'s title bar.');
}
