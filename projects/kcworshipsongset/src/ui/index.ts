function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Admin')
    .addItem('Create Sheet...', 'showPrompt')
    .addToUi();
}

function showPrompt(): void {
  const ui = SpreadsheetApp.getUi();
  let monthIndex = 0;
  let year = 0;

  try {
    const yearPrompt = ui.prompt('Which year?');
    if (yearPrompt.getSelectedButton() !== ui.Button.OK) return;

    try {
      year = parseInt(yearPrompt.getResponseText());
    } catch (e) {
      throw new Error('Please specify a valid year.');
    }

    const termPrompt = ui.prompt('Which month number of the year? (1-12)');
    if (termPrompt.getSelectedButton() !== ui.Button.OK) return;

    try {
      monthIndex = parseInt(termPrompt.getResponseText());
    } catch (e) {
      throw new Error('Please specify a number between 1 and 12.');
    }

    if (monthIndex < 1 || monthIndex > 12) {
      throw new Error('Invalid month. Specify a number between 1 and 12.');
    }
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  main(year, monthIndex);
}
