function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Admin')
    .addItem('Create Sheet...', 'showPrompt')
    .addToUi();
}

function showPrompt(): void {
  const ui = SpreadsheetApp.getUi();
  let term;
  let year;

  try {
    const yearPrompt = ui.prompt('Which year?');
    if (yearPrompt.getSelectedButton() !== ui.Button.OK) return;

    try {
      year = parseInt(yearPrompt.getResponseText());
    } catch (e) {
      throw new Error('Please specify a valid year.');
    }

    const termPrompt = ui.prompt(
      'What term?\n\n(1) Jan-Feb\n(2) Mar-Apr-May\n(3) Jun-Jul-Aug\n(4) Sep-Oct-Nov-Dec\n',
    );
    if (termPrompt.getSelectedButton() !== ui.Button.OK) return;

    try {
      term = parseInt(termPrompt.getResponseText());
    } catch (e) {
      throw new Error('Please specify a number between 1 and 4.');
    }

    if (term < 1 || term > 4) {
      throw new Error('Invalid term. Only terms 1-4 are allowed.');
    }
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  main(term, year);
}
