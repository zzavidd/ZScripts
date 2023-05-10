function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Custom Menu')
    .addItem('Create Sheet', 'main')
    .addToUi();
}

function showPrompt(): void {
  const html = HtmlService.createTemplateFromFile('prompt')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Prompt');
}
