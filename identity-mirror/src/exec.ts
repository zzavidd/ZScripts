const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const PARTICIPANTS = ['Ama', 'Grace', 'Michael', 'Zavid'];
const TOTAL_COLUMNS = PARTICIPANTS.length + 1;
const TOTAL_ROWS = ADJECTIVES.length + 1;

const TEXT_STYLE = SpreadsheetApp.newTextStyle()
  .setFontFamily('Nunito')
  .setFontSize(12)
  .build();

const HEADER_COLUMN_COLOR = '#e1effb';
const HEADER_ROW_COLOR = '#0b5394';

function createSheets() {
  PARTICIPANTS.forEach((participant) => {
    const pollSheetName = `Poll: ${participant}`;

    try {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(pollSheetName)!);
    } catch (e: any) {
      console.warn(e.message);
    }

    const pollSheet = spreadsheet.insertSheet();
    pollSheet.setName(pollSheetName);
    const pollRange = pollSheet.getRange(1, 1, TOTAL_ROWS, TOTAL_COLUMNS);

    formatRange(pollSheet);

    pollSheet
      .getRange(2, 1, TOTAL_ROWS - 1, 1)
      .setBackground(HEADER_COLUMN_COLOR);

    // Set data validation
    pollSheet
      .getRange(2, 2, TOTAL_ROWS - 1, TOTAL_COLUMNS - 1)
      .setDataValidation(VALIDATION);

    pollRange.getCell(1, 2).setValue('You');

    const otherParticipants = [...PARTICIPANTS];
    otherParticipants.splice(otherParticipants.indexOf(participant), 1);
    otherParticipants.forEach((p, i) => {
      pollRange.getCell(1, i + 3).setValue(p);
    });

    // List out adjectives in first column
    ADJECTIVES.forEach((adjective, i) => {
      pollSheet.getRange(i + 2, 1).setValue(adjective.word);
      PARTICIPANTS.forEach((p, j) => {
        pollRange.getCell(i + 2, j + 2).setValue('None');
      });
    });
  });
}

function formatRange(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  sheet
    .getRange(1, 1, TOTAL_ROWS, TOTAL_COLUMNS + 1)
    .setTextStyle(TEXT_STYLE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const rules = getConditionalFormatRules(
    sheet.getRange(1, 1, TOTAL_ROWS, TOTAL_COLUMNS),
  );
  sheet.setConditionalFormatRules(rules);

  sheet
    .getRange(1, 1, 1, TOTAL_COLUMNS)
    .setBackground(HEADER_ROW_COLOR)
    .setFontColor('#fff')
    .setFontWeight('bold');
  sheet
    .setColumnWidth(1, 150)
    .setColumnWidths(2, TOTAL_COLUMNS - 1, 115)
    .setRowHeights(1, TOTAL_ROWS, 35);
}

function getConditionalFormatRules(range: GoogleAppsScript.Spreadsheet.Range) {
  const positive = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Yes')
    .setBackground('#d4edbc')
    .setFontColor('#11734b')
    .setRanges([range])
    .build();

  const neutral = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Not sure')
    .setBackground('#ffe5a0')
    .setFontColor('#473821')
    .setRanges([range])
    .build();

  const negative = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('No')
    .setBackground('#ffcfc9')
    .setFontColor('#b10202')
    .setRanges([range])
    .build();
  return [positive, neutral, negative];
}
