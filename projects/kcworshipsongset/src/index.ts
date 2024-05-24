const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

let TOTAL_NONEMPTY_ROWS = 2;
const MONTH_ROW_COUNT = 5;

let range: GoogleAppsScript.Spreadsheet.Range;
let sheet: GoogleAppsScript.Spreadsheet.Sheet;
const RANGES_TO_MERGE: GoogleAppsScript.Spreadsheet.Range[] = [];

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function main(year = 2024, monthIndex = 1): void {
  monthIndex -= 1;
  const month = new Intl.DateTimeFormat('en', { month: 'long' }).format(
    new Date().setMonth(monthIndex),
  );
  const sheetName = `${month} ${year}`;

  createSheet(sheetName);
  addColumns();
  populateSheet(year, monthIndex);
  formatSheet();
  mergeCells();
}

function addColumns(): void {
  console.info('Adding columns...');
  sheet.getRange(1, 1).setValue('Date');
  sheet.getRange(1, 2).setValue('Leader');
  sheet.getRange(1, 3).setValue('Songs');
  sheet.getRange(1, 4).setValue('Artist');
  sheet.getRange(1, 5).setValue('Key');
  sheet.getRange(1, 6).setValue('BPM');
  sheet.getRange(1, 7).setValue('Reference');
  sheet.getRange(1, 8).setValue('Comments');
}

function populateSheet(year: number, monthIndex: number): void {
  console.info('Populating sheet...');

  const sundaysInMonth = getSundaysInMonth(year, monthIndex);
  const colours = Object.values(COLOURS)
    .sort(() => 0.5 - Math.random())
    .slice(0, sundaysInMonth.length);

  sundaysInMonth.forEach((date, i) => {
    range = sheet
      .getRange(TOTAL_NONEMPTY_ROWS, 1, MONTH_ROW_COUNT, 1)
      .setValue(formatOrdinal(date));
    RANGES_TO_MERGE.push(range);
    range = sheet.getRange(TOTAL_NONEMPTY_ROWS, 2, MONTH_ROW_COUNT, 1);
    RANGES_TO_MERGE.push(range);
    sheet
      .getRange(TOTAL_NONEMPTY_ROWS, 1, MONTH_ROW_COUNT, TOTAL_NONEMPTY_COLUMNS)
      .setBackground(colours[i]);
    TOTAL_NONEMPTY_ROWS += MONTH_ROW_COUNT;
  });

  console.info('Sheet populated.');
}

function formatSheet(): void {
  console.info('Formatting sheet...');

  // Style column headers.
  sheet
    .getRange(1, 1, 1, TOTAL_NONEMPTY_COLUMNS)
    .setBackground('#666666')
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Nunito')
        .setFontSize(12)
        .setForegroundColor('#FFFFFF')
        .build(),
    )
    .setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, 6).setHorizontalAlignment('center');
  sheet.getRange(1, 7, 1, 2).setHorizontalAlignment('left');

  // Style first column
  sheet
    .getRange(2, 1, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Nunito')
        .setFontSize(18)
        .build(),
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style second column
  sheet
    .getRange(2, 2, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Nunito')
        .setFontSize(14)
        .build(),
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style third column
  sheet
    .getRange(2, 3, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Nunito')
        .setFontSize(12)
        .build(),
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style fourth & fifth column
  sheet
    .getRange(2, 4, TOTAL_NONEMPTY_ROWS, 2)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setFontFamily('Nunito')
        .setFontSize(12)
        .build(),
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style sixth column
  sheet
    .getRange(2, 6, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setBold(true)
        .setFontFamily('Nunito')
        .setFontSize(12)
        .build(),
    )
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style seventh column
  sheet
    .getRange(2, 7, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setFontFamily('Nunito')
        .setFontSize(12)
        .build(),
    )
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  // Style eighth column
  sheet
    .getRange(2, 8, TOTAL_NONEMPTY_ROWS, 1)
    .setTextStyle(
      SpreadsheetApp.newTextStyle()
        .setFontFamily('Nunito')
        .setFontSize(12)
        .build(),
    )
    .setVerticalAlignment('middle');

  console.info('Resizing rows and columns...');
  sheet.setRowHeight(1, 45);
  sheet.setRowHeights(2, TOTAL_NONEMPTY_ROWS - 1, 28);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 290);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 70);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 200);
  sheet.setColumnWidth(8, 375);
}

/**
 * Creates the sheet, deleting it if a sheet with the name already exists.
 * @returns The sheet.
 */
function createSheet(sheetName: string): void {
  sheet = spreadsheet.getSheetByName(sheetName)!;
  if (sheet) {
    console.info(`Sheet '${sheet.getSheetName()}' exists. Clearing...`);
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet();
    sheet.setName(sheetName);
    console.info(`Sheet '${sheet.getSheetName()}' created.`);
  }
}

function mergeCells(): void {
  for (const range of RANGES_TO_MERGE) {
    range.merge();
  }
}
