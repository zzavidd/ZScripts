const TERM: Term = 3;
const YEAR = 2023;

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const termMonths = displayMonthsForTerm(TERM);
const SHEET_NAME = `T${TERM} ${YEAR} (${termMonths})`;

const { SOLID, SOLID_MEDIUM } = SpreadsheetApp.BorderStyle;

let TOTAL_NONEMPTY_ROWS = 2;
let range: GoogleAppsScript.Spreadsheet.Range;
let sheet: GoogleAppsScript.Spreadsheet.Sheet;
const RANGES_TO_MERGE: GoogleAppsScript.Spreadsheet.Range[] = [];

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function main(): void {
  createSheet();
  addColumns();
  populateSheet();
  formatSheet();
  mergeCells();
  addDataValidation();
}

function addColumns(): void {
  console.info(`Adding columns...`);

  range = sheet.getRange(1, 1, 2, 2).setValue('Date');
  RANGES_TO_MERGE.push(range);

  range = sheet.getRange(1, 3, 1, 4).setValue('SINGERS');
  RANGES_TO_MERGE.push(range);

  range = sheet.getRange(1, 7, 1, 6).setValue('INSTRUMENTALISTS');
  RANGES_TO_MERGE.push(range);

  sheet.getRange(2, 3).setValue('Worship Leader');
  sheet.getRange(2, 4).setValue('BV1');
  sheet.getRange(2, 5).setValue('BV2');
  sheet.getRange(2, 6).setValue('BV3');

  sheet.getRange(2, 7).setValue('Keyboard');
  sheet.getRange(2, 8).setValue('Drums');
  sheet.getRange(2, 9).setValue('Bass');
  sheet.getRange(2, 10).setValue('Acoustic');
  sheet.getRange(2, 11).setValue('Saxophone');
  sheet.getRange(2, 12).setValue('Violin');
}

function populateSheet(): void {
  console.info(`Populating sheet...`);
  const colours = Object.values(COLOURS)
    .sort(() => 0.5 - Math.random())
    .slice(0, UF_TERMS[TERM].length);

  const { sundaysInTerm, numberOfSundays } = getSundaysInTerm();
  TOTAL_NONEMPTY_ROWS += numberOfSundays;

  let rowIndex = 3;
  Object.entries(sundaysInTerm).forEach(([month, dates], i) => {
    range = sheet.getRange(rowIndex, 1, dates.length, 1).setValue(month);
    RANGES_TO_MERGE.push(range);
    sheet
      .getRange(rowIndex, 1, dates.length, TOTAL_NONEMPTY_COLUMNS)
      .setBackground(colours[i]);
    dates.forEach((date) => {
      sheet.getRange(rowIndex, 2).setValue(formatOrdinal(date));
      rowIndex++;
    });
    sheet
      .getRange(rowIndex - 1, 1, 1, TOTAL_NONEMPTY_COLUMNS)
      .setBorder(null, null, true, null, null, null, '#000000', SOLID);
  });

  console.info('Sheet populated.');
}

function formatSheet(): void {
  console.info('Formatting sheet...');

  // Style column headers.
  sheet
    .getRange(1, 1, 2, TOTAL_NONEMPTY_COLUMNS)
    .setBackground('#666666')
    .setTextStyle(COLUMN_HEADER_TEXT_STYLE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style row headers.
  sheet
    .getRange(1, 1, TOTAL_NONEMPTY_ROWS, 2)
    .setTextStyle(TEXT_STYLE)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  // Style name cells
  sheet
    .getRange(3, 3, TOTAL_NONEMPTY_ROWS - 2, TOTAL_NONEMPTY_COLUMNS - 2)
    .setTextStyle(TEXT_STYLE)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  // Bottom border for first row.
  sheet
    .getRange(1, 3, 1, TOTAL_NONEMPTY_COLUMNS - 2)
    .setBorder(null, null, true, null, null, true, '#000000', SOLID);
  // Bottom border for second row.
  sheet
    .getRange(2, 1, 1, TOTAL_NONEMPTY_COLUMNS)
    .setBorder(null, null, true, null, null, null, '#000000', SOLID_MEDIUM);
  // Right border for Date columns.
  sheet
    .getRange(1, 2, TOTAL_NONEMPTY_ROWS, 1)
    .setBorder(null, null, null, true, true, null, '#000000', SOLID_MEDIUM);
  // Right border for SINGERS columns.
  sheet
    .getRange(1, 6, TOTAL_NONEMPTY_ROWS, 1)
    .setBorder(null, null, null, true, null, null, '#000000', SOLID_MEDIUM);
  // Right border for whole grid.
  sheet
    .getRange(1, TOTAL_NONEMPTY_COLUMNS, TOTAL_NONEMPTY_ROWS, 1)
    .setBorder(null, null, null, true, null, null, '#000000', SOLID_MEDIUM);
  // Bottom border for whole grid.
  sheet
    .getRange(TOTAL_NONEMPTY_ROWS, 1, 1, TOTAL_NONEMPTY_COLUMNS)
    .setBorder(null, null, true, null, null, null, '#000000', SOLID);

  console.info(`Resizing rows and columns...`);
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 40);
  sheet.setRowHeights(3, TOTAL_NONEMPTY_ROWS - 2, 40);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 94);
  sheet.setColumnWidth(3, 122);
  sheet.setColumnWidths(4, 6, 115);
  sheet.setColumnWidths(7, TOTAL_NONEMPTY_COLUMNS - 7, 100);
}

/**
 * Creates the sheet, deleting it if a sheet witht he name already exists.
 * @returns The sheet.
 */
function createSheet(): void {
  sheet = spreadsheet.getSheetByName(SHEET_NAME)!;
  if (sheet) {
    console.info(`Sheet '${sheet.getSheetName()}' exists. Clearing...`);
    sheet.clear();
  } else {
    sheet = spreadsheet.insertSheet();
    sheet.setName(SHEET_NAME);
    console.info(`Sheet '${sheet.getSheetName()}' created.`);
  }
}

function mergeCells(): void {
  RANGES_TO_MERGE.forEach((range) => range.merge());
}

function addDataValidation(): void {
  const set = (
    startColumn: number,
    list: string[],
    numberOfColumns = 1,
  ): void => {
    sheet
      .getRange(3, startColumn, TOTAL_NONEMPTY_ROWS - 2, numberOfColumns)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(list, true)
          .build(),
      );
  };

  set(3, LEAD_SINGERS);
  set(4, BV_SINGERS, 3);
  set(7, KEYBOARDISTS, 1);
  set(8, DRUMMERS, 1);
  set(9, BASSISTS, 1);
  set(10, ACOUSTICS, 1);
  set(11, SAXOPHONISTS, 1);
  set(12, VIOLINISTS, 1);
}
