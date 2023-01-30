const TERM: Term = 2;
const YEAR = 2023;

const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

const termMonths = displayMonthsForTerm(TERM);
const SHEET_NAME = `T${TERM} ${YEAR} (${termMonths})`;

function main() {
  const sheet = createSheet();
  populateSheet(sheet);
  resizeRows(sheet);
}

function populateSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  console.info(`Setting validation for ${sheet.getSheetName()}`);

  const [sundaysInTerm, numberOfSundays] = getSundaysInTerm();
  const range = sheet.getRange(3, 1, numberOfSundays, NUMBER_OF_COLUMNS);

  let rowIndex = 1;
  Object.entries(sundaysInTerm).forEach(([month, dates]) => {
    const rowNumber = rowIndex;
    range.getCell(rowNumber, 1).setValue(month);
    dates.forEach((date) => {
      rowIndex++;
      const rule = new Intl.PluralRules('en-US', { type: 'ordinal' });
      const ordinal = rule.select(date);
      range.getCell(rowNumber, 2).setValue(date + ordinal);
    });
  });
}

function getSundaysInTerm(): [Record<string, number[]>, number] {
  const monthIndices = UF_TERMS[TERM];
  const sundaysInTerm = monthIndices.reduce<Record<string, number[]>>(
    (monthMap, monthIndex) => {
      const date = new Date(YEAR, monthIndex);
      const month = date.toLocaleDateString('default', { month: 'long' });
      const addSunday = (dayNumber: number) => {
        monthMap[month] = [...(monthMap[month] || []), dayNumber];
      };

      // Get first Sunday.
      while (date.getDay() > 0) date.setDate(1);
      addSunday(date.getDate());

      // Get subsequent Sundays.
      while (date.getMonth() === monthIndex) {
        date.setDate(date.getDate() + 7);
        addSunday(date.getDate());
      }

      return monthMap;
    },
    {},
  );
  const dateEntryLength = Object.values(sundaysInTerm).reduce(
    (acc, { length }) => acc + length,
    0,
  );

  return [sundaysInTerm, dateEntryLength];
}

function displayMonthsForTerm(term: Term) {
  const monthIndices = UF_TERMS[term];
  return monthIndices
    .map((monthIndex) => {
      const date = new Date(YEAR, monthIndex);
      return date.toLocaleString('default', { month: 'short' });
    })
    .join(', ');
}

function resizeRows(sheet: GoogleAppsScript.Spreadsheet.Sheet) {
  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 40);
  sheet.setRowHeights(3, 20, 31);

  sheet.setColumnWidth(2, 57);
  sheet.setColumnWidth(3, 100);
}

/**
 * Creates the sheet, deleting it if a sheet witht he name already exists.
 * @returns The sheet.
 */
function createSheet() {
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }

  sheet = spreadsheet.insertSheet();
  sheet.setName(SHEET_NAME);
  return sheet;
}
