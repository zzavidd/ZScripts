/** A map of terms to the corresponding month indices. */
const UF_TERMS = new Map<Term, number[]>([
  [1, [0, 1, 2]],
  [2, [3, 4, 5]],
  [3, [6, 7, 8]],
  [4, [9, 10, 11]],
]);

/** The text style for the header. */
const COLUMN_HEADER_TEXT_STYLE = SpreadsheetApp.newTextStyle()
  .setBold(true)
  .setFontFamily('Nunito')
  .setFontSize(10)
  .setForegroundColor('#FFFFFF')
  .build();

const TEXT_STYLE = SpreadsheetApp.newTextStyle()
  .setFontFamily('Nunito')
  .setFontSize(11)
  .build();

/** The total number of expected non-empty columns. */
const TOTAL_NONEMPTY_COLUMNS = 13;

/** A map of date suffixes. */
const SUFFIXES = new Map([
  ['one', 'st'],
  ['two', 'nd'],
  ['few', 'rd'],
  ['other', 'th'],
]);

const COLOURS = {
  RED: '#e6b8af',
  ORANGE: '#fce5cd',
  YELLOW: '#fff2cc',
  GREEN: '#d9ead3',
  TEAL: '#d0e0e3',
  BLUE: '#c9daf8',
  PURPLE: '#d9d2e9',
  PINK: '#ead1dc',
};

type Term = 1 | 2 | 3 | 4;
