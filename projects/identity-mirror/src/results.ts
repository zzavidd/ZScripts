const RESULTS_OPTIONS: Record<string, number> = {
  'Yes': 1,
  'Not sure': 0,
  'No': -1,
} as const;

function displayResults() {
  PARTICIPANTS.forEach((participant) => {
    const resultsSheetName = `Results: ${participant}`;

    try {
      spreadsheet.deleteSheet(spreadsheet.getSheetByName(resultsSheetName)!);
    } catch (e: any) {
      console.warn(e.message);
    }

    const resultsSheet = spreadsheet.insertSheet();
    resultsSheet.setName(resultsSheetName);
    resultsSheet.setFrozenRows(1);
    resultsSheet.setTabColor('#0000ff');
    const resultsRange = resultsSheet.getRange(1, 1, TOTAL_ROWS, TOTAL_COLUMNS);

    formatRange(resultsSheet);
    resultsSheet
      .getRange(2, 1, TOTAL_ROWS - 1, 1)
      .setBackground(HEADER_COLUMN_COLOR);
    resultsSheet
      .getRange(1, 6, TOTAL_ROWS, 1)
      .setBackground(HEADER_ROW_COLOR)
      .setFontColor('#fff')
      .setFontWeight('bold');

    resultsSheet.setColumnWidth(6, 70);

    resultsRange.getCell(1, 2).setValue(`You (${participant})`);

    const otherParticipants = [...PARTICIPANTS];
    otherParticipants.splice(otherParticipants.indexOf(participant), 1);
    otherParticipants.forEach((p, i) => {
      resultsRange.getCell(1, i + 3).setValue(p);
    });

    ADJECTIVES.sort((a, b) => a.word.localeCompare(b.word)).forEach(
      (adjective, i) => {
        resultsSheet.getRange(i + 2, 1).setValue(adjective.word);
      },
    );

    resultsSheet
      .getRange(2, 2, TOTAL_ROWS - 1, 1)
      .setFormulaR1C1(`='Poll: ${participant}'!R[0]C[0]`);

    otherParticipants.forEach((p, i) => {
      const column = spreadsheet
        .getSheetByName(`Poll: ${p}`)!
        .getRange(1, 1, 1, TOTAL_COLUMNS)
        .getValues()[0]
        .findIndex((value) => value === participant);

      resultsSheet
        .getRange(2, i + 3, TOTAL_ROWS - 1, 1)
        .setFormulaR1C1(`='Poll: ${p}'!R[0]C${column + 1}`);
    });

    resultsSheet
      .getRange(2, 6, TOTAL_ROWS - 1, 1)
      .setFormulaR1C1('=calculateRowValue(R[0]C2:R[0]C5)');

    resultsSheet.sort(6);
  });
}

/**
 * Calculates the total value from the options.
 * @param input The input range.
 * @returns The total value.
 * @customfunction
 */
function calculateRowValue([row]: string[][]) {
  return row.reduce((acc, value) => acc + RESULTS_OPTIONS[value], 0);
}
