const pr = new Intl.PluralRules('en-US', { type: 'ordinal' });

function getSundaysInMonth(year: number, monthIndex: number): number[] {
  const sundaysInTerm = [];
  const date = new Date(year, monthIndex);

  // Get first Sunday.
  let dayIndex = 0;
  while (date.getDay() > 0) date.setDate(++dayIndex);

  // Get and add all Sundays.
  while (date.getMonth() === monthIndex) {
    sundaysInTerm.push(date.getDate());
    date.setDate(date.getDate() + 7);
  }

  return sundaysInTerm;
}

function formatOrdinal(n: number) {
  const rule = pr.select(n);
  const suffix = SUFFIXES.get(rule);
  return `${n}${suffix}`;
}
