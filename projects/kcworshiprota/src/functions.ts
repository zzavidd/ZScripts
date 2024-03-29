const pr = new Intl.PluralRules('en-US', { type: 'ordinal' });

function getSundaysInTerm(term: Term, year: number) {
  const monthIndices = UF_TERMS.get(term);
  if (!monthIndices) throw new Error('No month indices.');

  const sundaysInTerm = monthIndices.reduce<Record<string, number[]>>(
    (monthMap, monthIndex) => {
      const date = new Date(year, monthIndex);
      const month = date.toLocaleDateString('default', { month: 'long' });
      const addSunday = (dayNumber: number): void => {
        monthMap[month] = [...(monthMap[month] || []), dayNumber];
      };

      // Get first Sunday.
      let dayIndex = 0;
      while (date.getDay() > 0) date.setDate(++dayIndex);

      // Get and add all Sundays.
      while (date.getMonth() === monthIndex) {
        addSunday(date.getDate());
        date.setDate(date.getDate() + 7);
      }

      return monthMap;
    },
    {},
  );
  const numberOfSundays = Object.values(sundaysInTerm).reduce(
    (acc, { length }) => acc + length,
    0,
  );

  return { sundaysInTerm, numberOfSundays };
}

function displayMonthsForTerm(term: Term, year: number) {
  const monthIndices = UF_TERMS.get(term);
  if (!monthIndices) throw new Error('No month indices.');

  return monthIndices
    .map((monthIndex) => {
      const date = new Date(year, monthIndex);
      return date.toLocaleString('default', { month: 'short' });
    })
    .join(', ');
}

function formatOrdinal(n: number) {
  const rule = pr.select(n);
  const suffix = SUFFIXES.get(rule);
  return `${n}${suffix}`;
}
