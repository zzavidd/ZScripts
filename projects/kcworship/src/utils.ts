const pr = new Intl.PluralRules('en-US', { type: 'ordinal' });

function getSundaysInTerm() {
  const monthIndices = UF_TERMS[TERM];
  const sundaysInTerm = monthIndices.reduce<Record<string, number[]>>(
    (monthMap, monthIndex) => {
      const date = new Date(YEAR, monthIndex);
      const month = date.toLocaleDateString('default', { month: 'long' });
      const addSunday = (dayNumber: number) => {
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

function displayMonthsForTerm(term: Term) {
  const monthIndices = UF_TERMS[term];
  return monthIndices
    .map((monthIndex) => {
      const date = new Date(YEAR, monthIndex);
      return date.toLocaleString('default', { month: 'short' });
    })
    .join(', ');
}

function formatOrdinal(n: number) {
  const rule = pr.select(n);
  const suffix = SUFFIXES.get(rule);
  return `${n}${suffix}`;
}
