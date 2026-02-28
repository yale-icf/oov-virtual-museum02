const XLSX = require('xlsx');
const wb = XLSX.readFile('financial_documents_template.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const isEmpty = (val) => val === undefined || val === null || String(val).trim() === '';

// Find docs missing title, description, type, or Period but NOT fully empty
const targets = data.filter(r => {
  const missingTitle = isEmpty(r.title);
  const missingDesc = isEmpty(r.description);
  const missingType = isEmpty(r.type);
  const missingPeriod = isEmpty(r.Period);
  const hasSomething = !isEmpty(r.title) || !isEmpty(r.description) || !isEmpty(r.type) || !isEmpty(r.subjectCountry);
  return (missingTitle || missingDesc || missingType || missingPeriod);
});

console.log('Total docs missing at least one of title/description/type/Period:', targets.length);
console.log('');

// Group by what they're missing
const missingTitle = targets.filter(r => isEmpty(r.title));
const missingDesc = targets.filter(r => isEmpty(r.description));
const missingType = targets.filter(r => isEmpty(r.type));
const missingPeriod = targets.filter(r => isEmpty(r.Period));

console.log('Missing title (' + missingTitle.length + '):');
missingTitle.forEach(r => {
  console.log('  ' + (r.filename || '[no filename]') + ' | type: ' + (r.type || '') + ' | country: ' + (r.subjectCountry || ''));
});

console.log('\nMissing type (' + missingType.length + '):');
missingType.forEach(r => {
  console.log('  ' + (r.filename || '[no filename]') + ' | title: ' + (r.title || '').substring(0, 60));
});

console.log('\nMissing Period (' + missingPeriod.length + '):');
missingPeriod.forEach(r => {
  const title = (r.title || '').substring(0, 50);
  const date = r.issueDate || '';
  console.log('  ' + (r.filename || '[no filename]') + ' | title: ' + title + ' | date: ' + date);
});

console.log('\nMissing description only (have title, type, period) (' +
  targets.filter(r => isEmpty(r.description) && !isEmpty(r.title) && !isEmpty(r.type) && !isEmpty(r.Period)).length + '):');
targets.filter(r => isEmpty(r.description) && !isEmpty(r.title) && !isEmpty(r.type) && !isEmpty(r.Period)).forEach(r => {
  console.log('  ' + (r.filename || '[no filename]') + ' | title: ' + (r.title || '').substring(0, 60));
});
