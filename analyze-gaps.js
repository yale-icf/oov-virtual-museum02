const XLSX = require('xlsx');
const wb = XLSX.readFile('financial_documents_template.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

console.log('Total documents:', data.length);
console.log('');

const fields = ['title','description','type','subjectCountry','Period','currency','language','issueDate'];
fields.forEach(f => {
  const missing = data.filter(r => {
    const val = r[f];
    return val === undefined || val === null || String(val).trim() === '';
  });
  console.log(f + ' missing: ' + missing.length);
});

console.log('');

const keyFields = ['title','description','type','subjectCountry','Period'];
const partial = data.filter(r => {
  const hasOne = keyFields.some(f => {
    const val = r[f];
    return val !== undefined && val !== null && String(val).trim() !== '';
  });
  const missOne = keyFields.some(f => {
    const val = r[f];
    return val === undefined || val === null || String(val).trim() === '';
  });
  return hasOne && missOne;
});
console.log('Documents with PARTIAL key metadata:', partial.length);

keyFields.forEach(f => {
  const count = partial.filter(r => {
    const val = r[f];
    return val === undefined || val === null || String(val).trim() === '';
  }).length;
  if (count > 0) console.log('  - missing ' + f + ': ' + count);
});

const fullyEmpty = data.filter(r => {
  return keyFields.every(f => {
    const val = r[f];
    return val === undefined || val === null || String(val).trim() === '';
  });
});
console.log('\nFully empty (all key fields blank):', fullyEmpty.length);
if (fullyEmpty.length > 0) {
  console.log('  Files:', fullyEmpty.map(r => r.filename).join(', '));
}

// Show some examples of partial docs
console.log('\n--- Sample partial documents ---');
partial.slice(0, 10).forEach(r => {
  const missing = keyFields.filter(f => {
    const val = r[f];
    return val === undefined || val === null || String(val).trim() === '';
  });
  console.log(r.filename + ': missing [' + missing.join(', ') + ']');
});
if (partial.length > 10) {
  console.log('... and ' + (partial.length - 10) + ' more');
}
