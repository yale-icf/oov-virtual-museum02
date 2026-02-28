const XLSX = require('xlsx');
const wb = XLSX.readFile('financial_documents_template.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const isEmpty = (val) => val === undefined || val === null || String(val).trim() === '';

// First, look at existing patterns - how do docs with BOTH fields filled use them?
console.log('=== Existing docs with BOTH issuingCountry and subjectCountry ===');
const both = data.filter(r => r.issuingCountry && r.subjectCountry &&
  String(r.issuingCountry).trim() !== '' && String(r.subjectCountry).trim() !== '' &&
  String(r.issuingCountry).trim() !== String(r.subjectCountry).trim());
both.slice(0, 10).forEach(r => {
  console.log(r.filename + ': subject=' + r.subjectCountry + ' | issuing=' + r.issuingCountry);
});
console.log('Total with both (different):', both.length);

console.log('\n=== Documents we updated - current country state ===');
// Check all docs from our update batches
const ourFiles = data.filter(r => {
  const fn = r.filename || '';
  const num = parseInt(fn.replace('goetzmann','').replace('.jpg',''));
  return (num >= 630 && num <= 702) || num === 718 || num === 738 ||
    (num >= 966 && num <= 975) || (num >= 980 && num <= 993) ||
    (num >= 996 && num <= 1011) || (num >= 1022 && num <= 1026) ||
    num === 543 || num === 545 || num === 550 || num === 608 || num === 609 ||
    (num >= 710 && num <= 719);
});

let missingIssuing = 0;
let missingSubject = 0;

ourFiles.forEach(r => {
  const hasSubject = r.subjectCountry && String(r.subjectCountry).trim() !== '';
  const hasIssuing = r.issuingCountry && String(r.issuingCountry).trim() !== '';
  if (hasSubject && !hasIssuing) missingIssuing++;
  if (!hasSubject && hasIssuing) missingSubject++;
  if (hasSubject && !hasIssuing) {
    // Show ones where issuing might differ from subject
    const title = (r.title || '').toLowerCase();
    const subj = String(r.subjectCountry);
    // Flag if title hints at a different issuing country
    if (title.includes('amsterdam') || title.includes('dutch') || title.includes('netherlands') ||
        title.includes('london') || title.includes('paris') || title.includes('french') ||
        subj === 'Russia' || subj === 'China') {
      console.log(r.filename + ': subject=' + subj + ' | issuing=[EMPTY] | title: ' + (r.title || '').substring(0, 70));
    }
  }
});

console.log('\nOur updated docs: ' + ourFiles.length + ' total');
console.log('Have subjectCountry but no issuingCountry:', missingIssuing);
console.log('Have issuingCountry but no subjectCountry:', missingSubject);

// Show overall stats
console.log('\n=== Overall issuingCountry stats ===');
console.log('Total rows:', data.length);
console.log('issuingCountry filled:', data.filter(r => r.issuingCountry && String(r.issuingCountry).trim() !== '').length);
console.log('issuingCountry empty:', data.filter(r => isEmpty(r.issuingCountry)).length);
console.log('subjectCountry filled:', data.filter(r => r.subjectCountry && String(r.subjectCountry).trim() !== '').length);
console.log('subjectCountry empty:', data.filter(r => isEmpty(r.subjectCountry)).length);
