const XLSX = require('xlsx');
const path = require('path');

const SPREADSHEET = path.join(__dirname, 'financial_documents_template.xlsx');
const wb = XLSX.readFile(SPREADSHEET);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

const isEmpty = (val) => val === undefined || val === null || String(val).trim() === '';

// All filenames we touched in our 3 update scripts
const round1 = [ // update-metadata.js: 109 fully empty docs
  'goetzmann0630.jpg','goetzmann0631.jpg','goetzmann0632.jpg','goetzmann0633.jpg','goetzmann0634.jpg',
  'goetzmann0635.jpg','goetzmann0636.jpg','goetzmann0637.jpg','goetzmann0638.jpg','goetzmann0639.jpg',
  'goetzmann0640.jpg','goetzmann0641.jpg','goetzmann0642.jpg','goetzmann0643.jpg','goetzmann0644.jpg',
  'goetzmann0645.jpg','goetzmann0646.jpg','goetzmann0647.jpg','goetzmann0648.jpg','goetzmann0650.jpg',
  'goetzmann0651.jpg','goetzmann0652.jpg','goetzmann0653.jpg','goetzmann0654.jpg','goetzmann0655.jpg',
  'goetzmann0656.jpg','goetzmann0657.jpg','goetzmann0658.jpg','goetzmann0659.jpg','goetzmann0660.jpg',
  'goetzmann0661.jpg','goetzmann0662.jpg','goetzmann0663.jpg','goetzmann0664.jpg','goetzmann0665.jpg',
  'goetzmann0666.jpg','goetzmann0667.jpg','goetzmann0668.jpg','goetzmann0669.jpg','goetzmann0670.jpg',
  'goetzmann0671.jpg','goetzmann0672.jpg','goetzmann0673.jpg','goetzmann0674.jpg','goetzmann0675.jpg',
  'goetzmann0676.jpg','goetzmann0677.jpg','goetzmann0678.jpg','goetzmann0679.jpg','goetzmann0688.jpg',
  'goetzmann0689.jpg','goetzmann0690.jpg','goetzmann0691.jpg','goetzmann0692.jpg','goetzmann0693.jpg',
  'goetzmann0694.jpg','goetzmann0695.jpg','goetzmann0696.jpg','goetzmann0697.jpg','goetzmann0698.jpg',
  'goetzmann0699.jpg','goetzmann0701.jpg','goetzmann0702.jpg','goetzmann0718.jpg','goetzmann0733.jpg',
  'goetzmann0738.jpg','goetzmann0966.jpg','goetzmann0967.jpg','goetzmann0968.jpg','goetzmann0969.jpg',
  'goetzmann0970.jpg','goetzmann0971.jpg','goetzmann0974.jpg','goetzmann0975.jpg','goetzmann0980.jpg',
  'goetzmann0981.jpg','goetzmann0982.jpg','goetzmann0983.jpg','goetzmann0984.jpg','goetzmann0985.jpg',
  'goetzmann0986.jpg','goetzmann0987.jpg','goetzmann0988.jpg','goetzmann0989.jpg','goetzmann0990.jpg',
  'goetzmann0991.jpg','goetzmann0992.jpg','goetzmann0993.jpg','goetzmann0996.jpg','goetzmann0997.jpg',
  'goetzmann0998.jpg','goetzmann0999.jpg','goetzmann1000.jpg','goetzmann1001.jpg','goetzmann1002.jpg',
  'goetzmann1003.jpg','goetzmann1004.jpg','goetzmann1005.jpg','goetzmann1006.jpg','goetzmann1007.jpg',
  'goetzmann1008.jpg','goetzmann1009.jpg','goetzmann1010.jpg','goetzmann1011.jpg','goetzmann1022.jpg',
  'goetzmann1023.jpg','goetzmann1024.jpg','goetzmann1025.jpg','goetzmann1026.jpg'
];

const round2_image = [ // update-partial.js: image-analyzed partial docs
  'goetzmann0543.jpg','goetzmann0545.jpg','goetzmann0550.jpg','goetzmann0608.jpg','goetzmann0609.jpg',
  'goetzmann0710.jpg','goetzmann0711.jpg','goetzmann0712.jpg','goetzmann0713.jpg','goetzmann0714.jpg',
  'goetzmann0715.jpg','goetzmann0716.jpg','goetzmann0717.jpg','goetzmann0719.jpg'
];

const round2_period = [ // update-partial.js: period fixes
  'goetzmann0002.jpg','goetzmann0367.jpg','goetzmann0381.jpg','goetzmann0386.jpg','goetzmann0387.jpg',
  'goetzmann0392.jpg','goetzmann0401.jpg','goetzmann0403.jpg','goetzmann0412.jpg','goetzmann0413.jpg',
  'goetzmann0420.jpg','goetzmann0421.jpg','goetzmann0422.jpg','goetzmann0423.jpg','goetzmann0427.jpg',
  'goetzmann0433.jpg','goetzmann0435.jpg','goetzmann0499.jpg','goetzmann0506.jpg','goetzmann0508.jpg',
  'goetzmann0513.jpg','goetzmann0526.jpg','goetzmann0538.jpg','goetzmann0554.jpg','goetzmann0601.jpg'
];

const round2_desc = [ // update-partial.js: description fixes
  'goetzmann0410.jpg','goetzmann0486.jpg','goetzmann0495.jpg','goetzmann0528.jpg','goetzmann0536.jpg',
  'goetzmann0537.jpg','goetzmann0542.jpg','goetzmann0606.jpg','goetzmann1015.jpg','goetzmann1016.jpg',
  'goetzmann1017.jpg','goetzmann1018.jpg','goetzmann1019.jpg','goetzmann1020.jpg','goetzmann1021.jpg',
  'goetzmann1040.jpg','goetzmann1041.jpg'
];

const round2_type = [ // update-partial.js: type fixes
  'goetzmann0367.jpg','goetzmann0381.jpg','goetzmann0386.jpg','goetzmann0387.jpg','goetzmann0393.jpg',
  'goetzmann0401.jpg','goetzmann0412.jpg','goetzmann0413.jpg','goetzmann0416.jpg','goetzmann0420.jpg',
  'goetzmann0421.jpg','goetzmann0422.jpg','goetzmann0423.jpg','goetzmann0427.jpg','goetzmann0435.jpg',
  'goetzmann0499.jpg','goetzmann0506.jpg','goetzmann0508.jpg','goetzmann0510.jpg','goetzmann0512.jpg',
  'goetzmann0513.jpg','goetzmann0526.jpg','goetzmann0540.jpg','goetzmann0683.jpg','goetzmann1027.jpg'
];

// Combine all unique filenames
const allChanged = new Set([...round1, ...round2_image, ...round2_period, ...round2_desc, ...round2_type]);

// Also add the fix-countries.js docs (same files as round1 + round2_image plus some extras)
// Those 106 issuingCountry fixes overlap with the above sets

console.log('=== Documents Changed Today ===\n');
console.log('TOTAL UNIQUE DOCUMENTS MODIFIED:', allChanged.size);

console.log('\n--- Round 1: Fully empty docs filled from image analysis (108 updated + 1 skipped) ---');
console.log('Count:', round1.length);
console.log('Files:', round1.join(', '));

console.log('\n--- Round 2a: Partial docs filled from image analysis (14) ---');
console.log('Count:', round2_image.length);
console.log('Files:', round2_image.join(', '));

console.log('\n--- Round 2b: Period inferred from titles (25) ---');
console.log('Count:', round2_period.length);
console.log('Files:', round2_period.join(', '));

console.log('\n--- Round 2c: Descriptions added (17) ---');
console.log('Count:', round2_desc.length);
console.log('Files:', round2_desc.join(', '));

console.log('\n--- Round 2d: Type classifications added (25) ---');
console.log('Count:', round2_type.length);
console.log('Files:', round2_type.join(', '));

console.log('\n--- Round 3: issuingCountry added (106 docs, overlaps with above) ---');

console.log('\n=== What was changed per document ===');
const sortedFiles = [...allChanged].sort();
sortedFiles.forEach(fn => {
  const changes = [];
  if (round1.includes(fn)) changes.push('full metadata from image');
  if (round2_image.includes(fn)) changes.push('title/desc/type/period from image');
  if (round2_period.includes(fn)) changes.push('Period');
  if (round2_desc.includes(fn)) changes.push('description');
  if (round2_type.includes(fn)) changes.push('type');
  const r = data.find(d => d.filename === fn);
  const title = r ? (r.title || '[no title]').substring(0, 55) : '[not found]';
  console.log(fn + ' | ' + changes.join(', ') + ' | ' + title);
});
