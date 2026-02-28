// Fix rows 274-285: update numberPages from 12 to 18 and correct "Page X of 12" → "Page X of 18"
const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'oov_data_new.xlsx');
const wb = xlsx.readFile(filePath);
const ws = wb.Sheets['Documents'];
const data = xlsx.utils.sheet_to_json(ws, { header: 1 });
const headers = data[0];
const col = {};
headers.forEach((h, i) => { col[h] = i; });

function get(rowIdx, field) {
  if (col[field] === undefined) return '';
  return data[rowIdx][col[field]] || '';
}
function set(rowIdx, field, value) {
  if (col[field] === undefined) return;
  while (data[rowIdx].length <= col[field]) data[rowIdx].push('');
  data[rowIdx][col[field]] = value;
}

for (let r = 274; r <= 285; r++) {
  set(r, 'numberPages', 18);
  const title = get(r, 'title');
  const desc = get(r, 'description');
  set(r, 'title', title.replace('of 12)', 'of 18)'));
  set(r, 'description', desc.replace('of 12)', 'of 18)'));
}

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Fixed rows 274-285: numberPages updated to 18, page references corrected.');
