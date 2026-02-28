// Fixes rows 191-202: updates numberPages from 12 to 23 and corrects "Page X of 12" → "Page X of 23"
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

// Fix rows 191-202: numberPages 12 → 23, and fix title strings "Page X of 12" → "Page X of 23"
for (let r = 191; r <= 202; r++) {
  set(r, 'numberPages', 23);
  const title = get(r, 'title');
  const desc = get(r, 'description');
  set(r, 'title', title.replace('of 12)', 'of 23)'));
  set(r, 'description', desc.replace('of 12)', 'of 23)'));
}

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Fixed rows 191-202: numberPages updated to 23, page references corrected.');
