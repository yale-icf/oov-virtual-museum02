// Fix non-conforming 'type' field values across all previously written rows
const xlsx = require('xlsx');
const path = require('path');

const filePath = path.join(__dirname, '..', 'oov_data_new.xlsx');
const wb = xlsx.readFile(filePath);
const ws = wb.Sheets['Documents'];
const data = xlsx.utils.sheet_to_json(ws, { header: 1 });
const headers = data[0];
const col = {};
headers.forEach((h, i) => { col[h] = i; });

function set(rowIdx, field, value) {
  if (col[field] === undefined) return;
  while (data[rowIdx].length <= col[field]) data[rowIdx].push('');
  data[rowIdx][col[field]] = value;
}

// Rows 101-121: 'Illustration' → 'Certificate' (Dutch Windkaarten playing cards)
for (let r = 101; r <= 121; r++) set(r, 'type', 'Certificate');

// Rows 236-291: 'Legislative Document' → 'Pamphlet' (Acts of Parliament, 8 Geo I and 24 Geo II)
for (let r = 236; r <= 291; r++) set(r, 'type', 'Pamphlet');

// Share Warrant → Certificate
[297, 311, 330, 334, 366].forEach(r => set(r, 'type', 'Certificate'));

// Dividend Coupon → Coupon
[299, 355].forEach(r => set(r, 'type', 'Coupon'));

// Depositary Receipt → Certificate, Receipt
[313, 314].forEach(r => set(r, 'type', 'Certificate, Receipt'));

// Profit Share → Certificate
set(331, 'type', 'Certificate');

// Founders Share → Stock Certificate
set(358, 'type', 'Stock Certificate');

// Interest Share → Certificate
set(370, 'type', 'Certificate');

// Stock Warrant → Certificate
set(372, 'type', 'Certificate');

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Fixed type fields across rows 101-121, 236-291, 297, 299, 311, 313-314, 330-331, 334, 355, 358, 366, 370, 372.');
