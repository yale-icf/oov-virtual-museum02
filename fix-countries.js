const XLSX = require('xlsx');
const path = require('path');

const SPREADSHEET = path.join(__dirname, 'financial_documents_template.xlsx');

const isEmpty = (val) => val === undefined || val === null || String(val).trim() === '';

// For documents we updated: set issuingCountry based on what we know from image analysis
// Key: if the bond was issued internationally (e.g., Russian bonds traded in Amsterdam),
// issuingCountry = where it was traded/printed, subjectCountry = the underlying country
// If domestic, issuingCountry = subjectCountry

const issuingCountryOverrides = {
  // Russian bonds issued via Amsterdam/European markets
  'goetzmann0630.jpg': 'Netherlands', // already set
  'goetzmann0631.jpg': 'Netherlands', // already set
  // Russian bonds issued domestically by Imperial/Soviet government
  'goetzmann0667.jpg': 'Russia',
  'goetzmann0668.jpg': 'Russia',
  'goetzmann0669.jpg': 'Russia',
  'goetzmann0670.jpg': 'Russia',
  'goetzmann0671.jpg': 'Russia',
  'goetzmann0672.jpg': 'Russia',
  'goetzmann0673.jpg': 'Russia',
  'goetzmann0674.jpg': 'Russia',
  'goetzmann0675.jpg': 'Russia',
  'goetzmann0676.jpg': 'Russia',
  'goetzmann0677.jpg': 'Russia',
  'goetzmann0678.jpg': 'Russia',
  'goetzmann0679.jpg': 'United States', // Sallie Mae, domestic US
  'goetzmann0688.jpg': 'United States', // Sallie Mae
  'goetzmann0689.jpg': 'United States',
  'goetzmann0738.jpg': 'Russia',
  'goetzmann0980.jpg': 'Russia', // Grand Russian Railway
  'goetzmann0981.jpg': 'Russia',
  'goetzmann0982.jpg': 'Russia',
  'goetzmann0983.jpg': 'Russia',
  'goetzmann0998.jpg': 'Russia', // USSR bonds
  'goetzmann0999.jpg': 'Russia',
  'goetzmann1000.jpg': 'Russia',
  'goetzmann1001.jpg': 'Russia',
  'goetzmann1004.jpg': 'Russia', // Imperial Russian Consolidated Railway
  'goetzmann1005.jpg': 'Russia',
  'goetzmann1010.jpg': 'Netherlands', // Receipt issued in Amsterdam
  'goetzmann1011.jpg': 'Netherlands',
  // Chinese bonds issued through European banks
  'goetzmann0697.jpg': 'United Kingdom', // Chinese Imperial 4.5% Gold Loan, £100, issued London
  'goetzmann0698.jpg': 'United Kingdom',
  'goetzmann0699.jpg': 'United Kingdom',
  'goetzmann0970.jpg': 'Germany', // Chinese Imperial Gold Loan via Deutsch-Asiatische Bank
  'goetzmann0971.jpg': 'Germany',
  'goetzmann1002.jpg': 'China', // Republic of China domestic
  'goetzmann1003.jpg': 'China', // Nationalist Government, Canton
  'goetzmann1006.jpg': 'China', // Shanghai Pudong Taxi, domestic
  'goetzmann1007.jpg': 'China',
  'goetzmann1025.jpg': 'China', // Ming Dynasty, domestic
  'goetzmann1026.jpg': 'China',
  // Banque Industrielle de Chine - headquartered in Paris
  'goetzmann0988.jpg': 'France', // already set
  'goetzmann0989.jpg': 'France',
  // Bulgarian bonds
  'goetzmann0648.jpg': 'Bulgaria',
  'goetzmann0665.jpg': 'Bulgaria',
  'goetzmann0666.jpg': 'Bulgaria',
  'goetzmann0966.jpg': 'Bulgaria',
  'goetzmann0967.jpg': 'Bulgaria',
  'goetzmann0968.jpg': 'Bulgaria',
  'goetzmann0969.jpg': 'Bulgaria',
  // Ottoman/Turkish bonds
  'goetzmann0710.jpg': 'France', // French société anonyme
  'goetzmann0711.jpg': 'Netherlands', // Issued in Amsterdam
  'goetzmann0712.jpg': 'Netherlands',
  // Other domestic bonds
  'goetzmann0633.jpg': 'Austria',
  'goetzmann0634.jpg': 'Azerbaijan',
  'goetzmann0635.jpg': 'Bolivia',
  'goetzmann0636.jpg': 'Bolivia',
  'goetzmann0637.jpg': 'Bolivia',
  'goetzmann0638.jpg': 'Yugoslavia',
  'goetzmann0639.jpg': 'Austria',
  'goetzmann0640.jpg': 'Serbia',
  'goetzmann0641.jpg': 'Bosnia and Herzegovina',
  'goetzmann0642.jpg': 'Brazil',
  'goetzmann0643.jpg': 'Hungary',
  'goetzmann0644.jpg': 'Hungary',
  'goetzmann0645.jpg': 'Argentina',
  'goetzmann0646.jpg': 'Hungary',
  'goetzmann0651.jpg': 'Colombia',
  'goetzmann0652.jpg': 'Cuba',
  'goetzmann0653.jpg': 'Czechoslovakia',
  'goetzmann0654.jpg': 'Egypt',
  'goetzmann0655.jpg': 'United Kingdom',
  'goetzmann0656.jpg': 'United Kingdom',
  'goetzmann0657.jpg': 'Estonia',
  'goetzmann0658.jpg': 'Serbia',
  'goetzmann0659.jpg': 'Germany',
  'goetzmann0660.jpg': 'Greece',
  'goetzmann0661.jpg': 'United Kingdom',
  'goetzmann0662.jpg': 'Hungary',
  'goetzmann0690.jpg': 'Peru',
  'goetzmann0691.jpg': 'Chile',
  'goetzmann0692.jpg': 'Chile',
  'goetzmann0696.jpg': 'Serbia',
  'goetzmann0701.jpg': 'Honduras',
  'goetzmann0702.jpg': 'Honduras',
  'goetzmann0713.jpg': 'Peru',
  'goetzmann0714.jpg': 'Poland',
  'goetzmann0715.jpg': 'Poland',
  'goetzmann0716.jpg': 'Poland',
  'goetzmann0717.jpg': 'United Kingdom', // Portuguese debt conversion, issued in London
  'goetzmann0718.jpg': 'Portugal',
  'goetzmann0719.jpg': 'Netherlands', // Russian funds certificate issued in Amsterdam
  'goetzmann0974.jpg': 'France',
  'goetzmann0975.jpg': 'France',
  'goetzmann0984.jpg': 'Canada',
  'goetzmann0985.jpg': 'Canada',
  'goetzmann0986.jpg': 'Canada',
  'goetzmann0987.jpg': 'Canada',
  'goetzmann0990.jpg': 'Egypt',
  'goetzmann0991.jpg': 'Egypt',
  'goetzmann0992.jpg': 'Egypt',
  'goetzmann0993.jpg': 'Egypt',
  'goetzmann0996.jpg': 'Austria',
  'goetzmann0997.jpg': 'Austria',
  'goetzmann1008.jpg': 'Belgium',
  'goetzmann1009.jpg': 'Belgium',
  'goetzmann1022.jpg': 'Netherlands',
  'goetzmann1023.jpg': 'United Kingdom',
  'goetzmann1024.jpg': 'United Kingdom',
  'goetzmann1027.jpg': 'Italy', // Monte di Pieta
  // Notarial deeds issued in Netherlands
  'goetzmann0543.jpg': 'Netherlands', // already set
  'goetzmann0545.jpg': 'Netherlands',
  'goetzmann0550.jpg': 'Netherlands',
  'goetzmann0608.jpg': 'Netherlands',
  'goetzmann0609.jpg': 'Netherlands',
  // Warnsveld/England War Loan
  'goetzmann0632.jpg': 'Netherlands', // already set
  // India stock via Bank of England
  'goetzmann0663.jpg': 'United Kingdom', // already set
  'goetzmann0664.jpg': 'United Kingdom', // already set
  // Mozambique/Portugal
  'goetzmann0647.jpg': 'Portugal', // already set
  // Chrysler via Netherlands
  'goetzmann0650.jpg': 'Netherlands', // already set
  // Uncertain ones
  'goetzmann0693.jpg': 'France',
  'goetzmann0694.jpg': 'France',
  'goetzmann0695.jpg': 'France',
};

// Read spreadsheet
const wb = XLSX.readFile(SPREADSHEET);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

let fixedIssuing = 0;
let fixedSubject = 0;

for (const row of data) {
  const fn = row.filename;

  // Set issuingCountry if we have an override and it's currently empty
  if (issuingCountryOverrides[fn] && isEmpty(row.issuingCountry)) {
    row.issuingCountry = issuingCountryOverrides[fn];
    fixedIssuing++;
  }

  // For any doc that has issuingCountry but no subjectCountry,
  // and the subject is clearly the same as issuing (domestic bonds),
  // we can leave subjectCountry empty (matching existing convention).
  // But for docs where we set subjectCountry and it equals what issuingCountry would be,
  // that's fine too — just ensures both are populated.
}

// Write back
const newWs = XLSX.utils.json_to_sheet(data);
wb.Sheets[wb.SheetNames[0]] = newWs;
XLSX.writeFile(wb, SPREADSHEET);

console.log('issuingCountry fixes applied:', fixedIssuing);

// Final verification
console.log('\n--- Final country stats ---');
console.log('issuingCountry filled:', data.filter(r => r.issuingCountry && String(r.issuingCountry).trim() !== '').length);
console.log('issuingCountry empty:', data.filter(r => isEmpty(r.issuingCountry)).length);
console.log('subjectCountry filled:', data.filter(r => r.subjectCountry && String(r.subjectCountry).trim() !== '').length);
console.log('subjectCountry empty:', data.filter(r => isEmpty(r.subjectCountry)).length);

// Show our updated docs - verify both fields
console.log('\n--- Sample of our updated docs (both fields) ---');
const samples = ['goetzmann0630.jpg','goetzmann0670.jpg','goetzmann0697.jpg','goetzmann0970.jpg',
  'goetzmann0984.jpg','goetzmann0998.jpg','goetzmann1003.jpg','goetzmann0710.jpg','goetzmann0719.jpg'];
samples.forEach(fn => {
  const r = data.find(d => d.filename === fn);
  if (r) {
    console.log(fn + ': subject=' + (r.subjectCountry || '[empty]') + ' | issuing=' + (r.issuingCountry || '[empty]'));
  }
});
