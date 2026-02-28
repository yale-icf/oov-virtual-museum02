const XLSX = require('xlsx');
const path = require('path');

const SPREADSHEET = path.join(__dirname, 'financial_documents_template.xlsx');
const SEP = '\u001d';

const isEmpty = (val) => val === undefined || val === null || String(val).trim() === '';

// ====== PART 1: Image-analyzed documents (missing title/type/period) ======
const imageAnalyzed = {
  'goetzmann0543.jpg': {
    title: 'Contract for Sale of British Consolidated Stock (Consols), Amsterdam, 1805',
    description: 'Dutch-language contract for the sale of British Consolidated stock (Consols) in Pounds Sterling. Issued through Ricardo and De Lara at the Waarpoleis, Amsterdam. Dated May 13, 1805.',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'GBP',
    language: 'Dutch',
    issueDate: '1805'
  },
  'goetzmann0545.jpg': {
    title: 'Conditions of Negotiation for Essequebo and Demerary Plantation Fund, f 400,000',
    description: 'Prospectus/conditions for a plantation loan fund under the direction of Daniel Changuion, providing f 400,000 over 10 years at 6% interest to planters in Rio Essequebo and Rio Demerary (present-day Guyana). Dutch language.',
    subjectCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'Guilders',
    language: 'Dutch'
  },
  'goetzmann0550.jpg': {
    title: 'Municipality of Warnsveld Loan Certificate, 250 Guilders, 1883',
    description: 'Proof of share (Bewijs van Aandeel) in the municipal loan of f 12,000 by the Gemeente Warnsveld, Gelderland. Denomination of 250 Guilders at 4½% interest. Established by council resolution of September 16, 1882. Dated April 30, 1883. Stamped "UITGELOOT" (redeemed).',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'Guilders',
    language: 'Dutch',
    issueDate: '1883'
  },
  'goetzmann0608.jpg': {
    title: 'Notarial Copy of Mortgage Deed for Plantation Vauxhall, Dominica, 1777',
    description: 'Notarial copy (Copia) of a mortgage deed relating to Plantation Vauxhall on the island of Dominica. Involves Cornelis van Herpion, Notary in Middelburg, and Fredrik Cornelis Stolkert. References the Compagnie van Demerara and colonial properties. Dated 1777.',
    type: ['Document', 'Legal'].join(SEP),
    subjectCountry: 'Netherlands',
    Period: '18th Century or before',
    currency: 'GBP',
    language: 'Dutch',
    issueDate: '1777'
  },
  'goetzmann0609.jpg': {
    title: 'Notarial Deed of Mortgage for Plantation Vauxhall, Dominica, March 21, 1777',
    description: 'Notarial deed (Copie) of mortgage for Plantation Vauxhall on the island of Dominica. Executed before Hieronimus de Wolf Junior, Notaris Publiq in Amsterdam, on March 21, 1777. Involves Frans Jacob Heshuysen and the Societeyt of Adolff Jan Heshuysen en Compagnie. Mortgage of £7,100 Sterling with James Balmer, merchant in London.',
    type: ['Document', 'Legal'].join(SEP),
    subjectCountry: 'Netherlands',
    Period: '18th Century or before',
    currency: 'GBP',
    language: 'Dutch',
    issueDate: '1777'
  },
  'goetzmann0710.jpg': {
    title: 'Ottoman Damascus-Hamah Railway (Homs-Tripoli Extension) 4% Bond, 500 Francs, 1909',
    description: 'Bearer bond of the Société Ottomane du Chemin de Fer de Damas-Hamah et Prolongements (Ottoman Damascus-Hamah Railway and Extensions). 4% bond, denomination of 500 Francs or 22 Turkish Liras. Emprunt 1909 for the Homs-Tripoli line. Capital of 15,000,000 Francs. French and Arabic/Ottoman Turkish text.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Turkey',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French, Arabic',
    issueDate: '1909'
  },
  'goetzmann0711.jpg': {
    title: 'Ottoman Public Debt Receipt, Series A 1928, Liras 2.20 / £2 / 50 Francs',
    description: 'Bearer receipt (Reçu au Porteur) from the Conseil de la Dette Publique Répartie de l\'Ancien Empire Ottoman (Council of the Distributed Public Debt of the Former Ottoman Empire). Exchangeable for bonds representing arrears of the Ottoman Public Debt, Series A 1928. Denomination of Liras 2.20 / £2 / 50 Francs. Dated Amsterdam, April 22, 1930. Issued through Banque de Paris et des Pays-Bas.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Turkey',
    Period: '20th Century',
    currency: 'Turkish Liras',
    language: 'French',
    issueDate: '1930'
  },
  'goetzmann0712.jpg': {
    title: 'Ottoman Public Debt Receipt, Series A 1928, Liras 4.40 / £4 / 100 Francs',
    description: 'Bearer receipt from the Council of the Distributed Public Debt of the Former Ottoman Empire. Series A 1928. Denomination of Liras 4.40 / £4 / 100 Francs. Dated Amsterdam, April 22, 1930. Issued through Banque de Paris et des Pays-Bas.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Turkey',
    Period: '20th Century',
    currency: 'Turkish Liras',
    language: 'French',
    issueDate: '1930'
  },
  'goetzmann0713.jpg': {
    title: 'Republic of Peru Bond for Arica Customs House Construction, 100 Soles, 1871',
    description: 'Bond of the Republic of Peru for the construction of the Aduana de Arica (Customs House of Arica). Denomination of 100 Soles. Dated Lima, July 16, 1871. Features Peruvian coat of arms and decorative green border.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Peru',
    Period: '19th Century',
    currency: 'Soles',
    language: 'Spanish',
    issueDate: '1871'
  },
  'goetzmann0714.jpg': {
    title: 'City of Warsaw 4½% Municipal Bond, 1931',
    description: 'Municipal bond (Obligation/Obligacja) of the Capital City of Warsaw at 4½% interest, issued in 1931. Bilingual French and Polish text. Multiple currency denominations including Zloty, USD, GBP, and Francs. Green ornate design with Warsaw city views and coat of arms.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Poland',
    Period: '20th Century',
    currency: 'Zloty',
    language: 'French, Polish',
    issueDate: '1931'
  },
  'goetzmann0715.jpg': {
    title: 'Republic of Poland Series III Premium Dollar Loan Bond, $5, 1931',
    description: 'Bond (Obligacja) of the Republic of Poland (Rzeczpospolita Polska), Series III Premium Dollar Loan. Value of $5 US Dollars or 44.57 Zloty. Blue design. Dated Warsaw, February 1, 1931.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Poland',
    Period: '20th Century',
    currency: 'USD',
    language: 'Polish',
    issueDate: '1931'
  },
  'goetzmann0716.jpg': {
    title: 'Republic of Poland Series III Premium Dollar Loan Bond, $5, with Coupons, 1931',
    description: 'Bond of the Republic of Poland, Series III Premium Dollar Loan. Value of $5 or 44.57 Zloty. Same design as goetzmann0715 but with four attached coupons (numbered 17-20). Dated February 1, 1931.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Poland',
    Period: '20th Century',
    currency: 'USD',
    language: 'Polish',
    issueDate: '1931'
  },
  'goetzmann0717.jpg': {
    title: 'Conversion of Portuguese External Debt, Provisional Certificate of 3% Deferred Stock, 1852',
    description: 'Provisional certificate of 3% deferred stock issued under the Decree of December 18, 1852, for the conversion of the external debt of Portugal. Stock bears interest from January 1, 1863. Issued through the Portuguese Financial Agency in London. Green/teal decorative border.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Portugal',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1852'
  },
  'goetzmann0719.jpg': {
    title: 'Certificate of 6% Russian Funds in Bank Assignations, 1,000 Roubles, 1825',
    description: 'Bilingual Dutch/French certificate for 6% Russian government funds in bank assignations. Denomination of 1,000 Roubles. Issued through Hope en Comp., Ketwich et Voombergh, and Wed. W. Borski in Amsterdam. Dated May 11, 1825. Ornate decorative border.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Dutch, French',
    issueDate: '1825'
  }
};

// ====== PART 2: Documents missing only Period (can infer from title/date) ======
const periodFixes = {
  'goetzmann0002.jpg': '19th Century', // Burmese document, likely 19th century
  'goetzmann0367.jpg': '20th Century', // Phu-Quoc exploitation, French Indochina
  'goetzmann0381.jpg': '20th Century', // Same
  'goetzmann0386.jpg': '19th Century', // Kaiserlich Russische Regierung (Imperial Russian)
  'goetzmann0387.jpg': '19th Century', // Same
  'goetzmann0392.jpg': '19th Century', // La Platense Flotilla Co.
  'goetzmann0401.jpg': '19th Century', // Morris Canal
  'goetzmann0403.jpg': '19th Century', // New-York Central Rail Road Company
  'goetzmann0412.jpg': '18th Century or before', // Ord Der Paters Norbertynen
  'goetzmann0413.jpg': '19th Century', // Oregon and Transcontinental Company
  'goetzmann0420.jpg': '20th Century', // Polsko Amerykanskie (Polish-American)
  'goetzmann0421.jpg': '20th Century', // Ports Debarcadere Maritime
  'goetzmann0422.jpg': '19th Century', // Portuguese External Debt
  'goetzmann0423.jpg': '19th Century', // Pramien Obligation (Premium bond, likely Austrian)
  'goetzmann0427.jpg': '20th Century', // Reichsbanknote
  'goetzmann0433.jpg': '19th Century', // 4% Russian Government bond
  'goetzmann0435.jpg': '20th Century', // Schuldverschreibung
  'goetzmann0499.jpg': '19th Century', // Mexican Deferred Stock
  'goetzmann0506.jpg': '19th Century', // Poyaisian Land Grant
  'goetzmann0508.jpg': '20th Century', // Chinese Republic
  'goetzmann0513.jpg': '19th Century', // Rjanm Uralsk (Russian railway)
  'goetzmann0526.jpg': '20th Century', // Temporary Regulations
  'goetzmann0538.jpg': '18th Century or before', // Bond on behalf of Charles VI (1723)
  'goetzmann0554.jpg': '19th Century', // New Russia Company limited
  'goetzmann0601.jpg': '19th Century', // Compagnie de Colonisation
};

// ====== PART 3: Documents missing only description (have title/type/period) ======
const descriptionFixes = {
  'goetzmann0410.jpg': 'Share or bond certificate of Omnium Français du Film, a French film industry company.',
  'goetzmann0486.jpg': 'Financial document issued by Hope & Company and Ketwich & Voombergh, prominent Amsterdam banking houses involved in international finance.',
  'goetzmann0495.jpg': 'London certificate, likely a financial instrument or stock certificate issued in London.',
  'goetzmann0528.jpg': 'Certificate of Deferred Debt (Uitgestelde Schuld), a Dutch government debt instrument representing deferred obligations.',
  'goetzmann0536.jpg': 'Perpetual annuity bond issued by Wilhelm, Prince of Orange. A historical Dutch financial instrument providing perpetual annual payments.',
  'goetzmann0537.jpg': 'Master bond (Hoofd-Obligatie) of the Swedish Crown Loan, representing the principal obligation of a loan to the Swedish Crown.',
  'goetzmann0542.jpg': 'Bond certificate issued on behalf of the Princes of France, a royal French debt instrument.',
  'goetzmann0606.jpg': 'Private tontine certificate. A tontine is an investment scheme where subscribers share diminishing payments as members die, with the last survivor receiving the entire fund.',
  'goetzmann1015.jpg': 'Document related to the Tyler Building, likely a real estate bond or mortgage certificate.',
  'goetzmann1016.jpg': 'Document related to the Tyler Building, continuation or additional page.',
  'goetzmann1017.jpg': 'Document related to the Tyler Building, continuation or additional page.',
  'goetzmann1018.jpg': 'Document related to the Maplewood Suburban Home Company, likely a share certificate or bond for a suburban real estate development company.',
  'goetzmann1019.jpg': 'Document related to the Maplewood Suburban Home Company, continuation or additional page.',
  'goetzmann1020.jpg': 'Document related to the Maplewood Suburban Home Company, continuation or additional page.',
  'goetzmann1021.jpg': 'Bond of the Dutch East India Company (VOC), issued October 26, 16xx. One of the earliest corporate bonds, issued by the world\'s first multinational corporation.',
  'goetzmann1040.jpg': 'Plantation loan certificate for Rio Essequebo en Rio Demmerary (present-day Guyana). Dutch colonial-era financial instrument for funding plantation operations.',
  'goetzmann1041.jpg': 'Plantation loan certificate for Rio Essequebo en Rio Demmerary, continuation or additional page.',
};

// ====== PART 4: Missing type (can infer from title) ======
const typeFixes = {
  'goetzmann0367.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0381.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0386.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0387.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0393.jpg': 'Document',
  'goetzmann0401.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0412.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0413.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0416.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0420.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0421.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0422.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0423.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0427.jpg': ['Currency', 'Banknote'].join(SEP),
  'goetzmann0435.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0499.jpg': ['Equity', 'Security'].join(SEP),
  'goetzmann0506.jpg': ['Document', 'Land Grant'].join(SEP),
  'goetzmann0508.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0510.jpg': ['Bond', 'Debt', 'Annuity'].join(SEP),
  'goetzmann0512.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0513.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann0526.jpg': 'Document',
  'goetzmann0540.jpg': ['Document', 'Land Grant'].join(SEP),
  'goetzmann0683.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
  'goetzmann1027.jpg': ['Bond', 'Debt', 'Security'].join(SEP),
};

// ====== Apply updates ======
const wb = XLSX.readFile(SPREADSHEET);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

let counts = { image: 0, period: 0, desc: 0, type: 0 };

for (const row of data) {
  const fn = row.filename;

  // Part 1: Image-analyzed full updates
  if (imageAnalyzed[fn]) {
    const m = imageAnalyzed[fn];
    if (isEmpty(row.title) && m.title) row.title = m.title;
    if (isEmpty(row.description) && m.description) row.description = m.description;
    if (m.type && (isEmpty(row.type) || row.type === 'security' || row.type === 'Derivative')) row.type = m.type;
    if (isEmpty(row.subjectCountry) && m.subjectCountry) row.subjectCountry = m.subjectCountry;
    if (m.issuingCountry && isEmpty(row.issuingCountry)) row.issuingCountry = m.issuingCountry;
    if (isEmpty(row.Period) && m.Period) row.Period = m.Period;
    if (isEmpty(row.currency) && m.currency) row.currency = m.currency;
    if (isEmpty(row.language) && m.language) row.language = m.language;
    if (isEmpty(row.issueDate) && m.issueDate) row.issueDate = m.issueDate;
    counts.image++;
  }

  // Part 2: Period fixes
  if (periodFixes[fn] && isEmpty(row.Period)) {
    row.Period = periodFixes[fn];
    counts.period++;
  }

  // Part 3: Description fixes
  if (descriptionFixes[fn] && isEmpty(row.description)) {
    row.description = descriptionFixes[fn];
    counts.desc++;
  }

  // Part 4: Type fixes
  if (typeFixes[fn] && isEmpty(row.type)) {
    row.type = typeFixes[fn];
    counts.type++;
  }
}

const newWs = XLSX.utils.json_to_sheet(data);
wb.Sheets[wb.SheetNames[0]] = newWs;
XLSX.writeFile(wb, SPREADSHEET);

console.log('Image-analyzed updates:', counts.image);
console.log('Period fixes:', counts.period);
console.log('Description fixes:', counts.desc);
console.log('Type fixes:', counts.type);
console.log('Total updates:', counts.image + counts.period + counts.desc + counts.type);

// Final stats
const fields = ['title','description','type','Period'];
console.log('\n--- Remaining gaps ---');
fields.forEach(f => {
  const missing = data.filter(r => isEmpty(r[f]));
  console.log(f + ' still missing: ' + missing.length);
});
