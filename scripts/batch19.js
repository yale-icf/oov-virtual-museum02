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
  data[rowIdx][col[field]] = value;
}

function setDoc(rowIdx, title, description, meta) {
  set(rowIdx, 'title', title);
  set(rowIdx, 'description', description);
  if (meta.type !== undefined) set(rowIdx, 'type', meta.type);
  if (meta.subjectCountry !== undefined) set(rowIdx, 'subjectCountry', meta.subjectCountry);
  if (meta.issuingCountry !== undefined) set(rowIdx, 'issuingCountry', meta.issuingCountry);
  if (meta.creator !== undefined) set(rowIdx, 'creator', meta.creator);
  if (meta.issueDate !== undefined) set(rowIdx, 'issueDate', meta.issueDate);
  if (meta.currency !== undefined) set(rowIdx, 'currency', meta.currency);
  if (meta.language !== undefined) set(rowIdx, 'language', meta.language);
  if (meta.numberPages !== undefined) set(rowIdx, 'numberPages', meta.numberPages);
  if (meta.period !== undefined) set(rowIdx, 'period', meta.period);
  if (meta.notes !== undefined) set(rowIdx, 'notes', meta.notes);
}

// Row 545: Changuion Negotiatie Conditions, Article 1 text, 1816
setDoc(545,
  'Conditien van Negotiatie, Daniel Changuion, Essequibo & Demerara, Article 1, f.400,000 at 6%, 1816',
  'Printed conditions text (Article 1) of the negotiatie organized under the directorship of Daniel Changuion for planters in Rio Essequibo and Rio Demerara, to fund improvements to their plantations for a total of f.400,000 at 6% interest over 10 years. This page shows the body article text of the conditions, complementing the title page (see related document). Dated 1816.',
  {
    type: 'negotiatie',
    subjectCountry: 'Netherlands|Guyana',
    issuingCountry: 'Netherlands',
    creator: 'Daniel Changuion',
    issueDate: '1816',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Plantation negotiatie for Essequibo and Demerara (now Guyana) planters; f.400,000 at 6% for 10 years; Article 1 of conditions text; see related title page'
  }
);

// Row 546: Russian Perpetual Income Obligation, Bezstochny
setDoc(546,
  'Russian Imperial Perpetual Income Obligation (Bezstochny), Heavily Stamped, 19th Century',
  'Russian imperial government perpetual income obligation bearing the word "БЕЗСРОЧНЫЙ" (Bezstochny, meaning perpetual/without term). The document is heavily annotated with numerous oval revenue stamps in purple, red, and blue, and handwritten endorsements recording successive interest payments or coupon collections during the 1860s–1870s. Printed on pale paper with Cyrillic text.',
  {
    type: 'bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Russian Imperial Government',
    issueDate: 'ca. 1855',
    currency: 'Russian ruble',
    language: 'Russian',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: '"БЕЗСРОЧНЫЙ" = perpetual; multiple stamp cancellations indicating payment history across 1860s–1870s; perpetual government bond'
  }
);

// Row 547: Pekin Syndicate Limited, Bearer Certificate No. 835003, 5 Ordinary Shares, 1900
setDoc(547,
  'Pekin Syndicate Limited, Bearer Certificate No. 835003, Five Ordinary Shares, 1900',
  'Bilingual English/French bearer share certificate for the Pekin Syndicate Limited (London), No. 835003, representing five ordinary shares of £1 each in the company. The Pekin Syndicate was a British company formed to exploit mining and railway concessions in northern China (Shanxi province). Printed in green on cream with a fine decorative border. Coupon sheet (coupons 1–33) is attached at the bottom. Capital structure: £1,200,000 ordinary stock and £800,000 in 6% First Mortgage Debentures.',
  {
    type: 'share',
    subjectCountry: 'China',
    issuingCountry: 'United Kingdom',
    creator: 'Pekin Syndicate Limited',
    issueDate: 'November 1900',
    currency: 'British pound sterling',
    language: 'English|French',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'British concession company in northern China (Shanxi); bilingual English/French bearer certificate; No. 835003; coupon sheet coupons 1–33 attached'
  }
);

// Row 548: Pekin Syndicate Limited, Bearer Certificate No. 835003, Reverse Side
setDoc(548,
  'Pekin Syndicate Limited, Bearer Certificate No. 835003, Reverse Side with Coupon Sheet',
  'Reverse/back side of Pekin Syndicate Limited bearer certificate No. 835003, showing the conditions of transfer and a coupon sheet with coupons numbered in a grid format (coupons 1–33). This side contains standard bearer conditions and the reverse of the company logo. The coupon layout on the reverse is arranged in reverse column order from the front face.',
  {
    type: 'share',
    subjectCountry: 'China',
    issuingCountry: 'United Kingdom',
    creator: 'Pekin Syndicate Limited',
    issueDate: 'November 1900',
    currency: 'British pound sterling',
    language: 'English|French',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Reverse of bearer certificate No. 835003; coupon sheet verso showing coupons 1–33'
  }
);

// Row 549: Vlaardingen Weeshuis Negotiatie Prospectus, f.75,000, Batavian Republic, 1800
setDoc(549,
  'Negotiatie Prospectus/Conditions, Vlaardingen Weeshuis (Orphanage), f.75,000, Batavian Republic, 1800',
  'Printed lottery-style negotiatie prospectus and conditions document for a fund of f.75,000 established under the Batavian Republic\'s representative assembly on May 9, 1800, for the maintenance of the Weeshuis (Orphanage) of the city of Vlaardingen. Details the prize structure: 514 shares of f.150 each, with 14 premium prizes ranging from f.50,000 down to f.125, and 500 premia of f.60 each. Also outlines five age-class life annuity rates (under 15, 15–24, 25–36, 36–48, over 48 years). Signed at the Vlaardingsche Weeshuys.',
  {
    type: 'tontine',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Vertegenwoordigend Lichaam des Bataafsen Volks',
    issueDate: 'May 9, 1800',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Batavian Republic period',
    notes: 'Lottery-style charitable negotiatie for Vlaardingen orphanage; f.75,000 total; five age-class life annuity structure; prize structure listed; see also related obligation certificate'
  }
);

// Row 550: Municipality of Warnsveld Municipal Loan Bond, f.250 at 4.5%, 1883
setDoc(550,
  'Gemeente Warnsveld, Bewijs van Aandeel in Municipal Loan f.12,000, f.250 at 4.5%, 1883',
  'Dutch municipal bond share certificate (Bewijs van Aandeel) in a f.12,000 municipal loan of the Gemeente (Municipality) of Warnsveld (Gelderland province), bearing interest of 4.5% per annum. Face value f.250. The loan was authorized by resolution of the Municipal Council on September 16 and October 21, 1882, approved by the States of Gelderland on October 25, 1882, for a total loan amount of f.14,000. Redeemable by lottery over twelve years. Signed by the Burgemeester and Wethouder of Warnsveld. Dated April 30, 1883.',
  {
    type: 'bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Gemeente Warnsveld',
    issueDate: 'April 30, 1883',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Municipal loan; f.12,000 total; f.250 denomination; 4.5% interest; lottery redemption; Warnsveld, Gelderland'
  }
);

// Row 551: 3% Buitenlandsche Schuld van Portugal, Certificaat No. 1402, Amsterdam, 1892
setDoc(551,
  '3% Buitenlandsche Schuld van Portugal, Certificaat No. 1402, Amsterdam, 1892',
  'Certificate No. 1402 issued by the Vereeniging voor den Effectenhandel (Amsterdam Securities Exchange Association) representing an arrears claim against the Portuguese State for unpaid interest of two-thirds of the 3% Portuguese Foreign Debt (1853–1884), due July 1, 1892. Issued in connection with the Nederlandsche Brekeringe-Comité voor Portugeesche Fondsen te Amsterdam (Dutch Arbitration Committee for Portuguese Funds in Amsterdam). Dated Amsterdam, September 30, 1892.',
  {
    type: 'certificate',
    subjectCountry: 'Portugal',
    issuingCountry: 'Netherlands',
    creator: 'Vereeniging voor den Effectenhandel, Amsterdam',
    issueDate: 'September 30, 1892',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Portuguese 3% foreign debt 1853–1884; arrears interest certificate; issued by Amsterdam Securities Exchange; Dutch-Portuguese debt arbitration'
  }
);

// Row 552: New Russia Company Limited, Debenture No. 277, Coupon Sheet (large format)
setDoc(552,
  'New Russia Company Limited, 6% First Mortgage Debenture No. 277, £100, Coupon Sheet',
  'Large format coupon sheet for New Russia Company Limited 6% First Mortgage Debenture No. 277, denomination £100. Each coupon bears the Company name and debenture number. The New Russia Company Limited was a British company developing agricultural land in the Ekaterinoslav Government of southern Russia (now Ukraine). Coupons are printed in blue and arranged in a multi-column grid covering the full life of the debenture.',
  {
    type: 'bond',
    subjectCountry: 'Russia',
    issuingCountry: 'United Kingdom',
    creator: 'New Russia Company Limited',
    issueDate: 'ca. 1901',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Coupon sheet for debenture No. 277, £100; New Russia Company Ltd.; agricultural land in Ekaterinoslav/southern Russia; 6% first mortgage'
  }
);

// Row 553: New Russia Company Limited, 6% First Mortgage Debenture No. 277, £100, 1901
setDoc(553,
  'New Russia Company Limited, 6% First Mortgage Debenture No. 277, £100, London, 1901',
  '6% First Mortgage Debenture No. 277 of New Russia Company Limited, London, £100 denomination, from a total issue of £800,000. Secured on property of the Company consisting of land in the Government of Ekaterinoslav, southern Russia (now Ukraine). Repayable on October 1, 1933, with semi-annual 6% interest payments on April 1 and October 1. Issued under the Company\'s Deed of Trust signed by directors including Alfred Lyttelton KC and Charles Eden. Printed in blue/black on cream with decorative border.',
  {
    type: 'bond',
    subjectCountry: 'Russia',
    issuingCountry: 'United Kingdom',
    creator: 'New Russia Company Limited',
    issueDate: '1901',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'British company; £800,000 total issue; No. 277; 6% first mortgage debenture; land in Ekaterinoslav, southern Russia; repayable 1933'
  }
);

// Row 554: New Russia Company Limited, Debenture No. 277, Reverse with Amortization Table
setDoc(554,
  'New Russia Company Limited, 6% First Mortgage Debenture No. 277, Reverse Side with Amortization Table',
  'Reverse/back side of New Russia Company Limited 6% First Mortgage Debenture No. 277, £100, showing a detailed amortization table, blank transfer register, and condition clauses. The amortization table outlines the annual redemption schedule for the £800,000 debenture issue over its life to 1933. The right portion contains a partially printed transfer register in tabular form.',
  {
    type: 'bond',
    subjectCountry: 'Russia',
    issuingCountry: 'United Kingdom',
    creator: 'New Russia Company Limited',
    issueDate: '1901',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of debenture No. 277; amortization table and transfer register'
  }
);

// Row 555: New Russia Company Limited, Debenture No. 277, Coupon Sheet (Coupons 22-70)
setDoc(555,
  'New Russia Company Limited, 6% First Mortgage Debenture No. 277, Coupon Sheet (Coupons 22–70)',
  'Blue-printed coupon sheet for New Russia Company Limited 6% First Mortgage Debenture No. 277, £100, containing interest coupons numbered 22 through 70, covering the later payment periods of the debenture\'s life from approximately the 1910s to 1935. Coupons are arranged in three columns. Each coupon references the debenture number and bears the company name. This is the later-series coupon sheet accompanying the debenture.',
  {
    type: 'bond',
    subjectCountry: 'Russia',
    issuingCountry: 'United Kingdom',
    creator: 'New Russia Company Limited',
    issueDate: '1901',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Coupon sheet coupons 22–70; later-series coupons for debenture No. 277; blue printed on white'
  }
);

// Row 556: Société Selim et Samaan Sednaoui, 25 Registered Shares, LE.10, Cairo, UAR
setDoc(556,
  'Société Selim et Samaan Sednaoui S.A., 25 Registered Shares of LE.10 Each, Cairo, United Arab Republic',
  'Share certificate in Arabic and French for the Société Selim et Samaan Sednaoui, Société Anonyme R.A.U. (République Arabe Unie/United Arab Republic), headquartered in Cairo. Certificate represents 25 registered (nominative) shares of Egyptian Pound (LE) 10 each, total LE.250. The Sednaoui family founded one of Cairo\'s premier department stores in 1886; the company was later nationalized. The certificate is printed in red/terracotta with Arabic title text and has a detachable dividend coupon sheet at the bottom. Dated ca. 1958–1960.',
  {
    type: 'share',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Société Selim et Samaan Sednaoui S.A.',
    issueDate: 'ca. 1958',
    currency: 'Egyptian pound',
    language: 'Arabic|French',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'Famous Cairo department store founded 1886; United Arab Republic era (1958–1961); 25 registered shares at LE.10 each; later nationalized'
  }
);

// Row 557: Sednaoui Share Certificate, Transfer Register (Reverse)
setDoc(557,
  'Société Selim et Samaan Sednaoui, Share Certificate, Arabic Transfer Register (Reverse)',
  'Reverse/back side of the Société Selim et Samaan Sednaoui registered share certificate, showing the Arabic transfer of ownership register table headed "انتقال ملكية الأسهم" (Intiqal Milkiyat al-Asham / Transfer of Share Ownership). The table has columns for date, name, and countersignature in Arabic script. One transfer entry is partially visible recording a transfer to Sednaoui Sednaoui.',
  {
    type: 'share',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Société Selim et Samaan Sednaoui S.A.',
    issueDate: 'ca. 1958',
    currency: 'Egyptian pound',
    language: 'Arabic',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'Reverse of registered share certificate; Arabic transfer register "انتقال ملكية الأسهم"'
  }
);

// Row 558: Mexico 3% Consolidated National Debt Bond, 5,000 Pesos, Mid-19th Century
setDoc(558,
  'Mexico, Deuda Nacional Consolidada al 3%, Bono de 5,000 Pesos, Mid-19th Century',
  'Mexican 3% Consolidated National Debt bond (Deuda Nacional Consolidada al Tres por Ciento) for 5,000 Pesos. Printed with an ornate decorative border and featuring the Mexican eagle national seal. The bond is from Mexico\'s mid-19th century era of debt consolidation and renegotiation. Includes endorsement or coupon tables. Likely issued under one of Mexico\'s consolidated debt conversion operations of the 1840s–1860s.',
  {
    type: 'bond',
    subjectCountry: 'Mexico',
    issuingCountry: 'Mexico',
    creator: 'Gobierno de México',
    issueDate: 'ca. 1845-1860',
    currency: 'Mexican peso',
    language: 'Spanish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Mexican consolidated debt; 3% Deuda Nacional Consolidada; 5,000 pesos denomination; 19th-century Mexican debt restructuring'
  }
);

// Row 559: Mexican Republic Consolidated Treasury Debt, Letra F, $5,000
setDoc(559,
  'Mexican Republic, Deuda Consolidada del Tesoro Público, Letra F, $5,000, Mid-19th Century',
  'Large format Mexican Republic consolidated public debt bond (Deuda Consolidada del Tesoro Público de la República Mexicana), series Letra F, face value $5,000. Issued by the Secretaría de Estado y del Despacho de Hacienda (Treasury Ministry). The bond has a printed coupon schedule on the right side, signed by treasury officials. The Mexican eagle with republican inscription appears prominently. From Mexico\'s serial domestic debt consolidation operations of the 1840s–1860s.',
  {
    type: 'bond',
    subjectCountry: 'Mexico',
    issuingCountry: 'Mexico',
    creator: 'Secretaría de Hacienda, República Mexicana',
    issueDate: 'ca. 1845-1860',
    currency: 'Mexican peso',
    language: 'Spanish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Series Letra F; $5,000 denomination; Deuda Consolidada del Tesoro; coupon schedule at right; Mexican republican eagle seal'
  }
);

// Row 560: Compagnie Impériale Chemins de Fer Ethiopiens, Action 500F, No. 01,927
setDoc(560,
  'Compagnie Impériale Chemins de Fer Ethiopiens, Action de 500 Francs au Porteur, No. T 01,927',
  '500 Franc bearer share (action au porteur) of the Compagnie Impériale Chemins de Fer Ethiopiens (Imperial Ethiopian Railway Company), certificate No. T 01,927. Printed in terracotta/rust-red on cream paper. Features an illustration of an African landscape scene (trees and figures) on the left panel. Bilingual Amharic and French text throughout. The company was formed under Emperor Menelik II to build the railway from Djibouti to Addis Ababa. Coupon sheet attached at the bottom.',
  {
    type: 'share',
    subjectCountry: 'Ethiopia',
    issuingCountry: 'France',
    creator: 'Compagnie Impériale Chemins de Fer Ethiopiens',
    issueDate: 'ca. 1899-1905',
    currency: 'French franc',
    language: 'French|Amharic',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Djibouti–Addis Ababa railway; Emperor Menelik II era; 500 franc bearer share; bilingual French/Amharic; No. T 01,927'
  }
);

// Row 561: Compagnie Impériale Chemins de Fer Ethiopiens, Action No. 01,927, Reverse Side
setDoc(561,
  'Compagnie Impériale Chemins de Fer Ethiopiens, Action No. T 01,927, Reverse Side',
  'Reverse/back side of the Compagnie Impériale Chemins de Fer Ethiopiens 500 Franc bearer share No. T 01,927. Shows faint mirror-image printing from the obverse (Amharic text and railway illustration in reverse) and purple dividend distribution stamps: "1ère REPARTITION du 1 au 18..." and dated "du 30 MAI 1908," indicating the first dividend distribution was recorded in May 1908.',
  {
    type: 'share',
    subjectCountry: 'Ethiopia',
    issuingCountry: 'France',
    creator: 'Compagnie Impériale Chemins de Fer Ethiopiens',
    issueDate: 'ca. 1899-1905',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Reverse of share No. T 01,927; first dividend distribution stamp dated 30 May 1908'
  }
);

// Row 562: Compagnie Impériale Chemins de Fer Ethiopiens, Coupon Sheet, Exercices 1900-1947
setDoc(562,
  'Compagnie Impériale Chemins de Fer Ethiopiens, Dividend Coupon Sheet, Exercices 1900–1947',
  'Full dividend coupon sheet for Compagnie Impériale Chemins de Fer Ethiopiens shares, containing 48 annual dividend coupons labeled "EXERCICE" (fiscal year) running from 1900 through 1947. Each coupon shows the exercise year. The sheet is printed in black on white paper. Coupons are arranged in a four-column grid. These were detached annually as dividends were paid out.',
  {
    type: 'share',
    subjectCountry: 'Ethiopia',
    issuingCountry: 'France',
    creator: 'Compagnie Impériale Chemins de Fer Ethiopiens',
    issueDate: 'ca. 1899-1905',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Annual dividend coupon sheet; exercices 1900–1947; 48 coupons; likely separated from a share certificate'
  }
);

// Row 563: Compagnie Impériale Chemins de Fer Ethiopiens, Action No. 01,927, Coupon Sheet
setDoc(563,
  'Compagnie Impériale Chemins de Fer Ethiopiens, Action No. 01,927, Dividend Coupon Sheet (Coupons 1–48)',
  'Dividend coupon sheet belonging to Compagnie Impériale Chemins de Fer Ethiopiens bearer share No. 01,927, containing 48 annual dividend coupons numbered 1 through 48. Each coupon is labeled "Chemins de Fer Ethiopiens, Action N° 01,927" and printed in terracotta/rust-red. Coupons are arranged in a grid running from lower-left to upper-right. This sheet accompanied the bearer share No. 01,927 and enabled collection of annual dividends.',
  {
    type: 'share',
    subjectCountry: 'Ethiopia',
    issuingCountry: 'France',
    creator: 'Compagnie Impériale Chemins de Fer Ethiopiens',
    issueDate: 'ca. 1899-1905',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Coupon sheet for Action No. 01,927; coupons 1–48; terracotta printing'
  }
);

// Row 564: Chinese Republic 8% Railway Equipment Loan 1922, £20 / 1,260 Belgian Francs
setDoc(564,
  'Chinese Republic, 8% Railway Equipment Loan 1922, £20 Treasury Note / 1,260 Belgian Francs',
  'Treasury Note (Bon du Trésor) for £20 / 1,260 Belgian Francs from the Government of the Chinese Republic\'s 8% Railway Equipment Loan of 1922, First Series of £8,000,000. Bilingual English and French title. Features Chinese red seal stamps and Chinese characters alongside the French text. Signed by authorized republic officials. The loan was intended to finance railway rolling stock and equipment for Chinese railways and was issued in multiple currencies across international markets.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: '1922',
    currency: 'British pound sterling|Belgian franc',
    language: 'French|Chinese',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '8% Railway Equipment Loan 1922; First Series £8,000,000; Belgian franc tranche: £20 / 1,260 Belgian francs; Chinese Republic government bond'
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 545–564 (20 documents, batch19).');
