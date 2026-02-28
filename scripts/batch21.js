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

// Row 585: Kingdom of Bulgaria 5% Gold Loan, Large Landscape Coupon Sheet
setDoc(585,
  'Kingdom of Bulgaria 5% Gold Loan, Large Format Landscape Coupon Sheet (Royaume de Bulgarie)',
  'Large wide-format landscape coupon sheet for a Kingdom of Bulgaria (Royaume de Bulgarie) 5% Gold loan. Contains numerous small semi-annual interest coupons arranged in a dense grid. Coupons reference the "Royaume de Bulgarie 5% Or" series and show bond numbers in the 025,xxx range in orange. This sheet is one of the companion coupon sheets for the Kingdom of Bulgaria international gold loan series (post-1908).',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Kingdom of Bulgaria, Ministry of Finance',
    issueDate: 'ca. 1909-1914',
    currency: 'French franc',
    language: 'French|Bulgarian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Large landscape coupon sheet; "Royaume de Bulgarie" 5% Gold loan; bond numbers in 025,xxx range; companion to related coupon sheets'
  }
);

// Row 586: Kingdom of Bulgaria 5% Gold Loan, Coupon Sheet No. 025,472
setDoc(586,
  'Kingdom of Bulgaria 5% Gold Loan, Coupon Sheet for Bond No. 025,472 (Royaume de Bulgarie)',
  'Wide-format interest coupon sheet for Kingdom of Bulgaria (Royaume de Bulgarie) 5% Gold loan bearer bond No. 025,472. Coupons are labeled "EMPRUNT BULGARE / ROYAUME DE BULGARIE / 5% OR" with the bond number 025,472 shown in orange in each coupon. The sheet contains semi-annual interest coupons spanning the full life of the loan. Part of the same international gold loan series as the companion bond and coupon sheets (see related documents).',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Kingdom of Bulgaria, Ministry of Finance',
    issueDate: 'ca. 1909-1914',
    currency: 'French franc',
    language: 'French|Bulgarian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Coupon sheet for bond No. 025,472; "Royaume de Bulgarie" 5% Gold loan; orange bond number; companion to No. 025,444 and other related documents'
  }
);

// Row 587: Empire du Mexique, Dette Publique Extérieure, Obligation 500 Francs, No. 1,086,528, 1865
setDoc(587,
  'Empire du Mexique, Dette Publique Extérieure, Obligation 500 Francs au Porteur, No. 1,086,528, 1865',
  'French-language 500 Franc bearer bond (Obligation au Porteur) No. 1,086,528 of the Empire of Mexico\'s External Public Debt (Dette Publique Extérieure), 1865. Issued under Emperor Maximilian I of Mexico. The bond carries semi-annual interest coupons attached on left and right sides. A smaller subsidiary billet at the bottom records the bondholder\'s claim in more condensed form: "EMPIRE DE MEXIQUE 1865 DETTE PUBLIQUE EXTÉRIEURE / LE PORTEUR." The Mexican Imperial External Debt of 1864-1865 was raised on European capital markets to finance Maximilian\'s government.',
  {
    type: 'bond',
    subjectCountry: 'Mexico',
    issuingCountry: 'Mexico',
    creator: 'Empire of Mexico (Maximilian I)',
    issueDate: '1865',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Maximilian I of Mexico; French-language bond No. 1,086,528; 500 francs; external public debt 1865; coupon sheets attached left and right; companion to No. 1,085,528'
  }
);

// Row 588: Imperio de Mexico, Deuda Publica Exterior, Obligacion 500 Francos, No. 1,085,528, 1865
setDoc(588,
  'Imperio de Mexico, Deuda Publica Exterior, Obligacion de 500 Francos al Portador, No. 1,085,528, 1865',
  'Spanish-language 500 Franc bearer bond (Obligacion al Portador) No. 1,085,528 of the Empire of Mexico\'s External Public Debt (Deuda Publica Exterior), 1865. Issued under Emperor Maximilian I of Mexico. Identical structure to the French-language version (see No. 1,086,528): main obligation at top with coupon sheets on both sides, and a smaller subsidiary billet at the bottom reading "IMPERIO DE MEXICO 1865 DEUDA PUBLICA EXTERIOR." Both French and Spanish versions of these bonds were distributed across European markets during Maximilian\'s reign.',
  {
    type: 'bond',
    subjectCountry: 'Mexico',
    issuingCountry: 'Mexico',
    creator: 'Imperio de Mexico (Maximilian I)',
    issueDate: '1865',
    currency: 'French franc',
    language: 'Spanish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Maximilian I of Mexico; Spanish-language bond No. 1,085,528; 500 francos; Deuda Publica Exterior 1865; companion Spanish version to French No. 1,086,528'
  }
);

// Row 589: Uzinele de Fier și Domeniile din Resita S.A., 10 Shares at 500 Lei, Bucharest, 1928
setDoc(589,
  'Uzinele de Fier și Domeniile din Resita S.A. / Aciéries et Domaines de Resita S.A., 10 Registered Shares of 500 Lei, Bucharest, 1928',
  'Registered share certificate (titru/titre) for 10 ordinary registered shares (Actiuni Nominative) of Uzinele de Fier și Domeniile din Resita S.A. (Romanian Iron Works and Domains of Resita / Aciéries et Domaines de Resita S.A.), nominal value 500 Lei each, total 5,000 Lei. The certificate is bilingual Romanian and French. The Resita ironworks in Banat (now Romania) were one of the oldest and largest steel-producing facilities in Eastern Europe. The certificate includes a transfer register ("CESIUNI") on the right. Dated Bucharest, September 1928. Company seal "UD" at bottom.',
  {
    type: 'share',
    subjectCountry: 'Romania',
    issuingCountry: 'Romania',
    creator: 'Uzinele de Fier și Domeniile din Resita S.A.',
    issueDate: 'September 1928',
    currency: 'Romanian leu',
    language: 'Romanian|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Resita ironworks, Banat; 10 registered shares at 500 lei each; bilingual Romanian/French; CESIUNI transfer register attached; company seal "UD"'
  }
);

// Row 590: Uzinele de Fier din Resita S.A., TALON Dividend Coupon Sheet for 10 Shares
setDoc(590,
  'Uzinele de Fier și Domeniile din Resita S.A., TALON Dividend Coupon Sheet for 10 Registered Shares, Nos. 1875141–1875170',
  'Annual dividend coupon sheet (TALON) for Uzinele de Fier și Domeniile din Resita S.A. / Aciéries et Domaines de Resita S.A., covering 10 registered shares (Actiuni Nominative) Nos. 1875141 to 1875170. Bilingual Romanian/French header. Contains annual dividend coupons for successive years (shown as 1940, 1950, 1951, 1952, 1953, 1954, and others), each labeled "Cupon de Dividende pentru 10 Actiuni / Coupon de Dividende pour 10 Actions." The TALON coupon sheet accompanied the registered share certificate for these share numbers.',
  {
    type: 'share',
    subjectCountry: 'Romania',
    issuingCountry: 'Romania',
    creator: 'Uzinele de Fier și Domeniile din Resita S.A.',
    issueDate: 'September 1928',
    currency: 'Romanian leu',
    language: 'Romanian|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'TALON dividend coupon sheet; shares Nos. 1875141–1875170; annual dividend coupons 1940–1954+; bilingual Romanian/French; companion to share certificate'
  }
);

// Row 591: Empire Ottoman, Emprunt à Primes 400 Francs, No. 1,898,329, Constantinople, 1870
setDoc(591,
  'Empire Ottoman, Emprunt à Primes de 792,000,000 Francs, Obligation 400 Francs, No. 1,898,329, Constantinople, 1870',
  'Bearer prize bond (Obligation au Porteur) No. 1,898,329 of the Ottoman Empire Prize Loan (Emprunt à Primes / Prämien-Anleihe) of 792,000,000 Francs. Total emission of 1,860,000 obligations at a nominal capital of 400 Francs each. The Ottoman Imperial Government pledges to pay principal and interest, and to honor prize draws, established in Galata/Constantinople. Multilingual text in French, German, and Ottoman Turkish (Arabic script). Signed by the Imperial Ottoman Government\'s authorized agent in Constantinople and the General Agent in Paris. Dated Constantinople, January 3, 1870. Two Ottoman revenue stamps affixed at top.',
  {
    type: 'bond',
    subjectCountry: 'Turkey',
    issuingCountry: 'Turkey',
    creator: 'Ottoman Imperial Government',
    issueDate: 'January 3, 1870',
    currency: 'French franc',
    language: 'French|German|Ottoman Turkish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Ottoman Prize Loan (Emprunt à Primes); 792,000,000 francs total; 1,860,000 obligations; No. 1,898,329; 400 francs; Constantinople 1870; prize draw lottery-style bond; trilingual French/German/Ottoman'
  }
);

// Row 592: Ottoman Empire Prize Loan, Reverse with Prize Table and Amortization Table
setDoc(592,
  'Empire Ottoman, Emprunt à Primes de 792,000,000 Francs, Reverse with Prize Table and Amortization Table',
  'Reverse/back side of the Ottoman Empire Prize Loan (Emprunt à Primes) 400 Franc bearer bond, showing: (1) "TABLEAU DES PRIMES / Tabelle der Prämien" (Prize Table) listing draw dates from 1872 to ca. 1879 with prize amounts in French and German; (2) "TABLEAU D\'AMORTISSEMENT / Plan der Tilgung" (Amortization/Sinking Fund Table) with the annual redemption schedule; and repeated descriptions of the loan in German (left column: "Prämien-Anleihe von 792,000,000 Franken") and French (right column: "Emprunt à Primes de 792,000,000 de Francs").',
  {
    type: 'bond',
    subjectCountry: 'Turkey',
    issuingCountry: 'Turkey',
    creator: 'Ottoman Imperial Government',
    issueDate: 'January 3, 1870',
    currency: 'French franc',
    language: 'French|German|Ottoman Turkish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Reverse of Ottoman Prize Loan bond; Prize Table (Tableau des Primes) 1872-1879; Amortization Table; bilingual French/German descriptions'
  }
);

// Row 593: California Mexico Land Company Limited, Certificate No. 1423, £20/50 Hectares/Frs.500, 1885
setDoc(593,
  'California Mexico Land Company Limited, Land Certificate No. 1423, 50 Hectares, £20 / Frs. 500, 1885',
  'Bilingual English/French land certificate No. 1423 of The California Mexico Land Company, Limited (London), representing 50 hectares of land in Lower California (Baja California, Mexico), valued at £20 Sterling or Francs 500. The Company issued certificates to a total nominal value of £240,000 or 6,000,000 Francs, representing the purchase of 500,000 hectares of land situated in Lower California. The certificate entitles the bearer to a proportional share of proceeds from the development and realization of the Company\'s land. Conditions dated under the Deed of Trust of May 23, 1885. Signed by company directors and secretary. Coupon sheet attached at bottom.',
  {
    type: 'certificate',
    subjectCountry: 'Mexico',
    issuingCountry: 'United Kingdom',
    creator: 'California Mexico Land Company Limited',
    issueDate: '1885',
    currency: 'British pound sterling|French franc',
    language: 'English|French',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Baja California land certificate; No. 1423; 50 hectares = £20 = Frs. 500; total issue £240,000 / Frs. 6,000,000 for 500,000 hectares; bilingual English/French; London company; coupon sheet attached'
  }
);

// Row 594: California Mexico Land Company Limited, Certificate No. 1423, Reverse/Conditions
setDoc(594,
  'California Mexico Land Company Limited, Land Certificate No. 1423, Reverse with Conditions',
  'Reverse/back side of California Mexico Land Company Limited land certificate No. 1423, 50 Hectares, £20/Frs.500. The left portion shows the full conditions text in English and French, detailing the Company\'s obligation to distribute proceeds from land development and sales to certificate holders. The right portion shows the certificate stub/coupon stub panel identifying "No. 1423, THE CALIFORNIA MEXICO LAND COMPANY, LIMITED. CERTIFICATE FOR 50 HECTARES. £20 Fcs.500."',
  {
    type: 'certificate',
    subjectCountry: 'Mexico',
    issuingCountry: 'United Kingdom',
    creator: 'California Mexico Land Company Limited',
    issueDate: '1885',
    currency: 'British pound sterling|French franc',
    language: 'English|French',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Reverse of certificate No. 1423; conditions text in English and French; stub panel at right identifying 50 hectares, £20 / Fcs. 500'
  }
);

// Row 595: Japanese Financial Document, Handwritten Manuscript, Page 1
setDoc(595,
  'Japanese Financial Document, Handwritten Manuscript with Official Seals, Meiji Era',
  'Handwritten Japanese-language financial document on paper, densely written in vertical columns of cursive Japanese script (kana and kanji). The document covers multiple columns and appears to be a formal financial or legal instrument, possibly a loan deed, land pledge, bond condition document, or government obligation certificate. Two red official seals (hanko) are affixed at the bottom. A red paper label appears in the upper right. The document style and script suggest Meiji era (1868-1912) or possibly late Edo period.',
  {
    type: 'bond',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Japanese government or merchant (unidentified)',
    issueDate: 'ca. 1870-1910',
    currency: 'Japanese yen',
    language: 'Japanese',
    numberPages: '1',
    period: 'Meiji era',
    notes: 'Handwritten Japanese document; vertical cursive script; two red official seals; red label upper right; financial or legal instrument; Meiji era or late Edo'
  }
);

// Row 596: Japanese Financial Document, Handwritten Manuscript, Page 2
setDoc(596,
  'Japanese Financial Document, Handwritten Manuscript (Second Page/Related Document), Meiji Era',
  'Second page or related handwritten Japanese-language financial document, companion to the preceding manuscript. Written in dense vertical columns of cursive Japanese script (kana and kanji). The lower portion of the document appears to contain a tabular section with shorter entries, possibly a schedule of payments, a ledger section, or an endorsement record. Script and format are consistent with the companion document (see related manuscript). Meiji era.',
  {
    type: 'bond',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Japanese government or merchant (unidentified)',
    issueDate: 'ca. 1870-1910',
    currency: 'Japanese yen',
    language: 'Japanese',
    numberPages: '1',
    period: 'Meiji era',
    notes: 'Second page/related document to companion Japanese manuscript; dense vertical cursive script; tabular section at bottom; Meiji era financial instrument'
  }
);

// Row 597: Chinese Government 8% 10-Year Sterling Treasury Notes 1925-1926, No. 5654, £100
setDoc(597,
  'Chinese Government 8% 10-Year Sterling Treasury Note 1925-1926, No. 5654, £100, with Full Coupon Sheet',
  'Treasury Note No. 5654 for £100 Sterling from the Chinese Government 8 Per Cent 10 Years Sterling Treasury Notes of 1925-1926. The note is printed on green paper with the heading "THE GOVERNMENT OF THE CHINESE REPUBLIC. CHINESE GOVERNMENT 8 Per Cent 10 Years Sterling TREASURY NOTES 1925-1926." A very large attached coupon sheet extends below, containing numbered semi-annual interest coupons (approximately 71 coupons total) for bond No. 5654, each labeled "THE CHINESE REPUBLIC." The full coupon sheet spans the life of the 10-year note.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: '1925',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '8% Sterling Treasury Notes 1925-1926; No. 5654; £100; 10-year term; full coupon sheet attached with approximately 71 semi-annual coupons; green printed'
  }
);

// Row 598: Chinese Government 5% Reorganisation Gold Loan 1913, Obligation Frs. 505, No. 489,731
setDoc(598,
  'Chinese Government 5% Reorganisation Gold Loan of 1913, Obligation de Frs. 505 / £20, No. 489,731',
  'Quintilingual bearer bond No. 489,731 of the Chinese Government 5% Reorganisation Gold Loan of 1913, for £23,000,000 Sterling (or Marks 236,250,000 or Francs 631,250,000 or Roubles 236,500,000 or Yen 193.86). Single obligation: Frs. 505 = £20 = Marks 413.40 = Roubles 193.40 = Yen 193.82. Printed in five languages: English ("Bond for £20"), German ("Schuldverschreibung über 1913"), French ("Obligation de Frs. 505"), Russian ("Облигацiя въ 193.40 Рублей"), and Japanese. Features a central pastoral vignette. Signed by Hu Weide (胡惟德), Chinese Minister in Berlin, and co-signed by the consortium banks. Large Chinese Republic official seal in red. Issued in five international financial centres.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Chinese Government / Six-Power Consortium',
    issueDate: '1913',
    currency: 'British pound sterling|French franc|German mark|Russian ruble|Japanese yen',
    language: 'English|French|German|Russian|Japanese',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Reorganisation Gold Loan 1913; £23,000,000 total; No. 489,731; Frs. 505 / £20; quintilingual; signed by Hu Weide; Six-Power Consortium; major Chinese republican bond'
  }
);

// Row 599: Chinese 5% Reorganisation Loan 1913, Reverse/Conditions (+ Baku City Loan coupons visible)
setDoc(599,
  'Chinese Government 5% Reorganisation Gold Loan 1913, No. 489,731, Reverse with Conditions and Amortization Table (with City of Baku 5% Loan 1910 Coupons)',
  'Reverse/back side of the Chinese Government 5% Reorganisation Gold Loan 1913 bearer bond No. 489,731, showing multilingual conditions text and a sinking fund/amortization table. At the right edge of the scan, coupons from a separate document are also visible: these belong to the "5% LOAN OF THE CITY OF BAKU, 1910 / EMPRUNT 5% DE LA VILLE DE BAKOU, 1910" (a municipal bond of Baku, capital of the Baku Governorate of Imperial Russia, now Azerbaijan). The Baku City 5% Loan coupons appear to be bilingual English/French.',
  {
    type: 'bond',
    subjectCountry: 'China|Azerbaijan',
    issuingCountry: 'China|Russia',
    creator: 'Chinese Government; City of Baku',
    issueDate: '1913',
    currency: 'British pound sterling|French franc',
    language: 'English|French|German|Russian|Japanese',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of Chinese 5% Reorganisation Loan 1913 No. 489,731; conditions and amortization table; ALSO: City of Baku (Bakou) 5% Loan 1910 coupons visible at right (separate document scanned together)'
  }
);

// Row 600: Colombian National Railway Company Limited, First Mortgage Debenture No. 2081, £100 at 6%
setDoc(600,
  'Colombian National Railway Company Limited, First Mortgage Debenture No. 2081, £100 at 6% per annum, 1909',
  'First Mortgage Debenture No. 2081, £100 of The Colombian National Railway Company, Limited (incorporated under the Companies\' Acts 1862-1886). Share capital £900,000 (900,000 ordinary shares of £1 each). Total first mortgage debentures issued: £200,000, carrying interest at 6% per annum. The debenture is payable on January 1, 1929 (or on such earlier date as the principal moneys become payable). Interest payable semi-annually on January 1 and July 1. Given under the Company\'s Common Seal. Dated January 1, 1909. Signed by two directors and the company secretary.',
  {
    type: 'bond',
    subjectCountry: 'Colombia',
    issuingCountry: 'United Kingdom',
    creator: 'Colombian National Railway Company Limited',
    issueDate: 'January 1, 1909',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'British-registered company; £200,000 total debenture issue; No. 2081; £100 at 6%; Colombian railway concession; repayable 1929'
  }
);

// Row 601: Compagnie de Colonisation Américaine, Action 100 Acres Virginia/Kentucky, Série B No. 4258
setDoc(601,
  'Compagnie de Colonisation Américaine, Action de 100 Acres, Virginie & Kentucky, Série B No. 4258, Frs. 1,300',
  'Share certificate (Action) of the Compagnie de Colonisation Américaine (American Colonization Company, Paris), Series B, No. 4258, entitling the bearer to 100 acres of land in the states of Virginia and Kentucky, valued at Francs 1,300. The Compagnie de Colonisation Américaine was a French land company selling American land rights to European investors. The certificate details the rights of the shareholder to participate in the land development and distribution of proceeds. Signed by the company\'s gérant (managing director). Dividend coupon strips attached at left. The share dates from the mid-19th century.',
  {
    type: 'share',
    subjectCountry: 'United States',
    issuingCountry: 'France',
    creator: 'Compagnie de Colonisation Américaine',
    issueDate: 'ca. 1845-1860',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'French land company; 100 acres in Virginia and Kentucky; Série B No. 4258; Frs. 1,300; dividend coupon strips attached at left; mid-19th century American land speculation'
  }
);

// Row 602: Compañia Azucarera Baragua (Baragua Sugar Company), First Mortgage 7.5% Gold Bond No. M 728, $1,000
setDoc(602,
  'Compañia Azucarera Baragua (Baragua Sugar Company), First Mortgage 7.5% Sinking Fund Gold Bond No. M 728, $1,000',
  'First Mortgage, Fifteen-Year, Seven-and-One-Half Per Cent Sinking Fund Gold Bond No. M 728, $1,000 of Compañia Azucarera Baragua / Baragua Sugar Company. The bond is secured by a first mortgage on all properties of the Cuban sugar company. Features a central vignette depicting sugar cane cultivation and harvesting, with workers in a tropical setting, surrounded by an ornate green decorative border. Signed by Secretary (T.C. Loring) and Trustee (William M. Smith). A pink Cuban fiscal revenue stamp is affixed at lower left.',
  {
    type: 'bond',
    subjectCountry: 'Cuba',
    issuingCountry: 'Cuba',
    creator: 'Compañia Azucarera Baragua',
    issueDate: 'ca. 1920-1930',
    currency: 'United States dollar',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Cuban sugar company; $1,000 bond No. M 728; 7.5% first mortgage 15-year sinking fund gold bond; sugar harvest vignette; green border; Cuban fiscal stamp; Baragua, Cuba'
  }
);

// Row 603: Conditiën van een Fonds of Negotiatie, Theodore Passalaigue en Zoon, Suriname, Amsterdam
setDoc(603,
  'Conditiën van een Fonds of Negotiatie, onder Directie van Theodore Passalaigue en Zoon, Suriname, Amsterdam',
  'Printed conditions document (No. 306) for a plantation negotiatie fund under the direction of De Heeren Theodore Passalaigue en Zoon (Theodore Passalaigue and Son), merchants in Amsterdam, established for the benefit of planters in the Colony of Suriname. Article I states that the plantations of participating planters must be properly pledged as security, free of encumbrances, and a full inventory submitted. Article II requires planters to pledge a mortgage on all plantation production, pay interest annually at 10% per cent, and assigns a commission of van Justitie de Colonie Surinamen. Consistent with Dutch Suriname plantation negotiatie financing of the late 18th century.',
  {
    type: 'negotiatie',
    subjectCountry: 'Netherlands|Suriname',
    issuingCountry: 'Netherlands',
    creator: 'Theodore Passalaigue en Zoon',
    issueDate: 'ca. 1785-1800',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 18th century',
    notes: 'Suriname plantation negotiatie conditions; No. 306; Theodore Passalaigue & Zoon, Amsterdam; Article I-II of conditions; 10% interest; plantation mortgage; late 18th century'
  }
);

// Row 604: Confederate States of America, $1,000 Bond, Act of Congress April 30, 1863
setDoc(604,
  'Confederate States of America, $1,000 Bond, Loan Authorized by Act of Congress, April 30, 1863',
  '$1,000 bearer bond of the Confederate States of America, issued under the Loan Authorized by Act of Congress Approved April 30, 1863. Features the Confederate States of America title in ornate script at top, a central portrait of a Confederate official (possibly Secretary of the Treasury Christopher Memminger or another Confederate leader), and $1,000 / ONE THOUSAND DOLLARS denomination prominently displayed. Printed in black on gray/white paper. Dated June 1, 1863 (approximately). Signed by the Register of the Treasury in pursuance of the Act of Congress. This is one of the major Confederate war finance bond issues of the American Civil War.',
  {
    type: 'bond',
    subjectCountry: 'United States',
    issuingCountry: 'Confederate States of America',
    creator: 'Confederate States of America, Treasury Department',
    issueDate: 'June 1, 1863',
    currency: 'Confederate dollar',
    language: 'English',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Confederate States of America; $1,000 bond; Act of Congress April 30, 1863; American Civil War finance; portrait of Confederate official; Register of the Treasury signature'
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 585–604 (20 documents, batch21).');
