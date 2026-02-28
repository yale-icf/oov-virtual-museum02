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

// Row 565: Chinese Republic 5% Gold Industrial Loan 1914, Coupon Sheet (blue)
setDoc(565,
  'Chinese Republic 5% Gold Industrial Loan 1914, Coupon Sheet (Blue Series)',
  'Blue-printed coupon sheet for the Chinese Republic 5% Gold Industrial Loan of 1914 (Emprunt Industriel du Gouvernement de la République Chinoise, 5% Or 1914). Contains multiple semi-annual interest coupons arranged in a multi-column grid, each showing the obligation number in orange and the coupon amount. This sheet belongs to a different bond number in the same 150,000,000 Franc series as the bearer obligation shown separately.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: '1914',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Gold Industrial Loan 1914; 150,000,000 franc total issue; blue-printed semi-annual coupon sheet'
  }
);

// Row 566: British Company Bond Coupon Sheet (blue, likely Pekin Syndicate or New Russia Company)
setDoc(566,
  'British Company Bond, Coupon Sheet (Blue Series), Early 20th Century',
  'Blue-printed coupon sheet containing multiple rectangular dividend or interest coupons for a British-registered company bond or share. Each coupon bears the company name and a reference number. The format and blue printing is consistent with late 19th- to early 20th-century British corporate securities. The coupon arrangement follows standard English practice of the period.',
  {
    type: 'bond',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'British company (unidentified)',
    issueDate: 'ca. 1900-1910',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Blue-printed coupon sheet; British corporate bond or debenture; company name not fully legible at thumbnail resolution'
  }
);

// Row 567: Costa Rica Railway Company Limited, First Mortgage Debenture, £100, Conditions Reverse, CANCELLED
setDoc(567,
  'Costa Rica Railway Company Limited, First Mortgage Debenture £100, Conditions (Reverse), CANCELLED',
  'Reverse/conditions side of a Costa Rica Railway Company Limited First Mortgage Debenture, £100, marked CANCELLED. The conditions text outlines the principal terms of the debenture, including redemption at par, 5% interest, and first-charge security on the Company\'s property. The lower portion contains a printed transfer register table with spaces for recording successive registered holders. A blue "CANCELLED" stamp is prominently displayed.',
  {
    type: 'bond',
    subjectCountry: 'Costa Rica',
    issuingCountry: 'United Kingdom',
    creator: 'Costa Rica Railway Company Limited',
    issueDate: 'ca. 1886-1890',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'First mortgage debenture; £100; CANCELLED; reverse/conditions side; transfer register blank; British company operating Costa Rican railway'
  }
);

// Row 568: Costa Rica Railway Company Limited, Second Debenture No. 6886, £100, 5%, 1888
setDoc(568,
  'Costa Rica Railway Company Limited, Second Debenture No. 6886, £100 at 5%, 1888',
  'Front face of Costa Rica Railway Company Limited Second Debenture No. 6886, £100 denomination, bearing interest at 5% per annum. Issued as part of a total series of £600,000 in 6,000 Second Debentures of £100 each, numbered 6,551 to 12,550 inclusive. Features a central oval portrait of a bearded gentleman (possibly a company director or Costa Rican official). Interest payable on March 1 and September 1 each year. Dated February 1888. Signed by director and secretary.',
  {
    type: 'bond',
    subjectCountry: 'Costa Rica',
    issuingCountry: 'United Kingdom',
    creator: 'Costa Rica Railway Company Limited',
    issueDate: 'February 1888',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Second debenture No. 6886; £600,000 total issue; 6,000 bonds numbered 6,551–12,550; 5% p.a.; portrait vignette; British company operating Costa Rican railway'
  }
);

// Row 569: Costa Rica Railway Company Limited, Second Debenture No. 6886, Conditions/Reverse, CANCELLED
setDoc(569,
  'Costa Rica Railway Company Limited, Second Debenture No. 6886, £100, Conditions (Reverse), CANCELLED',
  'Reverse/conditions side of Costa Rica Railway Company Limited Second Debenture No. 6886, £100. The conditions text sets out terms for the Second Debentures including redemption provisions, interest payment, and priority. The lower portion contains a transfer register table; several transfer entries appear to be partially completed. Marked CANCELLED. Companion to the front face of the same debenture (No. 6886).',
  {
    type: 'bond',
    subjectCountry: 'Costa Rica',
    issuingCountry: 'United Kingdom',
    creator: 'Costa Rica Railway Company Limited',
    issueDate: 'February 1888',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Second debenture No. 6886; reverse/conditions side; CANCELLED stamp; transfer register with partial entries'
  }
);

// Row 570: Bética S.A. Cooperativa Agrícola Industrial, Obligación 500 Pesetas, Sevilla, 1920
setDoc(570,
  'Bética Sociedad Anónima Cooperativa Agrícola Industrial, Obligación Hipotecaria, 500 Pesetas, Sevilla, 1920',
  'Mortgage bond (Obligación Hipotecaria) of Bética Sociedad Anónima Cooperativa Agrícola Industrial, face value 500 Pesetas. Bética was constituted by public deed on December 15, 1919 in Seville, with a share capital of 17,000,000 Pesetas. This bond belongs to an issue of 4,000,000 Pesetas in mortgage obligations divided into 12,000 titles of 500 Pesetas nominal each, authorized at the shareholder general assembly of June 30, 1920, and formalized by public deed before notary Don Félix Sánchez Blanco y Sánchez, Seville, September 21, 1920. Features allegorical female figure with agricultural and industrial scene in background. Coupon sheet attached at bottom. Signed by El Presidente and El Secretario, Sevilla October 1, 1920.',
  {
    type: 'bond',
    subjectCountry: 'Spain',
    issuingCountry: 'Spain',
    creator: 'Bética Sociedad Anónima Cooperativa Agrícola Industrial',
    issueDate: 'October 1, 1920',
    currency: 'Spanish peseta',
    language: 'Spanish',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Agricultural-industrial cooperative; Seville; 4,000,000 pesetas mortgage bond issue; 12,000 titles at 500 pesetas; coupon sheet attached'
  }
);

// Row 571: Bética Bond, Reverse Side
setDoc(571,
  'Bética Sociedad Anónima Cooperativa Agrícola Industrial, Obligación Hipotecaria, Reverse Side',
  'Reverse/back side of the Bética Sociedad Anónima Cooperativa Agrícola Industrial mortgage bond (Obligación Hipotecaria), 500 Pesetas. Shows "500 PESETAS" in large latent-image style text and legal registration information in two columns, noting inscription in the Registro Mercantil de Sevilla (Seville Commercial Registry) on November 30, 1919, and the Registro de la Propiedad de Lora del Río (Land Registry of Lora del Río) on November 35 (sic), 1920. Printed in salmon/pink.',
  {
    type: 'bond',
    subjectCountry: 'Spain',
    issuingCountry: 'Spain',
    creator: 'Bética Sociedad Anónima Cooperativa Agrícola Industrial',
    issueDate: 'October 1, 1920',
    currency: 'Spanish peseta',
    language: 'Spanish',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of mortgage bond; "500 PESETAS" latent text; Registro Mercantil de Sevilla registration noted; salmon/pink printed reverse'
  }
);

// Row 572: Chinese Republic 5% Gold Industrial Loan 1914, Obligation 500 Francs, No. 026,111
setDoc(572,
  'Chinese Republic, Emprunt Industriel 5% Or 1914, Obligation 500 Francs au Porteur, No. 026,111, Paris',
  '500 Franc bearer bond (Obligation au Porteur) from the Chinese Republic Industrial Government Loan, 5% Gold 1914 (Emprunt Industriel du Gouvernement de la République Chinoise, 5% Or 1914), a total issue of 150,000,000 Francs in 300,000 obligations of 500 Francs each. Authorized by the President of the Chinese Republic. Bond No. 026,111. Exempt from all Chinese taxes. Term of 30 years. Issued in Paris, April 7, 1914. Features vignettes of Chinese buildings and red Chinese seals. Signed by Zhou Baiqi (周白齐) and Chinese finance officials. Coupon sheet attached at bottom.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: 'April 7, 1914',
    currency: 'French franc',
    language: 'French|Chinese',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Gold Industrial Loan 1914; 150,000,000 francs total; No. 026,111; 500 francs; Paris issue; 30-year term; exempt from Chinese taxes; coupon sheet attached'
  }
);

// Row 573: Chinese Republic 5% Gold Industrial Loan 1914, No. 026,111, Reverse with Chinese Seals
setDoc(573,
  'Chinese Republic, Emprunt Industriel 5% Or 1914, No. 026,111, Reverse Side with Chinese Seals',
  'Reverse/back side of Chinese Republic Industrial Government Loan 5% Gold 1914 bearer bond No. 026,111, 500 Francs. Shows faint mirror-image printing from the obverse face. Two prominent red Chinese official seals (chops) are affixed. The bond number 026,111 is printed in red. This reverse was typically left mostly blank except for official seals and bond number for authentication purposes.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: 'April 7, 1914',
    currency: 'French franc',
    language: 'Chinese|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of bearer bond No. 026,111; two Chinese official red seals affixed; bond number printed in red'
  }
);

// Row 574: Chinese Republic 5% Gold Industrial Loan 1914, Red Coupon Sheet
setDoc(574,
  'Chinese Republic, Emprunt Industriel 5% Or 1914, Coupon Sheet (Red Series)',
  'Large red-printed coupon sheet for the Chinese Republic Industrial Government Loan 5% Gold 1914. Contains numerous semi-annual interest coupons for 500 Franc obligations, each showing the bond number in red. The coupons are arranged in a multi-column grid on a large loose sheet. This red coupon series accompanies a different bond number within the same 150,000,000 Franc total issue.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: '1914',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Gold Industrial Loan 1914; red-printed coupon sheet; large format; semi-annual coupons'
  }
);

// Row 575: Chinese Republic 5% Gold Industrial Loan 1914, Coupon Sheet for No. 026,111
setDoc(575,
  'Chinese Republic, Emprunt Industriel 5% Or 1914, Coupon Sheet for Obligation No. 026,111',
  'Large format coupon sheet belonging to Chinese Republic Industrial Government Loan 5% Gold 1914 bearer bond No. 026,111, 500 Francs. Contains numerous semi-annual interest coupons arranged in a multi-column grid, each labeled "OBLIGATION N° 026,111" with the coupon number and payment amount. The bond number is printed in orange. Coupons are printed in brown/tan on white paper. This sheet was attached to the bearer bond and individual coupons were detached at each interest payment date.',
  {
    type: 'bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Chinese Republic',
    issueDate: '1914',
    currency: 'French franc',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Coupon sheet for obligation No. 026,111; 5% Gold Industrial Loan 1914; bond number in orange; semi-annual interest coupons'
  }
);

// Row 576: Republic of Peru 5% Gold Bonds 1920, Bond No. 71547, £10 Sterling
setDoc(576,
  'Republic of Peru, 5% Gold Bonds of 1920 for £720,620, Bond to Bearer No. 71547, £10 Sterling',
  'Bilingual English/French bearer bond No. 71547 of £10 Sterling from the Republic of Peru\'s 5% Gold Bond issue of 1920 for a total of £720,620. "Issue of 72,062 Bonds of £10 each, numbered 1 to 72,062 inclusive." Interest at 5% per annum payable semi-annually. Issued by the Embiricos Syndicate Limited as agents for the Peruvian government. The bond features the Peruvian national coat of arms and bilingual conditions in English and French. Signed by the Peruvian Minister of Finance (Lewise H. Walker or similar). Dated May 1920.',
  {
    type: 'bond',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Republic of Peru',
    issueDate: 'May 1920',
    currency: 'British pound sterling',
    language: 'English|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Gold Bonds 1920; £720,620 total; 72,062 bonds at £10; No. 71547; issued by Embiricos Syndicate Limited; bilingual English/French'
  }
);

// Row 577: Republic of Peru 5% Gold Bond No. 71547, £10, Reverse with Bilingual Conditions
setDoc(577,
  'Republic of Peru, 5% Gold Bond No. 71547, £10, Reverse with Bilingual General Bond Conditions',
  'Reverse/back side of Republic of Peru 5% Gold Bond No. 71547, £10, showing the full bilingual conditions text: "REPUBLIC OF PERU. ISSUE OF £720,620 5% GOLD BONDS 1920. GENERAL BOND." (English) and "REPUBLIQUE DU PEROU. EMISSION DE £720,620 D\'OBLIGATIONS 5% OR DE 1920. OBLIGATION GÉNÉRALE." (French) in two parallel columns. Outlines the full terms of the loan, security, and redemption provisions.',
  {
    type: 'bond',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Republic of Peru',
    issueDate: 'May 1920',
    currency: 'British pound sterling',
    language: 'English|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of bond No. 71547; bilingual general bond conditions; English and French parallel columns'
  }
);

// Row 578: Republic of Peru 5% Gold Bond 1920, French Coupon Sheet (Republique du Perou)
setDoc(578,
  'Republic of Peru, 5% Gold Bonds 1920, Coupon Sheet (French: Republique du Perou), Coupons 1–51',
  'French-language coupon sheet for the Republic of Peru 5% Gold Bonds of 1920, labeled "REPUBLIQUE DU PEROU." Contains semi-annual interest coupons numbered 1 through approximately 51, each for £1.0.5.0 (one pound, zero shillings, five pence — equivalent to half the annual 5% interest on £10). Coupons are printed in green/black on white paper and arranged in a multi-row grid. These French-language coupons accompanied the bilingual bearer bonds.',
  {
    type: 'bond',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Republic of Peru',
    issueDate: 'May 1920',
    currency: 'British pound sterling',
    language: 'French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'French coupon sheet; "Republique du Perou" 5% Gold Bonds 1920; coupons 1–51; semi-annual interest coupons'
  }
);

// Row 579: Republic of Peru 5% Gold Bond 1920, English Coupon Sheet for No. 71547
setDoc(579,
  'Republic of Peru, 5% Gold Bond No. 71547, £10, English Coupon Sheet',
  'English-language coupon sheet for Republic of Peru 5% Gold Bond No. 71547, £10, labeled "R. OF PERU." Contains numbered semi-annual interest coupons for bond No. 71547, printed in black on white. Coupons are arranged in a dense grid and reference the 5% Gold Bonds 1920 issue and the January/July payment dates. This black-printed English coupon sheet is the companion to the green French coupon sheet (see related document).',
  {
    type: 'bond',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Republic of Peru',
    issueDate: 'May 1920',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'English coupon sheet for bond No. 71547; black printed; "R. of Peru" 5% Gold Bonds 1920; companion to French coupon sheet'
  }
);

// Row 580: Principality of Bulgaria, 5% State Gold Loan 1904, 500 Francs, No. 079,150
setDoc(580,
  'Principality of Bulgaria, Bulgarian 5% State Gold Loan 1904, 500 Francs, No. 079,150',
  'Quadrilingual bearer bond No. 079,150 of the Principality of Bulgaria (Княжество България) 5% State Gold Loan of 1904. Face value 500 Francs / Пятьсот Франковъ / Fünfhundert Francs / Five Hundred Francs — presented in Bulgarian, Russian, German, and English. Total issue of 200,000 obligations. The bond is secured on Bulgarian state revenues and features ornate orange-rust decorative borders with allegorical figures. Dated Sofia, April 21, 1904. Signed by the Bulgarian Finance Minister T. Baiçounov (Байчунов) and the Director of the Public Debt.',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Principality of Bulgaria',
    issueDate: 'April 21, 1904',
    currency: 'French franc',
    language: 'Bulgarian|French|German|English|Russian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Principality of Bulgaria (before kingdom proclamation 1908); No. 079,150; 500 francs; quadrilingual; 200,000 bonds total; orange/rust decorative printing'
  }
);

// Row 581: Bulgaria 5% State Gold Loan 1904, No. 079,150, Reverse with Quadrilingual Conditions
setDoc(581,
  'Principality of Bulgaria, 5% State Gold Loan 1904, No. 079,150, Reverse with Quadrilingual Conditions and Amortization Table',
  'Reverse/back side of the Principality of Bulgaria 5% State Gold Loan 1904 bearer bond No. 079,150. Shows loan conditions in four languages (Bulgarian "УСЛОВІЯ НА ЗАЕМА," French "CONDITIONS DE L\'EMPRUNT," German "BEDINGUNGEN DER ANLEIHE," English "CONDITIONS OF THIS LOAN") arranged in parallel columns, covering interest payments, redemption by lottery, and security. Below is a multilingual amortization table (ТАБЛИЦА НА ПОГАШЕНИЕ / TABLEAU D\'AMORTISSEMENT / TILGUNGS-PLAN / TABLE OF AMORTISATION) with redemption schedule by year.',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Principality of Bulgaria',
    issueDate: 'April 21, 1904',
    currency: 'French franc',
    language: 'Bulgarian|French|German|English|Russian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of bond No. 079,150; quadrilingual conditions; amortization table in four languages; lottery redemption schedule'
  }
);

// Row 582: Bulgarian Government Bond, Large Coupon Sheet
setDoc(582,
  'Bulgarian Government Bond, Large Format Coupon Sheet, Early 20th Century',
  'Large format coupon sheet for a Bulgarian government bond, containing numerous small interest coupons arranged in a large multi-column grid. The coupons are lightly printed in a pale color on white paper. The sheet is printed in a wide landscape or large portrait format covering an extensive redemption/interest schedule. Likely belongs to one of Bulgaria\'s international gold loans of the early 20th century.',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Bulgarian Government',
    issueDate: 'ca. 1904-1920',
    currency: 'French franc',
    language: 'French|Bulgarian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Large-format coupon sheet; Bulgarian government gold loan; pale/light printed coupons in grid'
  }
);

// Row 583: Kingdom of Bulgaria 5% Gold Loan, Coupon Sheet, No. 025,444
setDoc(583,
  'Kingdom of Bulgaria, 5% Or Gold Loan, Coupon Sheet, No. 025,444 (Royaume de Bulgarie)',
  'Interest coupon sheet for a Kingdom of Bulgaria (Royaume de Bulgarie) 5% Gold loan, bond No. 025,444, face value 500 Francs. The coupons are labeled "EMPRUNT BULGARE / ROYAUME DE BULGARIE / 5% OR / Coupon Européen" and "МИНИСТЕРСТВО НА ФИНАНСИТЕ" (Ministry of Finance). Bond number 025,444 appears in orange in each coupon. Coupons are arranged in a multi-column grid and printed in orange/black on cream. This loan is from after Bulgaria\'s proclamation as a Kingdom in 1908, distinct from the earlier Principality loans.',
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
    notes: '"Royaume de Bulgarie" = Kingdom of Bulgaria (after 1908); No. 025,444; 500 francs; "Coupon Européen"; Ministry of Finance; orange/black printing'
  }
);

// Row 584: Kingdom of Bulgaria Loan, Conditions Reverse with Quadrilingual Amortization Table
setDoc(584,
  'Kingdom of Bulgaria 5% Gold Loan, Reverse with Quadrilingual Conditions and Amortization Table',
  'Reverse/back side of a Kingdom of Bulgaria 5% Gold Loan bearer bond, showing conditions in four languages (Bulgarian "УСЛОВИЯ НА ЗАЕМА," French "CONDITIONS DE L\'EMPRUNT," German "BEDINGUNGEN DER ANLEIHE," English "CONDITIONS OF THE LOAN") in parallel columns, and an extensive multilingual amortization table (ТАБЛИЦА ЗА ПОГАШЕНИЕ / TABLEAU D\'AMORTISSEMENT / TILGUNGS-PLAN / TABLE OF AMORTISATION) with the annual redemption schedule. The quadrilingual format and condition structure is similar to but distinct from the 1904 Principality of Bulgaria bond.',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Kingdom of Bulgaria',
    issueDate: 'ca. 1909-1914',
    currency: 'French franc',
    language: 'Bulgarian|French|German|English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Reverse of Kingdom of Bulgaria bond; quadrilingual conditions and amortization table; "УСЛОВИЯ НА ЗАЕМА" (modernized spelling vs. 1904 Principality bond)'
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 565–584 (20 documents, batch20).');
