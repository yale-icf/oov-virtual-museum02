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

// Row 635: Republica Boliviana, Empréstito Interior, Vale de Cien Pesos, Chuquisaca, 1827-1828
setDoc(635,
  'Republica Boliviana, Empréstito Interior, Vale de Cien Pesos at 6%, Chuquisaca, 1827–1828',
  'Bolivian Republic domestic debt certificate (Vale) for 100 Pesos from the Empréstito Interior (Internal Loan) of one million pesos, authorized by the Constituent Congress law of November 16, 1826. Bears 6% annual rent payable in thirds in January, May, and September. Negotiable and accepted by republican treasuries at face value for purchase of public properties and reduction of censos, per decrees of June 12, 1827. Threatens death penalty for counterfeiters ("Pena de Muerte al Falsificador y Cómplices"). Issued at Chuquisaca (now Sucre), July 24, 1827. Departmental annotation from Chuquisaca, February 7, 1828 (the 17th year of Independence). Endorsed to Ramon Molina, annotated at folio No. 22. Signed by the Ministro de Hacienda, El Prefecto, and El Administrador del Tesoro.',
  {
    type: 'bond',
    subjectCountry: 'Bolivia',
    issuingCountry: 'Bolivia',
    creator: 'República Boliviana, Ministerio de Hacienda',
    issueDate: 'July 24, 1827',
    currency: 'Bolivian peso',
    language: 'Spanish',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Bolivian domestic loan; 100 pesos; 6% annual; payable in thirds Jan/May/Sep; Chuquisaca 1827; law of November 16, 1826; endorsed to Ramon Molina; death penalty for counterfeiting'
  }
);

// Row 636: Bolivian Government Loan of 1872, Trust Certificate, Bond No. 242, £400, June 1880
setDoc(636,
  'Bolivian Government Loan of 1872, Trust Certificate for Bond No. 242, £400, June 1880',
  'Trust certificate issued by John Horatio Lloyd and Alfred James Lambert, Trustees of the net proceeds of the 5% Government Loan raised in 1872 by the Republic of Bolivia, certifying that the holder is entitled to a rateable proportion of the ultimate balance of the trust fund. The trust was established from proceeds intended for payment of public works as described in the original loan prospectus. Covers Bond No. 242, dated February 12, 1872, for £400. Payment of the ultimate proportion of the trust fund balance will be notified by advertisement when it becomes available for distribution. Dated June 1880. A note at the bottom references the judgment of Her Majesty\'s Court of Appeal (June 1879) in the action "Wilson v. Church and others" and the House of Lords judgment (March 1880) in the action "National Bolivian Navigation Company and others v. Wilson."',
  {
    type: 'certificate',
    subjectCountry: 'Bolivia',
    issuingCountry: 'United Kingdom',
    creator: 'Trustees: John Horatio Lloyd & Alfred James Lambert',
    issueDate: 'June 1880',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Bolivian 5% Government Loan 1872; trust certificate; Bond No. 242, £400, dated February 12, 1872; trustees Lloyd & Lambert; legal references to Wilson v. Church (Court of Appeal 1879, House of Lords 1880)'
  }
);

// Row 637: Republica Boliviana, Fondo Publico, Vale Seis Pesos, Sucre, April 1, 1846
setDoc(637,
  'Republica Boliviana, Fondo Publico, Vale de Seis Pesos (Capital Cien Pesos), Sucre, April 1, 1846',
  'Bolivian Republic public fund certificate (Vale) No. 2657 for 6 Pesos, representing the income on a capital of 100 Pesos, payable in 402 thirds (tercias partes). The Fondo Publico (Public Fund) was instituted by the Law of June 1, 1843. Issued at Sucre (capital of Bolivia), April 1, 1846. Signed by the Ministro de la Corte Suprema, the Ministro de Hacienda, and the Presidente del Consejo. Endorsed to C° Venancio Nair; annotated in register book No. 70.',
  {
    type: 'bond',
    subjectCountry: 'Bolivia',
    issuingCountry: 'Bolivia',
    creator: 'República Boliviana, Ministerio de Hacienda',
    issueDate: 'April 1, 1846',
    currency: 'Bolivian peso',
    language: 'Spanish',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: 'Fondo Publico; No. 2657; 6 pesos income on 100 peso capital; 402 thirds payable; Law of June 1, 1843; Sucre April 1, 1846; endorsed to Venancio Nair; libro 70'
  }
);

// Row 638: Kingdom of Yugoslavia Public Works Loan, 4% Six-Language Bearer Bond, No. 05,155
setDoc(638,
  'Kingdom of Yugoslavia, 4% Public Works Loan, Six-Language Bearer Bond No. 05,155',
  'Six-language bearer bond No. 05,155 of the Kingdom of Yugoslavia\'s 4% Public Works Loan. The bond title appears in six languages: Serbo-Croatian (OBVEZNICA), Serbian Cyrillic (ОБВЕЗНИЦА), German (SCHULDVERSCHREIBUNG), Hungarian (KÖTVÉNY), French (OBLIGATION), and English (BOND). The multi-language format reflects the multi-ethnic character of the Yugoslav state and the bond\'s distribution across European financial markets. A revenue stamp is affixed at upper left. The bond carries 4% annual interest.',
  {
    type: 'bond',
    subjectCountry: 'Yugoslavia',
    issuingCountry: 'Yugoslavia',
    creator: 'Kingdom of Yugoslavia',
    issueDate: 'ca. 1928-1933',
    currency: 'Yugoslav dinar',
    language: 'Serbo-Croatian|German|Hungarian|French|English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Yugoslav 4% Public Works Loan; No. 05,155; six languages (Serbo-Croatian, Cyrillic Serbian, German, Hungarian, French, English); Kingdom of Yugoslavia; revenue stamp'
  }
);

// Row 639: Austro-Hungarian 4½% Bond, 1,000 Kronen, Five Languages, Series C, No. 17,496
setDoc(639,
  'Austro-Hungarian 4½% Bearer Bond, 1,000 Kronen / Mille Couronnes, Five Languages, Series C, No. 17,496',
  'Five-language 4½% bearer bond (Schuldverschreibung / Obveznica / Обвезника / Kötvény / Obligation), Series C, No. 17,496, face value 1,000 Kronen / Tausend Kronen / Ezer Korona / Mille Couronnes / Тысяча Крон. The denomination is expressed in five languages, and the bond title in five languages covering Croatian, Cyrillic Serbian, German, Hungarian, and French. The denomination in Kronen (the Austro-Hungarian currency) and the five-language format are characteristic of public bonds issued in the multi-ethnic lands of the Austro-Hungarian Crown, likely from Croatia-Slavonia, Bosnia-Herzegovina, or another composite territory. Beautifully printed in blue and silver with ornamental borders.',
  {
    type: 'bond',
    subjectCountry: 'Austria|Croatia|Hungary',
    issuingCountry: 'Austria',
    creator: 'Austro-Hungarian Government or Crown Land Authority',
    issueDate: 'ca. 1900-1914',
    currency: 'Austro-Hungarian krone',
    language: 'Croatian|Serbian|German|Hungarian|French',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Austro-Hungarian bond; Series C No. 17,496; 4.5%; 1,000 Kronen; five-language: Croatian/Cyrillic/German/Hungarian/French; blue and silver printing; likely Crown Land loan'
  }
);

// Row 640: Socialist Republic of Serbia, 100 Dinar Bearer Bond, Series Φ, No. 0324809
setDoc(640,
  'Socialist Republic of Serbia, Economic Development Loan, 100 Dinar Bearer Bond, Series Φ, No. 0324809',
  'Yugoslav communist-era bearer bond (Обвезница / Obveznica) for 100 Dinars of the Socialist Republic of Serbia (Социјалистичка Република Србија), Series Φ (Phi), No. 0324809, issued for economic development loans (За уплаћени износ зајма за привредни развој у Социјалистичкој Републици Србији). Payable to bearer (Платива доносиоцу). The coupon sheet on the left shows multiple numbered coupons (Серија Φ, Купон) for the same bond. Signed by the President of the Socialist Republic of Serbia.',
  {
    type: 'bond',
    subjectCountry: 'Yugoslavia',
    issuingCountry: 'Yugoslavia',
    creator: 'Socijalistička Republika Srbija',
    issueDate: 'ca. 1965-1980',
    currency: 'Yugoslav dinar',
    language: 'Serbian',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'SFR Yugoslavia; Socialist Republic of Serbia; economic development loan; 100 dinars; Series Φ No. 0324809; bearer bond; coupon sheet attached'
  }
);

// Row 641: Socialist Republic of Bosnia and Herzegovina, Employment Loan, 10,000 Dinar Bond, No. CC 04238717
setDoc(641,
  'Socialist Republic of Bosnia and Herzegovina, Employment Loan (Zajam za Zapošljavanje), 10,000 Dinars, No. CC 04238717',
  'Yugoslav communist-era bearer bond (Обвезница / Obveznica) for 10,000 Dinars (Deset Hiljada Dinara) of the Socialist Republic of Bosnia and Herzegovina (Социјалистичка Република Босна и Херцеговина), issued as an Employment Loan (Zajam za Zapošljavanje / Зајма за Запошљавање), No. CC 04238717. The bond features a photographic portrait (youth or workers) on the right face. Attached annuity coupon sheet (Ануитетски Купон / Anujtetski Kupon) on the left shows four coupons of 3,190 Dinars each. Bilingual in Serbian (Cyrillic) and Bosnian/Croatian (Latin script).',
  {
    type: 'bond',
    subjectCountry: 'Yugoslavia',
    issuingCountry: 'Yugoslavia',
    creator: 'Socijalistička Republika Bosna i Hercegovina',
    issueDate: 'ca. 1970-1985',
    currency: 'Yugoslav dinar',
    language: 'Bosnian|Serbian',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'SFR Yugoslavia; Socialist Republic of Bosnia & Herzegovina; Employment Loan (Zajam za Zapošljavanje); 10,000 dinars; No. CC 04238717; 4 annuity coupons of 3,190 dinars; photo portrait'
  }
);

// Row 642: Imperial Brazilian 5% Loan Conversion Circular, N.M. Rothschild & Sons, £20,000,000 at 4%, 1888
setDoc(642,
  'Imperial Brazil, Conversion and Redemption Circular for 5% Loans of 1865–1886, £20,000,000 at 4%, N.M. Rothschild & Sons, 1888',
  'Printed circular document announcing the Conversion and Redemption of the Imperial Brazilian Five Per Cent Loans of 1865, 1871, 1875, and 1886. Issue of £20,000,000 in new 4 per cent Brazilian Bonds, proceeds to be applied exclusively to conversion and redemption of the above loans. Issued under authority of Emperor of Brazil under Laws No. 3,329 (September 3, 1884) and Nos. 3,396 and 3,397 (November 24, 1888). N.M. Rothschild & Sons announce readiness to receive subscriptions. Conversion terms: 4% Bonds issued at 90, exchanged for old 5% bonds at various premiums; interest begins from first half-yearly coupon after January 1, 1889. Old bonds not presented for conversion will bear 3% from October 1, 1889.',
  {
    type: 'bond',
    subjectCountry: 'Brazil',
    issuingCountry: 'Brazil',
    creator: 'N.M. Rothschild & Sons (agents for Imperial Brazil)',
    issueDate: '1888',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Imperial Brazil debt conversion circular; £20,000,000 new 4% bonds; converts 5% loans of 1865, 1871, 1875, 1886; N.M. Rothschild & Sons; Laws 3,329, 3,396 and 3,397 (1884/1888)'
  }
);

// Row 643: Hungarian Kingdom 4½% Járadékkölcsön, 480 Korona/408 Marks/504 Fcs/£20, No. 00,919, Budapest, 1913
setDoc(643,
  'A Magyar Korona Országai 4½% Járadékkölcsön (Hungarian Kingdom 4½% Annuity Loan), 480 Korona / £20, No. 00,919, Budapest, 1913',
  'Bearer state debt bond (Bemutatóra Szóló Államadóssági Kötvény) of the Lands of the Hungarian Crown (A Magyar Korona Országai), 4½% Annuity Loan (Járadékkölcsön), authorized under Act XIV of 1911, Act V of 1912, and Act LXVI of 1912. Face value: 480 Korona = 408 German Imperial Marks = 504 French Francs = 20 Pounds Sterling. No. 00,919 (German market, "Abgestempelt"). Interest payable annually on April 1, in Budapest (in Korona), in Frankfurt, Brussels, and Zürich (in Marks), in Paris (in Francs), and in London (in Pounds Sterling). Redeemable by lottery from 1923. Kelt Budapesten, 1913. március 12-én. Signed by the Hungarian Royal Treasury (m. kir. pénzügyminisztérium) and the Hungarian State Debt Administration.',
  {
    type: 'bond',
    subjectCountry: 'Hungary',
    issuingCountry: 'Hungary',
    creator: 'A Magyar Korona Országai (Lands of the Hungarian Crown)',
    issueDate: 'March 12, 1913',
    currency: 'Austro-Hungarian krone|German mark|French franc|British pound sterling',
    language: 'Hungarian|German',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Hungarian Kingdom 4.5% Annuity Loan; No. 00,919; 480 Korona / 408 Marks / 504 Fcs / £20; Acts XIV-1911, V-1912, LXVI-1912; Budapest March 12, 1913; lottery redemption from 1923; Abgestempelt (German market stamp)'
  }
);

// Row 644: City of Budapest 4% Loan 1911, 400 Korona/420 Francs Bearer Bond, No. 135202
setDoc(644,
  'Budapest Székesfőváros (City of Budapest) 4% Loan 1911, Kötelezvény/Obligation No. 135202, 400 Korona / 420 Francs',
  'Bilingual Hungarian/French bearer bond (Kötelezvény / Obligation) No. 135202 of the Budapest Székesfőváros (Royal Capital City of Budapest) 4% Loan of 1911. Face value: 400 Korona (Hungarian) or 420 French Francs. Total loan: 100,000,000 Korona / 105,000,000 French Francs, divided into 28,000 bonds of 420 Francs. Interest at 4% per annum payable May 1 and November 1. Features an engraved vignette of the Chain Bridge (Széchenyi Lánchíd) over the Danube. Budapest, May 1, 1911. Signed by the Mayor (Főpolgármester) and City Comptroller (Számvevő Főnök).',
  {
    type: 'bond',
    subjectCountry: 'Hungary',
    issuingCountry: 'Hungary',
    creator: 'Budapest Székesfőváros (City of Budapest)',
    issueDate: 'May 1, 1911',
    currency: 'Austro-Hungarian krone|French franc',
    language: 'Hungarian|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Budapest municipal 4% loan 1911; No. 135202; 400 Korona / 420 Francs; total 100,000,000 Korona; Chain Bridge vignette; interest May 1 & Nov 1; bilingual Hungarian/French'
  }
);

// Row 645: Province of Buenos Aires, 5% Consolidation Gold Loan 1915, Converted Bond No. 048151, £20/Frs.504
setDoc(645,
  'Province of Buenos Aires, 5% Consolidation Gold Loan of 1915, Converted Bearer Bond No. 048151, £20 / Frs. 504',
  'Bilingual English/French converted bearer bond No. 048151 of the Province of Buenos Aires 5% Consolidation Gold Loan of 1915, face value £20 Sterling or Frs. 504 (= 504 French Francs). Marked "CONVERTED BOND." Printed in blue with the Provincial arms of Buenos Aires. The lower portion contains bilingual conditions text in English (PROVINCE OF BUENOS AIRES) and French. Signed by the Gobernador de la Provincia de Buenos Aires. The Consolidation Loan of 1915 was used to refund earlier Argentine provincial debts.',
  {
    type: 'bond',
    subjectCountry: 'Argentina',
    issuingCountry: 'Argentina',
    creator: 'Province of Buenos Aires',
    issueDate: '1915',
    currency: 'British pound sterling|French franc',
    language: 'English|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: '5% Consolidation Gold Loan 1915; No. 048151; £20 / Frs. 504; CONVERTED BOND; bilingual English/French; Provincial arms; Buenos Aires provincial debt consolidation'
  }
);

// Row 646: Egyesült Izzólámpa és Villamossági Részvénytársaság (Tungsram), 10 Shares at 100 Pengő, 1927–1935
setDoc(646,
  'Egyesült Izzólámpa és Villamossági Részvénytársaság (Tungsram), 10 Részvény of 100 Pengő Each, Nos. 107,151–107,160, Újpest',
  'Hungarian share certificate for 10 ordinary shares (Részvény) of 100 Pengő each (total 1,000 Pengő / Egyezer Pengőről) of the Egyesült Izzólámpa és Villamossági Részvénytársaság (United Incandescent Lamp and Electrical Company), internationally known as Tungsram. Certificate Nos. 107,151–107,160. The company was founded in 1896 in Újpest (near Budapest) and became one of the world\'s leading manufacturers of tungsten incandescent light bulbs. Features an engraved vignette of the Újpest factory complex. Issued Újpest, ca. 1927–1935 (Pengő denomination introduced 1927).',
  {
    type: 'share',
    subjectCountry: 'Hungary',
    issuingCountry: 'Hungary',
    creator: 'Egyesült Izzólámpa és Villamossági Részvénytársaság (Tungsram)',
    issueDate: 'ca. 1927-1935',
    currency: 'Hungarian pengő',
    language: 'Hungarian',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Tungsram; 10 shares at 100 Pengő each; Nos. 107,151–107,160; Újpest factory vignette; world-leading tungsten lamp manufacturer; Hungarian industrial company'
  }
);

// Row 647: Companhia da Zambezia (Zambezi Company), Provisional Share Certificate, Lisbon, December 25, 1900
setDoc(647,
  'Companhia da Zambezia (Zambezi Company), Certificado Provisorio de Acções, Bearer, Lisbon, December 25, 1900',
  'Provisional bearer share certificate (Certificado Provisorio / Certificat Provisoire) of the Companhia da Zambezia (Zambezi Company / Compagnie du Zambèze), Lisbon. Trilingual: Portuguese (Portador), French (Porteur), English (Bearer). The company was a Portuguese colonial concessionaire with the right to administer the Zambezia region of Mozambique. Capital: 2,700,000$000 Réis / 15,000,000 Francs / £600,000, divided into 600,000 shares of 4$500 Réis / 25 Francs / £1. Shares issued in four series (1898–1900). This provisional certificate was issued pending delivery of definitive share certificates. Lisbon, December 25, 1900. Signed by the Administradores (Board of Directors).',
  {
    type: 'share',
    subjectCountry: 'Mozambique',
    issuingCountry: 'Portugal',
    creator: 'Companhia da Zambezia',
    issueDate: 'December 25, 1900',
    currency: 'Portuguese real|French franc|British pound sterling',
    language: 'Portuguese|French|English',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Provisional share certificate; Companhia da Zambezia; Mozambique (Portuguese East Africa) concessionaire; capital £600,000; 600,000 shares at £1; four series 1898-1900; Lisbon December 25, 1900'
  }
);

// Row 648: Principality of Bulgaria, 6% State Mortgage Loan 1892, Two Obligations 500 Leva/Fcs., Nos. 073,741–073,742
setDoc(648,
  'Principality of Bulgaria, 6% State Mortgage Loan (Emprunt Hypothécaire d\'État) 1892, Two Obligations of 500 Leva / Cinq Cents Francs Each, Nos. 073,741–073,742',
  'Certificate for TWO bearer obligations (Две Облигации / Deux Obligations) Nos. 073,741 and 073,742 of the Principality of Bulgaria (Княжество България / Principauté de Bulgarie) 6% State Mortgage Loan of 1892 (Emprunt Hypothécaire d\'État 6% de l\'année 1892). Each obligation is for 500 Leva = Cinq Cents Francs = 400 Austrian Florins = 810,000 German Marks (per applicable rate). Total issue: 688,500 obligations = total of Fr. 344,250,000. Bilingual Bulgarian and French. Printed in green on cream paper. Signed by Finance Minister Leon Sabbaitscheff (Лeonъ Саббайчефъ) and another official. Sofia, November 15, 1892.',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Principality of Bulgaria, Ministry of Finance',
    issueDate: 'November 15, 1892',
    currency: 'Bulgarian lev|French franc',
    language: 'Bulgarian|French',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Bulgaria 6% State Mortgage Loan 1892; TWO obligations Nos. 073,741–073,742; 500 leva / 500 francs each; 688,500 total obligations; signed Leon Sabbaitscheff; green/cream printing; Sofia November 15, 1892'
  }
);

// Row 649: Bulgarian 6% State Mortgage Loan 1892, Horizontal Coupon Strip
setDoc(649,
  'Principality of Bulgaria, 6% State Mortgage Loan 1892, Horizontal Interest Coupon Strip',
  'Horizontal coupon strip for the Principality of Bulgaria 6% State Mortgage Loan of 1892. Contains a row of rectangular semi-annual interest coupons arranged in a long horizontal format, each showing the bond number and coupon number. This strip format was typical for Bulgarian government bonds of the period. The coupons enabled holders to collect the 6% annual interest in two semi-annual instalments. Part of the same bond document as the companion two-obligation certificate (see related document).',
  {
    type: 'bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Principality of Bulgaria, Ministry of Finance',
    issueDate: '1892',
    currency: 'Bulgarian lev|French franc',
    language: 'Bulgarian|French',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Coupon strip for Bulgaria 6% State Mortgage Loan 1892; horizontal format with rectangular coupons in row; semi-annual interest coupons'
  }
);

// Row 650: Chrysler Corporation, Dutch Certificaat SPECIMEN, 10 Ordinary Shares, Amsterdam, ca. 1926
setDoc(650,
  'Chrysler Corporation, Dutch Certificaat voor Tien Gewone Aandelen, SPECIMEN, Amsterdam, ca. 1926–1930',
  'Dutch specimen bearer certificate (Certificaat / SPECIMEN) for ten ordinary shares (without nominal value) of Chrysler Corporation (incorporated under the laws of the State of Delaware, U.S.A.). The shares are registered in the name of the Administratietrust established by Hubrecht, van Harrenraegel & Van Vlisser N.V., Amsterdam, under a notarial deed of July 16, 1926 by Notaris R. Nicolai. The certificate enables Dutch investors to hold Chrysler Corporation ordinary shares through a Dutch bearer instrument. Conditions for exchange into actual Chrysler shares in America are outlined. Marked SPECIMEN. Amsterdam, ca. 1926–1930. Dividend coupon sheet (coupons 29–40 and TALON) attached at bottom.',
  {
    type: 'share',
    subjectCountry: 'United States',
    issuingCountry: 'Netherlands',
    creator: 'Hubrecht, van Harrenraegel & Van Vlisser N.V., Amsterdam (Administratietrust)',
    issueDate: 'ca. 1926-1930',
    currency: 'United States dollar',
    language: 'Dutch',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'SPECIMEN; Chrysler Corporation; Dutch bearer certificaat for 10 ordinary shares; Hubrecht, van Harrenraegel & Van Vlisser N.V.; notarial deed July 16, 1926; dividend coupons 29–40 and talon attached'
  }
);

// Row 651: Republic of Colombia, 4% Funding Certificate Scrip, Lazard Brothers, September 1, 1934
setDoc(651,
  'Republic of Colombia, 4% Funding Certificate Scrip, No. 5131/1882, Lazard Brothers & Co. Ltd., September 1, 1934',
  'Bearer scrip certificate No. 5131/1882 for a nominal amount from the Republic of Colombia\'s 4% Funding Certificates, 1934/46, issued by Lazard Brothers & Co. Ltd. (41, Old Broad Street, London, E.C.2), limited to a total nominal amount not exceeding £155,892 14s. 0d. as announced in "The Times" of February 22, 1934. The Funding Certificates covered outstanding obligations under multiple old Colombian loans: 5% Bogotá-Sahana Railway Loan 1906, 8% External Gold Loan 1911, 8% External Debt 1913, 8% External Debt 1913 (French Issue), 5% (1920) Bonds, and Agricultural Mortgage Bank of Colombia 6% Guaranteed Mortgage Bonds. Dated September 1, 1934.',
  {
    type: 'certificate',
    subjectCountry: 'Colombia',
    issuingCountry: 'United Kingdom',
    creator: 'Lazard Brothers & Co. Ltd. (for Republic of Colombia)',
    issueDate: 'September 1, 1934',
    currency: 'British pound sterling',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Colombia 4% Funding Certificates 1934/46; scrip No. 5131/1882; Lazard Brothers & Co. London; total issue £155,892 14s.; covers 6 old Colombian loans including Bogotá-Sahana Railway (1906) and others; September 1, 1934'
  }
);

// Row 652: Republic of Cuba, 4½% External Gold Bond Due 1949 SPECIMEN, $500,000, Havana, 1909
setDoc(652,
  'Republic of Cuba, 4½% External Gold Bond Due 1949 SPECIMEN, $500,000 US Gold / £102,880.13.2 / M 2,100,000 / 2,590,000 Fcs., Havana, August 2, 1909',
  'Specimen external gold bond (SPECIMEN) of the Republic of Cuba\'s 4½% Gold Bond External Loan due 1949, No. F 2124 / No. 00000. Face value in four currencies: $500,000 US Gold / £102,880.13.2 Sterling / M 2,100,000 German Marks / 2,590,000 French Francs. This large-denomination general bond represents a temporary bond issued in pursuance of a decree of the President of Cuba and agreement between the Republic of Cuba and Speyer & Co., dated January 24, 1901, for a loan of $16,500,000. The bond is exempt from all Cuban taxes and may be drawn for redemption at par with interest. Issued in Havana, Cuba, August 2, 1909.',
  {
    type: 'bond',
    subjectCountry: 'Cuba',
    issuingCountry: 'Cuba',
    creator: 'Republic of Cuba (agent: Speyer & Co.)',
    issueDate: 'August 2, 1909',
    currency: 'United States dollar|British pound sterling|German mark|French franc',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'SPECIMEN; Cuba 4.5% External Gold Bond due 1949; No. F 2124 / 00000; $500,000 US Gold; £102,880 / M 2,100,000 / 2,590,000 Fcs.; Speyer & Co. loan $16,500,000; Havana August 2, 1909'
  }
);

// Row 653: Czechoslovak State 8% Bond 1922, $100 (Cancelled)
setDoc(653,
  'Czechoslovak State 8% Bond 1922, $100, CANCELLED',
  'US Dollar-denominated bearer bond of the Czechoslovak State (or a Czechoslovak state enterprise), 8% interest, 1922, face value $100. The bond is stamped CANCELLED (annulled) with a large red overprint. The bond\'s legal basis and conditions are printed in English (and possibly Czech). The Czechoslovak State issued dollar-denominated bonds on the US market in the early 1920s to raise foreign currency capital for post-WWI reconstruction. The cancellation suggests this bond was redeemed or annulled during a debt restructuring.',
  {
    type: 'bond',
    subjectCountry: 'Czechoslovakia',
    issuingCountry: 'Czechoslovakia',
    creator: 'Czechoslovak State',
    issueDate: '1922',
    currency: 'United States dollar',
    language: 'English|Czech',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Czechoslovak State 8% bond 1922; $100; CANCELLED (large red stamp); US dollar-denominated; post-WWI reconstruction bond; English and Czech text'
  }
);

// Row 654: Egyptian Credit Foncier (Crédit Foncier Egyptien), Share £20/Frs.500, No. 5,000,246, Cairo, 1904
setDoc(654,
  'Egyptian Credit Foncier / Crédit Foncier Egyptien, Bearer Share £20 or Frs. 500, No. 5,000,246, Cairo, January 1904',
  'Bilingual English/Arabic/French bearer share No. 5,000,246 of the Egyptian Credit Foncier (Crédit Foncier Egyptien / البنك العقاري المصري), created by decree of H.H. the Khedive bearing date February 15, 1880. Société Anonyme. Capital: £4,000,000 or Frs. 100,000,000. Share value: £20 or Frs. 500 to bearer (Action au Porteur). Presented in both English and French with Arabic title text and decorative Moorish/Islamic arch motif framing the certificate. Cairo, January 1904. Has endorsement and transfer stamps visible at bottom. The Egyptian Credit Foncier provided real estate and agricultural mortgage credit in Egypt under the Khedival era.',
  {
    type: 'share',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Egyptian Credit Foncier / Crédit Foncier Egyptien',
    issueDate: 'January 1904',
    currency: 'British pound sterling|French franc',
    language: 'English|French|Arabic',
    numberPages: '1',
    period: 'Late 19th/Early 20th century',
    notes: 'Egyptian Credit Foncier; Khedival decree February 15, 1880; No. 5,000,246; £20 / Frs. 500 bearer share; capital £4,000,000; Cairo January 1904; Islamic arch motif; bilingual English/French with Arabic; transfer stamps'
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 635–654 (20 documents, batch23).');
