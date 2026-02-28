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

function setDoc(rowIdx, title, description, meta) {
  set(rowIdx, 'title', title);
  set(rowIdx, 'description', description);
  Object.entries(meta).forEach(([k, v]) => set(rowIdx, k, v));
}

// --- Row 376: Société des Grands Hôtels Indochinois, Saigon ---
setDoc(376,
  'Société des Grands Hôtels Indochinois: Action au Porteur de 50 Piastres (No. 01529, Saïgon)',
  'This printed share certificate is a bearer action (action au porteur) of 50 piastres in the Société des Grands Hôtels Indochinois (Grand Hotels of Indochina Company), a French colonial hospitality corporation operating luxury hotels in French Indochina. Capital 600,000 piastres. Certificate No. 01529. Issued in Saïgon (present-day Ho Chi Minh City, Vietnam). The company was one of several French colonial enterprises that sought to develop tourism and business travel infrastructure in Indochina during the early twentieth century.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Vietnam',
    issuingCountry: 'France',
    creator: 'Société des Grands Hôtels Indochinois',
    issueDate: '1920-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société des Grands Hôtels Indochinois, Action au porteur, 50 piastres, No. 01529. Capital 600,000 piastres. Saïgon, French Indochina, ca. 1920.',
  }
);

// --- Row 377: Great Atlantic & Pacific Tea Company (A&P) ---
setDoc(377,
  'Great Atlantic & Pacific Tea Company (A&P): 100 Shares, $1 Par Value (No. 458293)',
  'This engraved stock certificate represents 100 shares of $1 par value in The Great Atlantic & Pacific Tea Company (commonly known as A&P), incorporated in the State of Maryland. Certificate No. 458293. At its peak in the mid-twentieth century, A&P was the largest retail grocery chain in the United States and one of the largest retail businesses in the world, operating thousands of supermarkets. This certificate dates from the company\'s dominant era and represents a significant piece of American retail and corporate history.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Great Atlantic & Pacific Tea Company',
    issueDate: '1940-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Great Atlantic & Pacific Tea Company (A&P). 100 shares, $1 par value. No. 458293. Maryland corporation. Ca. 1940.',
  }
);

// --- Row 378: Hungarian Fund, New York, 1852 ---
setDoc(378,
  'Hungarian Fund: $1 Bond (New York, February 2, 1852)',
  'This small printed certificate is a $1 bearer bond of the Hungarian Fund, issued in New York on February 2, 1852, during the American tour of Lajos Kossuth, the exiled leader of the 1848–49 Hungarian Revolution. Following the crushing of the Hungarian independence movement by Austrian and Russian forces, Kossuth traveled to the United States in 1851–52 to raise money and political support for a renewed Hungarian bid for independence. The Hungarian Fund bonds were sold to American supporters of the Hungarian cause, predominantly Hungarian and German immigrants, at face values as low as $1 to enable broad participation. Though the hoped-for second revolution never materialized, these bonds represent a significant early episode in the history of diaspora finance and international political fundraising.',
  {
    type: 'Bond',
    subjectCountry: 'Hungary',
    issuingCountry: 'United States',
    creator: 'Hungarian Fund (Kossuth exile government)',
    issueDate: '1852-02-02',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Hungarian Fund $1 bearer bond. New York, February 2, 1852. Issued during Lajos Kossuth\'s American fundraising tour for Hungarian independence.',
  }
);

// --- Row 379: Imperial Russian Government 3% State Loan 1859, £100 ---
setDoc(379,
  'Imperial Russian Government: 3% State Loan of 1859, £100 Sterling (No. 31016)',
  'This is a bearer bond of £100 sterling in the Imperial Russian Government 3% State Loan of 1859. Certificate No. 31016. Issued as part of an external British-market loan denominated in pounds sterling, this bond formed part of a series of Russian government borrowings in London during the mid-nineteenth century to finance infrastructure and state expenditures. Interest was payable in London at 3% per annum. The bond is an example of Russian sovereign debt instruments designed to attract British capital.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1859-01-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Imperial Russian Government 3% State Loan of 1859. £100 bearer bond. No. 31016.',
  }
);

// --- Row 380: Romanian Public Debt, Internal Consolidation Loan 3% 1935 ---
setDoc(380,
  'Romanian Public Debt: Internal Consolidation Loan 3% 1935, Lei 5,000 (No. 242914)',
  'This printed bearer bond of 5,000 Romanian Lei represents the Internal Consolidation Loan of 1935, issued by the Romanian state as part of a Depression-era debt restructuring program. No. 242914. The bond carries 3% annual interest and was issued with an attached coupon sheet for periodic interest payments. It reflects the fiscal pressures faced by Romania during the economic crises of the 1930s, when many European governments were forced to consolidate and restructure their domestic debt obligations.',
  {
    type: 'Bond',
    subjectCountry: 'Romania',
    issuingCountry: 'Romania',
    creator: 'Romania. Ministry of Finance',
    issueDate: '1935-01-01',
    currency: 'ROL',
    language: 'Romanian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Romanian Public Debt, Internal Consolidation Loan 3% 1935. Lei 5,000 bearer bond. No. 242914. With coupon sheet.',
  }
);

// --- Row 381: Imperial Russian Government 6% Perpetual Bond, 5,000 Rubles, 1859 ---
setDoc(381,
  'Imperial Russian Government: 6% Perpetual Bond, 5,000 Rubles (1859)',
  'This large-denomination bearer bond of 5,000 rubles was issued by the State Commission for the Repayment of State Debts of the Imperial Russian Government, bearing 6% perpetual interest. Dated 1859. Perpetual bonds (often called "consols" after the British model) paid interest indefinitely without a fixed redemption date. This bond was part of Russia\'s substantial domestic funded debt in the mid-nineteenth century. The 5,000-ruble denomination indicates this was a wholesale instrument likely held by institutional investors or wealthy individuals.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government, State Commission for Debt Repayment',
    issueDate: '1859-01-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Imperial Russian Government 6% Perpetual Bond. 5,000 rubles. Issued by the State Commission for Debt Repayment. 1859.',
  }
);

// --- Row 382: Japan Industrial Bank Savings Bond, 55 Yen, 1960 ---
setDoc(382,
  'Japan Industrial Bank (日本勧業銀行): Savings Bond (貯蓄債券), 88th Issue, 55 Yen (No. 009951, 1960)',
  'This printed savings bond (貯蓄債券, chochiku saiken) of 55 yen was issued by the Japan Industrial Bank (Nippon Kangyo Ginko, 日本勧業銀行), 88th issue (第88回). No. 009951. Dated Showa 35 (1960). The Japan Industrial Bank (founded 1897) was a major government-linked agricultural and industrial development bank that issued savings bonds to the general public. It merged with the Dai-Ichi Bank in 1971 to form Dai-Ichi Kangyo Bank. These small-denomination savings bonds were a key mechanism of Japanese postwar household savings mobilization.',
  {
    type: 'Bond',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Japan Industrial Bank (日本勧業銀行, Nippon Kangyo Ginko)',
    issueDate: '1960-01-01',
    currency: 'JPY',
    language: 'Japanese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Japan Industrial Bank (日本勧業銀行) Savings Bond (貯蓄債券), 88th Issue. 55 yen. No. 009951. Showa 35 (1960).',
  }
);

// --- Row 383: Japanese WWII War Bond, 5 Yen, August 1942 ---
setDoc(383,
  'Japanese Wartime Patriotic Bond (戦時報国債券): 5 Yen, 4th Issue (No. 083986, August 1942)',
  'This printed wartime patriotic bond (戦時報国債券, Senji Hōkoku Saiken) of 5 yen is the 4th issue (第4回), No. 083986, dated August 1942 (Showa 17). Issued by the Japanese Imperial Government during the Pacific War (World War II), these small-denomination bonds were sold to the general public as part of the war finance effort. The name "patriotic bond" (hōkoku saiken) reflects the ideological framing of wartime borrowing as a civic duty. This 5-yen denomination was accessible to workers and small savers, and represents the grassroots dimension of Japanese wartime finance.',
  {
    type: 'Bond',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Japanese Imperial Government',
    issueDate: '1942-08-01',
    currency: 'JPY',
    language: 'Japanese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Japanese Wartime Patriotic Bond (戦時報国債券), 4th Issue (第4回). 5 yen. No. 083986. August 1942 (Showa 17).',
  }
);

// --- Row 384: Japanese WWII War Bond, 10 Yen, August 1942 ---
setDoc(384,
  'Japanese Wartime Patriotic Bond (戦時報国債券): 10 Yen, 4th Issue (No. 095184, August 1942)',
  'This printed wartime patriotic bond (戦時報国債券, Senji Hōkoku Saiken) of 10 yen is the 4th issue (第4回), No. 095184, dated August 1942 (Showa 17). Companion piece to No. 083986 (5-yen denomination, also 4th issue). Issued by the Japanese Imperial Government during World War II. The 10-yen denomination represents a slightly larger investment and together these bonds illustrate the tiered savings structure of Japanese wartime bond programs designed to reach all income levels of the population.',
  {
    type: 'Bond',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Japanese Imperial Government',
    issueDate: '1942-08-01',
    currency: 'JPY',
    language: 'Japanese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Japanese Wartime Patriotic Bond (戦時報国債券), 4th Issue (第4回). 10 yen. No. 095184. August 1942 (Showa 17).',
  }
);

// --- Row 385: Kahetian Railway Company, 4.5% Bond, 189 Rubles / £20, 1912 ---
setDoc(385,
  'Kahetian Railway Company: 4½% Bond, 189 Rubles / £20 Sterling (No. A14533, St. Petersburg, 1912)',
  'This printed bearer bond of 189 rubles (equivalent to £20 sterling) at 4½% annual interest was issued by the Kahetian Railway Company (Кахетинская железная дорога). No. A14533. St. Petersburg, 1912. The Kahetian (Kakhetian) Railway served the Kakheti wine-growing region of eastern Georgia in the Caucasus, linking Tiflis (Tbilisi) with the wine-producing areas. The dual denomination in rubles and British pounds indicates the bond was marketed to both domestic and international investors. This document reflects the active development of Caucasian railway infrastructure by private companies in the Russian Empire during the early twentieth century.',
  {
    type: 'Bond',
    subjectCountry: 'Georgia',
    issuingCountry: 'Russia',
    creator: 'Kahetian Railway Company',
    issueDate: '1912-01-01',
    currency: 'RUB',
    language: 'Russian, English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Kahetian Railway Company 4½% Bond. 189 rubles = £20 sterling. No. A14533. St. Petersburg, 1912.',
  }
);

// --- Row 386: Imperial Russian Government 4% State Loan 1902, 1,000 German Marks ---
setDoc(386,
  'Imperial Russian Government: 4% State Loan 1902, 1,000 German Marks (No. 116485)',
  'This printed bearer bond of 1,000 German marks at 4% annual interest forms part of the Imperial Russian Government 4% State Loan of 1902. No. 116485. Issued as an external loan denominated in German marks and targeted at German investors, this bond represents one of the largest and most active foreign government bond markets in pre-World War I Germany. Russian imperial bonds were widely held by German, French, and Dutch investors and formed a major component of European international capital flows before 1914. After the Bolshevik Revolution of 1917, the new Soviet government repudiated all Tsarist debt, rendering these bonds worthless.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1902-01-01',
    currency: 'DEM',
    language: 'Russian, German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Imperial Russian Government 4% State Loan 1902. 1,000 German marks. No. 116485.',
  }
);

// --- Row 387: Imperial Russian Government 4.5% State Loan 1905, 500 German Marks ---
setDoc(387,
  'Imperial Russian Government: 4½% State Loan 1905, 500 German Marks (No. 319459)',
  'This printed bearer bond of 500 German marks at 4½% annual interest forms part of the Imperial Russian Government 4½% State Loan of 1905. No. 319459. Like the 1902 loan (No. 116485), this bond was issued in German marks to attract German capital. The 1905 loan was issued in the midst of the Russo-Japanese War and the 1905 Revolution, making it a particularly significant instrument from a turbulent period in Russian financial history. The rate of 4½% (higher than the 1902 issue) reflected the increased credit risk perceived by investors during this period of political and military crisis.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1905-01-01',
    currency: 'DEM',
    language: 'Russian, German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Imperial Russian Government 4½% State Loan 1905. 500 German marks. No. 319459.',
  }
);

// --- Row 388: Dutch Kettingverklaring for De Woning Maatschappij Batavia, 1941 ---
setDoc(388,
  'Kettingverklaring (Securities Chain Declaration) for De Woning Maatschappij Batavia Shares (No. 1499, January 24, 1941)',
  'This printed Dutch legal form is a kettingverklaring (chain declaration), No. 1499, certifying a chain of ownership for shares in De Woning Maatschappij Batavia (Batavia Housing Company, Netherlands East Indies). Dated January 24, 1941. Under Dutch securities law of the period, shares in companies registered in the Netherlands East Indies (present-day Indonesia) that were traded on Dutch exchanges required a "chain declaration" linking each successive owner in an unbroken chain of declarations. This document records the holder\'s assertion of unbroken ownership, enabling valid transfer on the Dutch stock exchange. This document is likely related to the De Woning Maatschappij Batavia share certificate (cf. No. 355) elsewhere in the collection.',
  {
    type: 'Certificate',
    subjectCountry: 'Indonesia',
    issuingCountry: 'Netherlands',
    creator: 'De Woning Maatschappij Batavia (shareholder / Dutch notary)',
    issueDate: '1941-01-24',
    currency: '',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Dutch Kettingverklaring (chain declaration) No. 1499 for De Woning Maatschappij Batavia shares. January 24, 1941. Related to collection item No. 355.',
  }
);

// --- Row 389: Royal Dutch Petroleum Company Stock Purchase Warrant, 1937 ---
setDoc(389,
  'Koninklijke Nederlandsche Maatschappij (Royal Dutch Petroleum): Stock Purchase Warrant No. 015281 (The Hague, February 1937)',
  'This printed document is a stock purchase warrant (inschrijvingsrecht) No. 015281 issued by the Koninklijke Nederlandsche Maatschappij tot Exploitatie van Petroleumbronnen in Nederlandsch-Indië (Royal Dutch Petroleum Company, commonly known as Royal Dutch). The Hague, February 1937. The warrant entitles the bearer to subscribe for a specified number of new shares at F. 4,500 per share until March 31, 1940, or at F. 5,000 until March 31, 1943. Royal Dutch (founded 1890) was one of the forerunners of the Royal Dutch Shell oil company that merged with Shell Transport and Trading Company in 1907. This warrant, issued as part of a rights offering, reflects the company\'s capital-raising activities in the late 1930s.',
  {
    type: 'Certificate',
    subjectCountry: 'Indonesia',
    issuingCountry: 'Netherlands',
    creator: 'Koninklijke Nederlandsche Maatschappij tot Exploitatie van Petroleumbronnen in Nederlandsch-Indië (Royal Dutch Petroleum)',
    issueDate: '1937-02-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Royal Dutch Petroleum Company (Koninklijke Nederlandsche Maatschappij) stock purchase warrant. No. 015281. The Hague, February 1937. Subscription price F. 4,500 (until March 31, 1940) or F. 5,000 (until March 31, 1943).',
  }
);

// --- Row 390: Austrian Third War Loan, 10,000 Kronen, 1915 ---
setDoc(390,
  'Austrian Third War Loan (Dritte Österreichische Kriegsanleihe): 5½% Staatsanleihe, 10,000 Kronen (Serie E, No. 100356, Vienna, October 1, 1915)',
  'This printed bearer bond of 10,000 kronen at 5½% annual tax-free interest is one of the large-denomination notes of the Third Austrian War Loan (Dritte Fünfeinhalbrozentige Österreichische Kriegsanleihe). Serie E, No. 100356. Vienna, October 1, 1915. Austria-Hungary issued eight successive war loans between 1914 and 1918 to finance its participation in World War I. By the third issue in 1915 the war was already proving far more costly than anticipated, and subscriptions were promoted through aggressive public campaigns. The 10,000-krone denomination was among the largest available, targeting institutional and wealthy individual investors. The postwar hyperinflation and collapse of the Habsburg Empire rendered these bonds effectively worthless.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'Austro-Hungarian Government',
    issueDate: '1915-10-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Austrian Third War Loan (Dritte Österreichische Kriegsanleihe), 5½% amortisable. 10,000 kronen, Serie E, No. 100356. Vienna, October 1, 1915.',
  }
);

// --- Row 391: La Pelleterie Russo-Américaine, Paris, 1926 ---
setDoc(391,
  'La Pelleterie Russo-Américaine (Russian-American Fur Company): Action de 100 Francs (No. 036857, Paris, July 31, 1926)',
  'This printed bearer share (action au porteur) of 100 francs is No. 036857 of the Société Anonyme La Pelleterie Russo-Américaine (Russian-American Fur Company). Capital 5,000,000 francs. Paris, July 31, 1926. Organized under French law, this company traded in furs sourced from Russia and North America, operating at the intersection of the Soviet fur export industry (which used fur sales to earn hard currency in the 1920s) and the Western luxury goods market. The certificate dates from the New Economic Policy (NEP) era in the Soviet Union, when limited private and foreign commercial activity was briefly tolerated.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'France',
    creator: 'La Pelleterie Russo-Américaine',
    issueDate: '1926-07-31',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'La Pelleterie Russo-Américaine, S.A. Action au porteur, 100 francs. Capital 5,000,000 francs. No. 036857. Paris, July 31, 1926.',
  }
);

// --- Row 392: La Platense Flotilla Company Limited, share warrant, £10 ---
setDoc(392,
  'La Platense Flotilla Company Limited: Share Warrant to Bearer, £10 (No. 15821)',
  'This printed share warrant to bearer entitles the holder to one fully paid share of £10 in La Platense Flotilla Company Limited. No. 15821. The company operated a river flotilla service on the Río de la Plata and its tributaries in Argentina and Uruguay. Capital £1,000,000 (100,000 shares of £10 each). Registered in the United Kingdom. The Río de la Plata river system was a critical commercial waterway linking the agricultural hinterlands of Argentina, Uruguay, Paraguay, and southern Brazil to the port of Buenos Aires, and several British-registered companies competed to provide steamship and flotilla services on the river during the late nineteenth and early twentieth centuries.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Argentina',
    issuingCountry: 'United Kingdom',
    creator: 'La Platense Flotilla Company Limited',
    issueDate: '1905-01-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'La Platense Flotilla Company Limited. Share warrant to bearer, £10. No. 15821. One of 100,000 shares. River flotilla, Río de la Plata. Ca. early 1900s.',
  }
);

// --- Row 393: Ancient Greek/Coptic Papyrus ---
setDoc(393,
  'Ancient Papyrus: Greek or Coptic Cursive Financial or Legal Document (Ptolemaic–Byzantine Egypt, ca. 1st–7th century CE)',
  'This fragment of an ancient papyrus bears cursive script in Greek or Coptic, consistent with documents produced in Egypt during the Ptolemaic, Roman, or Byzantine period (ca. 3rd century BCE to 7th century CE). The text appears to be a financial or administrative document—possibly a tax receipt, grain account, loan agreement, or land lease—of the type commonly produced in pharaonic and Hellenistic administrative practice. Papyri of this kind constitute the oldest surviving form of financial documentation in the Western world and are among the earliest physical evidence of organized credit, taxation, and contractual obligation. This is likely the oldest item in the collection.',
  {
    type: 'Certificate',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Unknown (Egyptian administrative scribal office)',
    issueDate: '0200-01-01',
    currency: '',
    language: 'Greek',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Ancient papyrus fragment with cursive Greek or Coptic script. Likely a financial or administrative document (tax receipt, grain account, or land lease). Ptolemaic, Roman, or Byzantine Egypt, ca. 1st–7th century CE. Oldest item in the collection.',
  }
);

// --- Row 394: Little Miami Rail Road Company, 10 shares, 1862 ---
setDoc(394,
  'Little Miami Rail Road Company: 10 Shares at $50 Each (No. 4644, Cincinnati, June 17, 1862)',
  'This engraved stock certificate represents 10 shares of $50 each in the Capital Stock of the Little Miami Rail Road Company, State of Ohio. Issued in the name of James G. King\'s Sons, Trustees. No. 4644. Cincinnati, June 17, 1862. The Little Miami Rail Road (incorporated 1836) was one of Ohio\'s earliest railroads, running from Cincinnati to Springfield, Ohio. By 1862 it had been leased to the Pittsburgh, Cincinnati and St. Louis Railway (later part of the Pennsylvania Railroad system). The James G. King\'s Sons firm was a prominent New York banking house that frequently acted as trustee for railroad securities.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Little Miami Rail Road Company',
    issueDate: '1862-06-17',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Little Miami Rail Road Company, Ohio. 10 shares at $50 each. No. 4644. Issued to James G. King\'s Sons, Trustees. Cincinnati, June 17, 1862.',
  }
);

// --- Row 395: Republic of China, 8% Lung-Tsing-U-Hai Railway Treasury Bill, F. 1,000, 1923 ---
setDoc(395,
  'Republic of China: 8% Lung-Tsing-U-Hai Railway Treasury Bill, F. 1,000 Dutch Guilders (No. 11638, 1923)',
  'This printed bearer bond of F. 1,000 Dutch guilders at 8% annual interest is part of the 8% Lung-Tsing-U-Hai Railway (隴秦豫海鐵路) Treasury Bill Loan of 1923, issued by the Government of the Republic of China. No. 11638. Total issue F. 16,667,000 guilders. The Lung-Tsing-U-Hai Railway (Gansu–Shaanxi–Henan–Jiangsu line) was a planned trans-China railway linking Gansu province in the northwest to the eastern coast; financing was secured through multiple bond issues in different European markets—this Dutch guilder issue targeted Netherlands investors, while companion issues were denominated in Belgian francs (cf. row 430) and British pounds. The railway was only partially completed due to political instability and the ongoing Chinese civil conflicts.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Republic of China',
    issueDate: '1923-01-01',
    currency: 'NLG',
    language: 'Dutch, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: '8% Lung-Tsing-U-Hai Railway Treasury Bill Loan 1923. F. 1,000 Dutch guilders. No. 11638. Total issue F. 16,667,000. Issued by Republic of China.',
  }
);

// --- Row 396: Merchants Insurance Company, Providence, RI, 1858 ---
setDoc(396,
  'Merchants Insurance Company: Share Transfer Certificate, 12 Shares to James S. Ham (No. 178, Providence, April 1, 1858)',
  'This printed share transfer certificate records the transfer of 12 shares in the Merchants Insurance Company to James S. Ham. Certificate No. 178. Providence, Rhode Island, April 1, 1858. The Merchants Insurance Company was one of several Providence-based marine and fire insurance companies active in the mid-nineteenth century, serving the prosperous commercial and manufacturing community of Rhode Island. This document reflects the active secondary market for insurance company shares in antebellum New England.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Merchants Insurance Company',
    issueDate: '1858-04-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Merchants Insurance Company (Providence, Rhode Island). Transfer of 12 shares to James S. Ham. No. 178. April 1, 1858.',
  }
);

// --- Row 397: Mexican Telephone Company, 25 shares, 1905 (Cancelled) ---
setDoc(397,
  'Mexican Telephone Company: 25 Shares at $10 Each (No. 16621, New York, 1905) [Cancelled]',
  'This engraved stock certificate represents 25 shares of $10 each in the Mexican Telephone Company. Certificate No. 16621. Issued in New York to Harry E. Jacobs; countersigned by the Boston Safe Deposit & Trust Co. on February 28, 1905. Stamped CANCELLED. The Mexican Telephone Company was one of several foreign-controlled telecommunications ventures operating in Mexico during the Porfiriato era (the long presidency of Porfirio Díaz, 1876–1911), when extensive foreign capital was invited into Mexico to develop its infrastructure. The company was later absorbed into the Ericsson telephone network that became dominant in Mexico.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Mexico',
    issuingCountry: 'United States',
    creator: 'Mexican Telephone Company',
    issueDate: '1905-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Mexican Telephone Company. 25 shares at $10 each. No. 16621. Issued to Harry E. Jacobs. Countersigned by Boston Safe Deposit & Trust Co., February 28, 1905. CANCELLED.',
  }
);

// --- Row 398: Minas Geraes Goldfields Limited, 25 shares, London, 1912 ---
setDoc(398,
  'Minas Geraes Goldfields Limited: Share Warrant to Bearer, 25 Shares at £1 Each (No. 023770, London, July 1912)',
  'This printed share warrant to bearer entitles the holder to 25 fully paid shares of £1 each in Minas Geraes Goldfields Limited. Capital £150,000. No. 023,770. London, July 1912. The company was incorporated in the United Kingdom to exploit gold mining properties in the state of Minas Gerais (often spelled "Minas Geraes" in period documents), Brazil, one of the world\'s richest mineral regions, historically famous for its eighteenth-century gold rush and continuing gold, iron ore, and diamond production into the twentieth century. British-registered mining ventures in Brazil were common during the Edwardian era.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Brazil',
    issuingCountry: 'United Kingdom',
    creator: 'Minas Geraes Goldfields Limited',
    issueDate: '1912-07-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Minas Geraes Goldfields Limited. Share warrant to bearer, 25 shares at £1 each. Capital £150,000. No. 023,770. London, July 1912.',
  }
);

// --- Row 399: Mines d'Or de Nam Kok, Paris ---
setDoc(399,
  'Mines d\'Or de Nam Kok: Action de 100 Francs au Porteur (No. 059820, Paris)',
  'This printed bearer share (action au porteur) of 100 francs is No. 059,820 of the Mines d\'Or de Nam Kok (Nam Kok Gold Mines), Société Anonyme. Capital 30,000,000 francs. Paris. The company\'s name Nam Kok (南谷 or similar) suggests operations in the mountainous regions of French Indochina or southern China, where French colonial enterprises actively exploited mineral resources. Gold mining in these regions was an important sector of French colonial investment during the late nineteenth and early twentieth centuries.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Vietnam',
    issuingCountry: 'France',
    creator: 'Mines d\'Or de Nam Kok',
    issueDate: '1910-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Mines d\'Or de Nam Kok. Action au porteur, 100 francs. Capital 30,000,000 francs. No. 059,820. Paris. Ca. early 1900s. Gold mining in Indochina or southern China.',
  }
);

// --- Row 400: Moscow-Kazan Railway Company 4% Bond, 2,000 German Marks, 1901 ---
setDoc(400,
  'Moscow-Kazan Railway Company: 4% Bond Loan of 1901, 2,000 German Marks (No. 06181)',
  'This trilingual (Russian/German/Dutch) printed bearer bond of 2,000 German marks at 4% annual interest is part of the Moscow-Kazan Railway Company Loan of 1901, issued under guarantee of the Imperial Russian Government. No. 06181. The Moscow-Kazan Railway (Московско-Казанская железная дорога) was one of Russia\'s major trunk lines, running from Moscow southeast to Kazan and the Volga region. The trilingual format—Russian, German, and Dutch—reflects the broad European investor base targeted by this bond issue, which was simultaneously marketed in Germany and the Netherlands. The Imperial Government guarantee made these railway bonds attractive to risk-averse European investors.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Moscow-Kazan Railway Company',
    issueDate: '1901-01-01',
    currency: 'DEM',
    language: 'Russian, German, Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Moscow-Kazan Railway Company 4% Bond Loan of 1901. 2,000 German marks. No. 06181. Trilingual Russian/German/Dutch. Imperial Russian Government guarantee.',
  }
);

// --- Row 401: Morris Canal and Banking Company transfer receipts, 1845 ---
setDoc(401,
  'Morris Canal and Banking Company: Two Stock Transfer Receipts (June 28 and July 10, 1845, Phenix Bank, New York)',
  'This document contains two stock transfer receipts for shares of the Morris Canal and Banking Company of 1844, issued at the Agency at Phenix Bank, New York City. The first receipt (June 28, 1845) records a transfer to Hopkins & Berton; the second receipt (July 10, 1845) records a transfer to E.S. Map. Both documents appear on a single sheet. The Morris Canal and Banking Company operated the Morris Canal in New Jersey, a combined canal and banking enterprise chartered in 1824. By the 1840s the company was undergoing restructuring, and these transfers reflect active secondary market trading in its shares through New York financial intermediaries. The Phenix Bank was a prominent New York commercial bank that served as a transfer agent for numerous securities.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Morris Canal and Banking Company (Agency at Phenix Bank, NYC)',
    issueDate: '1845-06-28',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Morris Canal and Banking Company. Two stock transfer receipts: (1) June 28, 1845, to Hopkins & Berton; (2) July 10, 1845, to E.S. Map. Phenix Bank agency, New York City.',
  }
);

// --- Row 402: New York Central and Hudson River Railroad Debt Certificate, $1,000, 1892 ---
setDoc(402,
  'New York Central and Hudson River Railroad Company: Debt Certificate No. 1390, $1,000 at 4% (December 24, 1892)',
  'This engraved certificate is a $1,000 debt certificate at 4% annual interest, due May 1, 1902, issued by the New York Central and Hudson River Railroad Company. No. 1390. Dated December 24, 1892. The New York Central and Hudson River Railroad, formed in 1869 through the consolidation of the New York Central and Hudson River Railroads by Cornelius Vanderbilt, was one of the most important trunk railroads in the United States, forming the core of what became the New York Central System. Debt certificates of this type were a form of short- to medium-term railroad financing used alongside bonded debt and equity.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'New York Central and Hudson River Railroad Company',
    issueDate: '1892-12-24',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'New York Central and Hudson River Railroad Company. Debt Certificate No. 1390. $1,000 at 4%. Due May 1, 1902. December 24, 1892.',
  }
);

// --- Row 403: New York Central Rail Road Company, $1,000 bond, 1858 (Cancelled) ---
setDoc(403,
  'New York Central Rail Road Company: $1,000 Bond at 6% (No. 1890, August 1858) [Cancelled]',
  'This engraved bearer bond of $1,000 at 6% annual interest was issued by the New York Central Rail Road Company. No. 1890. Dated August 1858. Stamped CANCELLED. The New York Central Rail Road was created in 1853 by the merger of several smaller railroads connecting Albany to Buffalo and was one of the flagship lines of the Cornelius Vanderbilt railroad empire. At 6% interest, this bond was issued during a period when American railroads were paying relatively high rates to attract capital. The CANCELLED stamp indicates this bond was redeemed or retired before the company reorganized into the New York Central and Hudson River Railroad in 1869.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'New York Central Rail Road Company',
    issueDate: '1858-08-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'New York Central Rail Road Company. $1,000 bond, 6%. No. 1890. August 1858. CANCELLED.',
  }
);

// --- Row 404: NY New Haven Hartford Railroad Gold Bond, $10,000, 1912 (Surrendered) ---
setDoc(404,
  'New York, New Haven and Hartford Railroad: Harlem River–Port Chester First Mortgage 4% 50-Year Gold Bond, $10,000 (No. 331, 1912) [Surrendered]',
  'This large-denomination printed bond of $10,000 at 4% annual interest is a Harlem River–Port Chester Branch First Mortgage Gold Bond, with a 50-year term, issued by the New York, New Haven and Hartford Railroad Company. No. 331. Dated 1912. Stamped SURRENDERED. The New Haven Railroad was a major northeastern rail carrier. The Harlem River–Port Chester Branch (now part of the Metro-North Harlem and New Haven Lines) was a key commuter corridor linking Manhattan to Westchester County and southern Connecticut. First mortgage gold bonds—backed by a pledge of railway assets and denominated in gold—were among the most creditworthy railway securities available to early twentieth-century investors.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'New York, New Haven and Hartford Railroad Company',
    issueDate: '1912-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'New York, New Haven and Hartford Railroad. Harlem River–Port Chester First Mortgage 4% 50-Year Gold Bond. $10,000. No. 331. 1912. SURRENDERED.',
  }
);

// --- Row 405: Nuovo Monte Sussidio Vacabile della Città di Firenze, 1693 ---
setDoc(405,
  'Nuovo Monte Sussidio Vacabile della Città di Firenze: Public Debt Certificate (Florence, 1693)',
  'This handwritten document is a public debt certificate of the Nuovo Monte Sussidio Vacabile della Città di Firenze (New Vacatable Subsidy Monte of the City of Florence), issued under Grand Duke Cosimo III de\' Medici at 6% annual interest. Florence, 1693. The Florentine monti (plural of monte, meaning "fund") were among the earliest instruments of transferable public debt in Europe, originating in the thirteenth and fourteenth centuries as forced loans that were later made voluntary and negotiable. The "Monte Sussidio Vacabile" was a specific fund type in which the income right "vacated" (reverted) to the Monte upon the original holder\'s death, giving it features of both a bond and an annuity. This 1693 document is one of the oldest financial instruments in the collection and represents a direct link to the origins of modern government debt finance in Renaissance Italy.',
  {
    type: 'Bond',
    subjectCountry: 'Italy',
    issuingCountry: 'Italy',
    creator: 'Città di Firenze (City of Florence, Grand Duchy of Tuscany)',
    issueDate: '1693-01-01',
    currency: 'ITL',
    language: 'Italian',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Nuovo Monte Sussidio Vacabile della Città di Firenze. Public debt certificate at 6%. Florence, 1693. Grand Duke Cosimo III de\' Medici. Among the earliest transferable public debt instruments in Europe.',
  }
);

// --- Row 406: Compañía del Ferro-Carril de Palencia a Ponferrada, bilingual bond ---
setDoc(406,
  'Compañía del Ferro-Carril de Palencia a Ponferrada / Chemins de Fer du Nord-Ouest de l\'Espagne: Bilingual Bond No. 32,105',
  'This bilingual (Spanish/French) printed bearer bond No. 32,105 was issued by the Compañía del Ferro-Carril de Palencia a Ponferrada (Compagnie des Chemins de Fer du Nord-Ouest de l\'Espagne / Railway Company of Palencia to Ponferrada). Total emission: 68,420 obligations at 1,900 Reales Vellón = 500 Francs each. Issued in Madrid. The Palencia–Ponferrada railway crossed the Cantabrian Mountains in northern Spain, connecting the Castilian grain plains to the coal-mining Bierzo region of León. The bilingual format (Spanish and French) indicates the bonds were marketed to both domestic Spanish and French investors, who were major providers of railway finance in mid-nineteenth-century Spain.',
  {
    type: 'Bond',
    subjectCountry: 'Spain',
    issuingCountry: 'Spain',
    creator: 'Compañía del Ferro-Carril de Palencia a Ponferrada',
    issueDate: '1865-01-01',
    currency: 'ESP',
    language: 'Spanish, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Compañía del Ferro-Carril de Palencia a Ponferrada (Nord-Ouest de l\'Espagne). Bilingual Spanish/French bearer bond No. 32,105. 1,900 Reales = 500 Francs. 68,420 obligations in total. Madrid, ca. 1860s.',
  }
);

// --- Row 407: Norwich & Worcester Rail Road Company, $1,000 bond ---
setDoc(407,
  'Norwich & Worcester Rail Road Company: $1,000 Bond (No. 468, Hartford)',
  'This printed bearer bond of $1,000 was issued by the Norwich & Worcester Rail Road Company, incorporated in the State of Connecticut. No. 468. Printed by Evans, Hartford. Total authorized issue $500,000. Interest payable semi-annually. The Norwich & Worcester Railroad (incorporated 1832) ran from Norwich on the Thames River in Connecticut northwest to Worcester, Massachusetts, providing an important link in the coastal rail network of southern New England. This bond, issued in Hartford, Connecticut, was part of the company\'s mid-nineteenth-century capital program.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Norwich & Worcester Rail Road Company',
    issueDate: '1845-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Norwich & Worcester Rail Road Company (State of Connecticut). $1,000 bearer bond. No. 468. Total authorized issue $500,000. Printed by Evans, Hartford. Ca. 1840s–1850s.',
  }
);

// --- Row 408: Spanish Colonial Libramento (Payment Order), Lima, 1816 ---
setDoc(408,
  'Spanish Colonial Government: Libramento (Payment Order) No. 444 – Lima, Peru, October 1816 [Empréstito Patriótico]',
  'This handwritten document is a libramento (government payment order/draft) No. 444 issued by the Spanish colonial treasury in Lima, Peru, in October 1816. Related to the "Empréstito Patriótico" (Patriotic Loan), an emergency forced loan imposed by the colonial government on the population of Lima to finance the royalist war effort during the South American Wars of Independence. Son 500 pesos. The document specifies a remarkably high annual interest rate of 25%. An additional notation from 1817 attests to further processing or partial payment. At the time of issue, the Vice-royalty of Peru was one of the last remaining strongholds of Spanish colonial rule in South America; independence would come only in 1821. This document is a rare surviving financial instrument from the final phase of Spanish colonial governance in Peru.',
  {
    type: 'Draft',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Spanish Colonial Treasury, Lima (Tesorería de Lima)',
    issueDate: '1816-10-01',
    currency: 'PEN',
    language: 'Spanish',
    numberPages: 1,
    period: '19th Century',
    notes: 'Spanish colonial libramento No. 444. Lima, Peru, October 1816. Empréstito Patriótico (forced loan). 500 pesos at 25% annual interest. Additional notation from 1817.',
  }
);

// --- Row 409: Russian Internal 5% Prize Loan, 100 Rubles, 1864 ---
setDoc(409,
  'Imperial Russian Government: Internal 5% Prize Loan (Билет Внутреннего 5% с Выигрышами Займа), 100 Rubles (Series 15871, No. 45, 1864)',
  'This bilingual (Russian/German) printed bearer lottery bond of 100 rubles at 5% interest with prize draws (выигрыши) is from the Internal 5% Prize Loan of 1864. Russian title: Билетъ Внутреннего 5% с Выигрышами Займа. German title: "Obligation der Russischen 5% Inneren Anleihe mit Prämien-Verlösungen von 1864." Series 15871, No. 45. 1864. Prize loans (займы с выигрышами) combined regular interest payments with periodic lottery drawings offering large cash prizes (выигрыши), making them more attractive to small and medium investors than plain bonds. They were widely used by European governments in the nineteenth century. The bilingual German text indicates this domestic loan was also actively traded in German financial markets.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1864-01-01',
    currency: 'RUB',
    language: 'Russian, German',
    numberPages: 1,
    period: '19th Century',
    notes: 'Imperial Russian Government Internal 5% Prize Loan 1864. 100 rubles. Series 15871, No. 45. Bilingual Russian/German. Lottery prize-bond format.',
  }
);

// --- Row 410: Omnium Français du Film, Paris, 1928 ---
setDoc(410,
  'Omnium Français du Film: Action Série A de 100 Francs (No. 022884, Paris, May 10, 1928)',
  'This printed bearer share (action au porteur) Série A of 100 francs is No. 022884 of the Omnium Français du Film (French Film Omnibus Company), Société Anonyme. Capital 3,000,000 francs divided into 30,000 shares (26,000 Série A at 100 francs each + 4,000 Série B). Paris, May 10, 1928. The company was incorporated at the height of the French silent film industry\'s prosperity, just before the introduction of sound films (the "talkies") began in 1927–28. The Omnium Français du Film was engaged in film production, distribution, or related activities in the French film industry of the 1920s.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Omnium Français du Film',
    issueDate: '1928-05-10',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Omnium Français du Film, S.A. Action au porteur Série A, 100 francs. Capital 3,000,000 francs (26,000 Série A + 4,000 Série B). No. 022884. Paris, May 10, 1928.',
  }
);

// --- Row 411: Generale Keyserlycke Indische Compagnie (Ostend Company) dividend receipt, 1734 ---
setDoc(411,
  'Generale Keyserlycke Indische Compagnie (Ostend Company): Dividend Receipt, 360 Guilders (Antwerp, August 11, 1734)',
  'This handwritten document is a dividend receipt from the Generale Keyserlycke Indische Compagnie (General Imperial East India Company), also known as the Ostend Company, signed by Joan Baptista Cogels in his capacity as Cassier (Treasurer) of the Company. Folio 247:B. Dated August 11, 1734, in Antwerp. The receipt confirms payment of 360 guilders (gls. 360 = w.t.) to C.J. De Witte(?), representing 60 guilders per share at a repartition rate of 6 per cent on 6 shares held in the Company. A handwritten annotation "P. Orange 1734" appears at the top right. The Ostend Company was chartered by the Holy Roman Emperor Charles VI in 1722 to trade with the East and West Indies, but was suppressed under pressure from Britain, the Dutch Republic, and France in 1731. The 1734 date indicates this receipt was issued during the continued winding-down of the Company\'s affairs and distribution of its remaining assets to shareholders—more than three years after suspension of trading operations.',
  {
    type: 'Receipt',
    subjectCountry: 'Belgium',
    issuingCountry: 'Belgium',
    creator: 'Generale Keyserlycke Indische Compagnie (Ostend Company)',
    issueDate: '1734-08-11',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Generale Keyserlycke Indische Compagnie (Ostend Company). Dividend receipt, 360 guilders (60 per share × 6 shares). Signed by Joan Baptista Cogels, Cassier. Folio 247:B. Antwerp, August 11, 1734.',
  }
);

// --- Row 412: Orde der Paters Norbertynen, 8% Mortgage Bond, F. 1,000, 1928 ---
setDoc(412,
  'Orde der Paters Norbertynen (Jassauer Praemonstratenser Koorheeren-Orde): 8% First Mortgage Bond, F. 1,000 (No. 0.274, Budapest, August 1, 1928)',
  'This printed bearer bond of F. 1,000 Dutch guilders at 8% annual interest is the No. 0.274 of the First Mortgage Loan (1e Hypothecaire Leening) issued by the Jassauer Praemonstratenser Koorheeren-Orde (Norbertine / Premonstratensian Fathers of Jászó) at Gödöllő, near Budapest. Guaranteed by the Royal Hungarian Government. Total issue F. 1,600,000 (1,200 bonds of F. 1,000 + 800 bonds of F. 500). Budapest, August 1, 1928. Administered by the Hollandsche Garantie & Trust Compagnie (Dutch Guarantee and Trust Company) in Amsterdam. This unusual instrument was issued by a Catholic religious order—the Premonstratensians of Jászó Abbey in Hungary—secured on its agricultural and monastic properties, with a Dutch guilder denomination to attract Dutch investors and backed by a Hungarian government guarantee.',
  {
    type: 'Bond',
    subjectCountry: 'Hungary',
    issuingCountry: 'Hungary',
    creator: 'Orde der Paters Norbertynen (Jassauer Praemonstratenser Koorheeren-Orde)',
    issueDate: '1928-08-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Orde der Paters Norbertynen (Jászó Premonstratensian / Norbertine Fathers) 8% First Mortgage Bond. F. 1,000 Dutch guilders. No. 0.274. Budapest, August 1, 1928. Total issue F. 1,600,000. Royal Hungarian Government guarantee. Administered by Hollandsche Garantie & Trust Compagnie.',
  }
);

// --- Row 413: Oregon and Transcontinental Company, 100 shares, 1888 ---
setDoc(413,
  'Oregon and Transcontinental Company: 100 Shares at $100 Each (No. 10583, September 1888)',
  'This engraved stock certificate represents 100 shares of $100 each in the Capital Stock of the Oregon and Transcontinental Company. Capital $50,000,000. No. 10583. Dated September 1888. Signed by Edward Ede (Assistant Secretary) and Mayor Fiske (President). The Oregon and Transcontinental Company was a holding company organized by Henry Villard in 1881 to control and develop railway, steamship, and other transportation properties in the Pacific Northwest, including the Oregon Railway and Navigation Company and the Northern Pacific Railroad. Villard\'s empire collapsed in 1883, but the company continued under new management; this certificate dates from the reorganization period.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Oregon and Transcontinental Company',
    issueDate: '1888-09-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Oregon and Transcontinental Company. 100 shares at $100 each. Capital $50,000,000. No. 10583. September 1888. President: Mayor Fiske. Secretary: Edward Ede.',
  }
);

// --- Row 414: Österreichisches Bau-Los (Austrian Building Lottery), 10 Schilling, 1926 ---
setDoc(414,
  'Österreichisches Bau-Los (Austrian Building Lottery): 10 Schilling, Emission 1926 (Serie 2,485, No. 049, Vienna)',
  'This printed lottery bond (Los) of 10 Schilling is the Emission 1926 of the Österreichisches Bau-Los (Austrian Building Lottery). Serie 2,485, No. 049. Vienna, August 1926. Issued by the Federal Housing and Settlement Fund (Bundeswohn- und Siedlungsfonds), backed by the Austrian federal government (Bund) and guaranteed by state law for housing purposes (Wohnungszwecke). Lottery bonds (Lose) that combined fixed interest payments with periodic prize drawings were a popular savings instrument in Central Europe throughout the nineteenth and early twentieth centuries. This Bau-Los was specifically designed to fund the construction of affordable housing in Austria during the severe housing crisis of the postwar period.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'Bundeswohn- und Siedlungsfonds (Austrian Federal Housing and Settlement Fund)',
    issueDate: '1926-08-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Österreichisches Bau-Los, Emission 1926. 10 Schilling. Serie 2,485, No. 049. Vienna, August 1926. Bundeswohn- und Siedlungsfonds (Federal Housing and Settlement Fund). Lottery bond for housing purposes.',
  }
);

// --- Row 415: Compagnie Universelle du Canal Interocéanique de Panama, 500 Francs, 1888 ---
setDoc(415,
  'Compagnie Universelle du Canal Interocéanique de Panama: Titre Provisoire, 500 Francs (No. 0,269,304, Paris, June 20, 1888) [Redeemed via Receivership]',
  'This printed provisional bearer bond (titre provisoire au porteur négociable) for one fully paid 500-franc obligation is No. 0,269,304 of the Compagnie Universelle du Canal Interocéanique de Panama (Universal Interoceanic Canal Company of Panama), Société civile with limited liability. Part of an 720-million franc loan authorized by the laws of May 21, 1886, and June 8, 1888. Paris, June 20, 1888. Signed by H. Cottu and J.F. de Lesseps (Ferdinand de Lesseps\'s son). Stamped multiple times "REMBOURSE PAR LE SEQUESTRE" (redeemed by the receivership administrator). The Panama Canal Company had been launched in 1881 under Ferdinand de Lesseps, builder of the Suez Canal, with great fanfare and public investment. It collapsed in 1889 in one of the nineteenth century\'s greatest financial disasters, with hundreds of thousands of small French investors losing their savings—a scandal that implicated French politicians and led to a major political crisis. This bond represents one of the company\'s last capital-raising efforts before bankruptcy.',
  {
    type: 'Bond',
    subjectCountry: 'Panama',
    issuingCountry: 'France',
    creator: 'Compagnie Universelle du Canal Interocéanique de Panama',
    issueDate: '1888-06-20',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Compagnie Universelle du Canal Interocéanique de Panama. Titre provisoire 500 francs. No. 0,269,304. Paris, June 20, 1888. Signed by H. Cottu and J.F. de Lesseps. Stamped REMBOURSE PAR LE SEQUESTRE. Panama Canal Company scandal.',
  }
);

// --- Row 416: Café de la Paix (Béziers), Part de Fondateur, 1921 ---
setDoc(416,
  'Café de la Paix (Béziers): Part de Fondateur No. 1335 (Béziers, January 1, 1921)',
  'This elaborately printed founder\'s share (part de fondateur) No. 1335 belongs to the Café de la Paix, Société Anonyme, located at 1, Allées Paul Riquet, Béziers (in the Hérault department of southern France). Capital 1,200,000 francs divided into 2,400 shares of 500 francs each. Signed by two administrators. Béziers, January 1, 1921. The certificate features a striking Art Nouveau illustration of the Grand Café de la Paix building—a grand Haussmann-style café—with an allegorical draped female figure in a monumental pose. "Parts de fondateur" (founder\'s parts) were non-par-value shares issued to the founders of a company in recognition of their promotional work and carried entitlements to a share of future profits above a minimum dividend threshold.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Café de la Paix (Béziers)',
    issueDate: '1921-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Café de la Paix, Béziers. Part de fondateur No. 1335. Capital 1,200,000 francs / 2,400 shares of 500 francs. Béziers, January 1, 1921. Art Nouveau design. 1, Allées Paul Riquet.',
  }
);

// --- Row 417: Peruvian Republic Congressional Loan, Lima, 1823 ---
setDoc(417,
  'Peruvian Republic: Congressional War Loan Certificate (Lima, 1823)',
  'This handwritten certificate, bearing the official seal of the newly independent Republic of Peru, records participation in a 100,000-peso emergency loan (empréstito) authorized by the Sovereign Congress (Soberano Congreso) of Peru in 1823. The document states that the treasury will pay the bearer the stated amount "más su respectivo interés al cinco por ciento mensual" (plus its respective interest at five per cent per month) as part of the contribution to the congressional loan of 100,000 pesos "determinado por el Soberano Congreso para subvenir a las necesidades públicas, y dar al ejército un impulso capas de terminar la guerra" (determined by the Sovereign Congress to meet public needs and give the army an impetus capable of ending the war). Principal and interest are declared irremissible national debt with priority over all other obligations. Lima, 1823. The 5% monthly interest rate reflects the extreme fiscal stress of the independence struggle; Peru had declared independence in 1821 but fighting continued until 1824–26.',
  {
    type: 'Bond',
    subjectCountry: 'Peru',
    issuingCountry: 'Peru',
    creator: 'Peruvian Republic (Soberano Congreso)',
    issueDate: '1823-01-01',
    currency: 'PEN',
    language: 'Spanish',
    numberPages: 1,
    period: '19th Century',
    notes: 'Peruvian Republic Congressional War Loan certificate. Lima, 1823. 100,000-peso loan authorized by the Soberano Congreso. 5% monthly interest. Irremissible national debt with priority claim. Independence war finance.',
  }
);

// --- Row 418: Philadelphia & West Chester Turnpike Road Co., unissued stock certificate ---
setDoc(418,
  'Philadelphia & West Chester Turnpike Road Company: Unissued/Blank Stock Certificate, $25 Per Share',
  'This beautifully engraved stock certificate template for the Philadelphia & West Chester Turnpike Road Company is blank and unissued—no share numbers, holder name, or date have been filled in. Par value: $25 per share. Printed by Evans, Printer, Fourth & Library Streets, Philadelphia. The certificate features a prominent engraved vignette of a winged allegorical figure (Fame or Commerce) above a globe, flanked by a covered Conestoga wagon with horses. The Philadelphia & West Chester Turnpike Road was one of Pennsylvania\'s oldest and most important toll roads, chartered in 1803 and linking central Philadelphia to the borough of West Chester in Chester County. The surviving blank certificate, ca. 1850s–1860s, is a specimen of antebellum turnpike company stock printing.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Philadelphia & West Chester Turnpike Road Company',
    issueDate: '1855-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Philadelphia & West Chester Turnpike Road Company. Unissued/blank stock certificate. $25 per share. Printed by Evans, Printer, Philadelphia. Ca. 1850s–1860s.',
  }
);

// --- Row 419: Société des Plantations d'Agaves de l'Annam, Saïgon, 1925 ---
setDoc(419,
  'Société des Plantations d\'Agaves de l\'Annam: Action de 300 Francs au Porteur (No. 05863, Saïgon, 1925)',
  'This printed bearer share (action de 300 francs au porteur, entièrement libérée) No. 05863 is a share in the Société des Plantations d\'Agaves de l\'Annam (Agave Plantation Society of Annam), a French colonial agricultural company. Capital 9,750,000 francs divided into 3,500 shares of 300 francs each. Siège social: Saïgon. Statutes deposited before a notary in Saïgon (Cochinchine), March 27, 1924. 1925. Agave was cultivated in the Annam region of central French Indochina (present-day central Vietnam) primarily for its sisal fiber, used in rope, twine, and textile manufacture. French colonial companies developed agave plantations as part of the broader agricultural exploitation of Indochina during the interwar period.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Vietnam',
    issuingCountry: 'France',
    creator: 'Société des Plantations d\'Agaves de l\'Annam',
    issueDate: '1925-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société des Plantations d\'Agaves de l\'Annam. Action au porteur, 300 francs, entièrement libérée. Capital 9,750,000 francs / 3,500 shares. No. 05863. Saïgon, 1925.',
  }
);

// --- Row 420: Polsko Amerykańskie Towarzystwo Naftowe "Columbia," 50 shares, 1922 ---
setDoc(420,
  'Polsko Amerykańskie Towarzystwo Naftowe "Columbia" S.A.: 50 Shares at 1,000 Polish Marks Each (No. 291101–291150, 1922)',
  'This printed block certificate represents 50 shares at 1,000 Polish marks each (total 50,000 marks) in the Polsko Amerykańskie Towarzystwo Naftowe "Columbia" S.A. (Polish-American Petroleum Company "Columbia"), Warsaw. Shares No. 291101 to No. 291150. Company statute (Statut spółki) approved February 6, 1922. Signed by the President of the Council (Prezes Rady), Management Board (Zarząd), Accountant (Księgowy), and Treasurer (Skarbnik). The certificate features a photographic illustration of oil derricks at the "Columbia" oil field in Galicia (southeastern Poland), one of Europe\'s oldest oil-producing regions. The company was founded in 1922 during a period of hyperinflation that rapidly eroded the value of the Polish mark.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Poland',
    issuingCountry: 'Poland',
    creator: 'Polsko Amerykańskie Towarzystwo Naftowe "Columbia" S.A.',
    issueDate: '1922-01-01',
    currency: 'PLN',
    language: 'Polish',
    numberPages: 1,
    period: '20th Century',
    notes: 'Polsko Amerykańskie Towarzystwo Naftowe "Columbia" S.A., Warsaw. 50 shares at 1,000 Polish marks each. Shares No. 291101–291150. Statute approved February 6, 1922. Oil fields in Galicia.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 376–420 (45 documents, batch13).');
