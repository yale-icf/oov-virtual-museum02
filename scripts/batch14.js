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

// --- Row 421: Compagnie Privilégiée des Ports, Débarcadère Maritime et Terrains de Cadix, 400 Francs ---
setDoc(421,
  'Compagnie Privilégiée des Ports, Débarcadère Maritime et Terrains de Cadix: Obligation Hypothécaire, 400 Francs (No. 41,869, Paris, ca. 1864)',
  'This printed mortgage bond (obligation hypothécaire au porteur) of 400 francs is No. 41,869 of the Compagnie Privilégiée des Ports, Débarcadère Maritime et Terrains de Cadix (Privileged Company of the Ports, Maritime Landing Stage and Lands of Cádiz), Société Anonyme Française, with registered offices on Rue de la Chaussée d\'Antin, Paris. Capital: 10 million francs; total loan: 20 million francs comprising 50,000 mortgage bonds of 400 francs each. Annual interest: 25 francs, payable quarterly. The bond bears a coupon sheet for quarterly interest payments. The company held a French concession to develop the port facilities and surrounding lands of Cádiz in southern Spain, representing a typical example of French capital investment in Spanish infrastructure during the Second Empire period.',
  {
    type: 'Bond',
    subjectCountry: 'Spain',
    issuingCountry: 'France',
    creator: 'Compagnie Privilégiée des Ports, Débarcadère Maritime et Terrains de Cadix',
    issueDate: '1864-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Compagnie Privilégiée des Ports, Débarcadère Maritime et Terrains de Cadix. Obligation hypothécaire, 400 francs. No. 41,869. Capital 10M francs; 20M franc loan / 50,000 bonds. Annual interest 25 francs. Paris, ca. 1864.',
  }
);

// --- Row 422: Portuguese External Debt, Certificate of Interest, £100, 1855 ---
setDoc(422,
  'Portuguese External Debt: Certificate of Interest on £100 Three Per Cent Bond, Letter B (London, December 13, 1855)',
  'This printed certificate is a Certificate of Interest on the sum of £100 in the capital of the Three Per Cent Portuguese External Debt Bond (Letter B), issued under the Decree of December 18, 1852. Issued by the Financial Agent of the Portuguese Government in London (identified as José Jorge Loureiro), and countersigned in London on December 13, 1855. The document certifies that the holder of Bond Letter B is entitled to receive additional interest at one per cent per annum from the date of the arrangement, or at such other rates as provided under the Law of July 20, 1855. Portugal suspended payments on its external debt in the 1840s, and the 1852 Decree launched a restructuring of outstanding obligations; this certificate evidences rights arising from that restructuring.',
  {
    type: 'Bond',
    subjectCountry: 'Portugal',
    issuingCountry: 'Portugal',
    creator: 'Portuguese Government (Financial Agent in London)',
    issueDate: '1855-12-13',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Portuguese External Debt, Certificate of Interest on £100 Three Per Cent Bond (Letter B). Decree of December 18, 1852. Countersigned London, December 13, 1855. Financial Agent: José Jorge Loureiro.',
  }
);

// --- Row 423: Kingdom of Serbia Prize Loan, 10 Gold Dinars, ca. 1908 ---
setDoc(423,
  'Kingdom of Serbia: Prize Loan (Prämien-Anleihe), 10 Gold Francs/Dinars (Series 1922, No. 35)',
  'This trilingual (German/Serbian/French) printed lottery prize bond (Prämien-Obligation / Обвезница са догонима / Obligation à Prime) was issued by the Kingdom of Serbia (Königreich Serbien / Royaume de Serbie). Series 1922, No. 35. Denomination: 10 gold francs (dinars) in gold (Zehn Francs (Dinars) Gold / десет динара у злату / Dix francs (dinars) en or). Signed by the Serbian Royal Finance Ministry (Die kön. Serbische Finanz-Ministerium / Le Ministre des Finances du Royaume de Serbie). With attached coupon sheet on the left. Prize loans (Prämien-Anleihen) combined regular interest payments with lottery draws awarding cash prizes, making them popular with small investors. The trilingual format targeted German, Serbian, and French-speaking investors in pre-World War I Balkan securities markets.',
  {
    type: 'Bond',
    subjectCountry: 'Serbia',
    issuingCountry: 'Serbia',
    creator: 'Kingdom of Serbia, Finance Ministry',
    issueDate: '1908-01-01',
    currency: 'RSD',
    language: 'German, Serbian, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Kingdom of Serbia Prize Loan (Prämien-Anleihe). Trilingual German/Serbian/French. 10 gold francs/dinars. Series 1922, No. 35. Signed by Serbian Royal Finance Ministry.',
  }
);

// --- Row 424: Adelsverein (Verein zum Schutze deutscher Einwanderer in Texas), 500 Gulden Priority Bond, 1850 ---
setDoc(424,
  'Verein zum Schutze deutscher Einwanderer in Texas (Adelsverein): Priority Bond, 500 Gulden (Wiesbaden, July 1, 1850)',
  'This printed document is a Prioritäts-Obligation (priority bond) of 500 gulden at 4% annual interest issued by the Verein zum Schutze deutscher Einwanderer in Texas (Society for the Protection of German Immigrants in Texas), commonly known as the Adelsverein ("nobles\' association"). Wiesbaden, July 1, 1850. The bearer is entitled to a 500-gulden share of 1,600,000 gulden total principal, carrying priority over ordinary Stamm-Actien (common shares) in the Verein\'s assets. Issued under statutes approved by the Herzoglich Nassau\'sche Landes-Regierung, Biebrich, July 23, 1847, per §26 of the Vereins statutes. Signed by Das Comite. The Adelsverein was a society of German nobility and merchants that organized and financed the emigration of over 7,000 Germans to Texas between 1844 and 1847, founding the towns of New Braunfels and Fredericksburg. The venture ultimately failed financially, leaving the society deeply indebted; this bond represents one of its financing instruments. The companion Texas and German Emigration Company stock (cf. row 459) is also in this collection.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'Germany',
    creator: 'Verein zum Schutze deutscher Einwanderer in Texas (Adelsverein)',
    issueDate: '1850-07-01',
    currency: 'NLG',
    language: 'German',
    numberPages: 1,
    period: '19th Century',
    notes: 'Verein zum Schutze deutscher Einwanderer in Texas (Adelsverein). Priority Bond, 500 gulden at 4%. Wiesbaden, July 1, 1850. Total issue 1,600,000 gulden. Priority over common shares. Chartered by Nassau government. Founded New Braunfels and Fredericksburg, Texas.',
  }
);

// --- Row 425: Trésorerie Nationale, Promesse de Mandat Territorial, 25 Francs, 1796 ---
setDoc(425,
  'Trésorerie Nationale: Promesse de Mandat Territorial, 25 Francs (No. 19891, Série 30, An IV/1796)',
  'This printed note is a Promesse de Mandat Territorial (promise of a territorial mandate) of 25 francs, created by the law of 28 Ventôse Year 4 of the French Republic (March 18, 1796). No. 19891, Série 30. Signed by Mugaret and Chevallier of the Trésorerie Nationale (National Treasury). The mandats territoriaux were paper currency instruments issued by the French Directory government to replace the rapidly depreciating assignats (the paper currency backed by nationalized church lands). Like the assignats they replaced, the mandats were backed by confiscated biens nationaux (national lands, formerly ecclesiastical and émigré properties) and were theoretically exchangeable for land at fixed prices. However, the mandats collapsed even faster than the assignats—within months they had lost 97% of their face value—and were withdrawn from circulation in 1797. This "Promesse de Mandat" was a preliminary issue, circulating before the full mandat territorial notes were printed, and is among the rarer instruments of French Revolutionary monetary history.',
  {
    type: 'Promissory Note',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Trésorerie Nationale (French Republic)',
    issueDate: '1796-03-18',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Trésorerie Nationale. Promesse de Mandat Territorial, 25 francs. No. 19891, Série 30. Law of 28 Ventôse An IV (March 18, 1796). Signed by Mugaret and Chevallier. French Revolutionary paper currency.',
  }
);

// --- Row 426: Reconstitution de Rentes à 3 Pour Cent, French Royal, Paris, July 20, 1763 ---
setDoc(426,
  'Reconstitution de Rentes à 3 Pour Cent: Royal French Annuity Certificate (Paris, July 20, 1763)',
  'This handwritten and witnessed notarial certificate reconstitutes (replaces lost or damaged originals of) 3% royal rentes (annual interest bonds) created by the Edict of May 1760, secured on the royal leather tax revenues (Denier sur les Cuirs, established by the Edict of April 1759) and other royal revenues. Executed before royal notaries at the Châtelet de Paris, July 20, 1763. Witnessed by several prominent Parisian financiers. The French crown issued these reconstitution certificates to replace rente holders whose original documents had been lost or destroyed, maintaining their claim to the annual income stream. Issued during the financially ruinous Seven Years\' War (1756–63), these rentes were part of the enormous debt burden that would ultimately contribute to the fiscal crisis leading to the French Revolution.',
  {
    type: 'Bond',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'France. Trésor Royal (Royal Treasury)',
    issueDate: '1763-07-20',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Reconstitution de Rentes à 3 Pour Cent. Royal French annuity certificate. Paris, July 20, 1763. Secured on royal leather tax (Denier sur les Cuirs, Edit d\'Avril 1759) and other royal revenues. Executed before notaries at the Châtelet de Paris. Seven Years\' War era fiscal debt.',
  }
);

// --- Row 427: Deutsches Reich Reichsbanknote, 1,000 Marks, April 21, 1910 ---
setDoc(427,
  'Deutsches Reich: Reichsbanknote, Ein Tausend Mark (No. 6125588N, Berlin, April 21, 1910)',
  'This printed note is a one-thousand mark (Ein Tausend Mark) Reichsbanknote. Serial No. 6,125,588N. Berlin, April 21, 1910. Issued by the Reichsbankdirektorium (Imperial Bank Directorate) under signatures of the Bank directors and deputy directors. The note states: "zahlt die Reichsbankhauptkasse in Berlin ohne Legitimationsprüfung dem Einlieferer dieser Banknote" (the Reichsbank main cashier in Berlin will pay the bearer of this banknote without identity verification). This pre-World War I German 1,000-mark note was one of the largest-denomination banknotes in regular circulation before 1914. The Reichsbank was Germany\'s central bank, established in 1876 upon German unification. This note was printed in the era of the classical gold standard, when the Reichsmark was fully convertible to gold.',
  {
    type: 'Banknote',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Reichsbank (German Imperial Bank)',
    issueDate: '1910-04-21',
    currency: 'DEM',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Deutsches Reich Reichsbanknote. 1,000 Marks (Ein Tausend Mark). No. 6125588N. Berlin, April 21, 1910. Reichsbankdirektorium.',
  }
);

// --- Row 428: Deutsches Reich Reichsbanknote, 1,000 Marks, September 15, 1922 (Hyperinflation) ---
setDoc(428,
  'Weimar Republic: Reichsbanknote, Tausend Mark (Pa 688969, Berlin, September 15, 1922)',
  'This printed note is a one-thousand mark (Tausend Mark) Reichsbanknote of the Weimar Republic. Serial No. Pa 688969, Series KH. Berlin, September 15, 1922. Issued by the Reichsbankdirektorium. The note carries the notice: "Vom 1. Januar 1923 ab kann diese Banknote aufgerufen und unter Umtausch gegen andere gesetzliche Zahlungsmittel eingezogen werden" (From January 1, 1923, this banknote may be called in and exchanged for other legal tender). This note was issued during the onset of the great German hyperinflation of 1921–23, during which the value of the mark collapsed catastrophically—by November 1923 the exchange rate reached 4.2 trillion marks per US dollar. The 1,000-mark note, a large denomination in 1910, had become almost worthless by the date of this issue. This note illustrates the extreme paper money production of the Weimar inflationary period.',
  {
    type: 'Banknote',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Reichsbank (German Weimar Republic Bank)',
    issueDate: '1922-09-15',
    currency: 'DEM',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Weimar Republic Reichsbanknote. 1,000 Marks (Tausend Mark). Pa 688969, KH. Berlin, September 15, 1922. Hyperinflation era. Note callable from January 1, 1923.',
  }
);

// --- Row 429: République Chinoise 5% Gold Bond 1925, $50 U.S. Gold Dollars ---
setDoc(429,
  'République Chinoise: 5% Gold Bond of 1925, $50 U.S. Gold Dollars (No. 350,356, London, May 27, 1925)',
  'This bilingual (French/English) printed bearer bond (bon au porteur / to bearer) of $50 U.S. gold dollars at 5% annual interest is No. 350,356 of the Republic of China 5% Gold Bond of 1925. London, May 27, 1925. Issued as part of the international borrowings of the Republic of China in the mid-1920s to fund railway development and other national projects. Denominated in U.S. gold dollars and issued through London, this bond was designed to attract international investors seeking security in gold-standard currency. Chinese sovereign bonds of the Republic era were subject to significant default risk given the political instability of the warlord period.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Republic of China',
    issueDate: '1925-05-27',
    currency: 'USD',
    language: 'French, English',
    numberPages: 1,
    period: '20th Century',
    notes: 'République Chinoise 5% Gold Bond 1925. $50 U.S. gold dollars. No. 350,356. London, May 27, 1925.',
  }
);

// --- Row 430: République Chinoise, 8% Lung-Tsing-U-Hai Railway Bond, 500 Francs, Brussels, 1921 ---
setDoc(430,
  'Gouvernement de la République Chinoise: 8% Treasury Bond 1921, Lung-Tsing-U-Hai Railway, 500 Francs (No. 012918, Brussels, July 1, 1921)',
  'This printed bearer bond (bon du trésor) of 500 francs at 8% annual interest is No. 012918 of the Republic of China 8% Lung-Tsing-U-Hai Railway (隴秦豫海鐵路) Loan of 1921. Total issue: 50,000,000 francs, divided into 100,000 bonds of 500 francs each. Brussels, July 1, 1921. Signed by the Chinese Finance Minister and the Compagnie Générale de Chemins de Fer et Tramways en Chine (General Company of Railways and Tramways in China). The Lung-Tsing-U-Hai Railway (Gansu–Shaanxi–Henan–Jiangsu) was a major planned trans-China line. This Belgian franc issue complemented Dutch guilder (cf. row 395) and other currency issues for the same railway, reflecting the multi-country European capital-raising strategy of the Chinese Republic. The Belgian-based Compagnie Générale de Chemins de Fer et Tramways en Chine was the principal financial intermediary.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'China',
    creator: 'Government of the Republic of China',
    issueDate: '1921-07-01',
    currency: 'BEF',
    language: 'French, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: 'République Chinoise 8% Lung-Tsing-U-Hai Railway (隴秦豫海鐵路) Treasury Bond 1921. 500 francs. No. 012918. Brussels, July 1, 1921. Total issue 50,000,000 francs / 100,000 bonds. Compagnie Générale de Chemins de Fer et Tramways en Chine.',
  }
);

// --- Row 431: Royaume de Bulgarie 4.5% Gold Loan 1909, 500 Gold Francs ---
setDoc(431,
  'Royaume de Bulgarie: 4½% Or State Loan of 1909, 500 Gold Francs / Cinq Cents Leva Zlatni (No. 055,199)',
  'This bilingual (Bulgarian/French) printed bearer bond (Облигация / Obligation) No. 055,199 is a 500-gold-franc denomination of the amortizable 4½% Gold State Loan of Bulgaria, 1909. Denomination: 500 gold francs = 500 Bulgarian gold leva = 476 Austrian crowns = 405 German marks = 19.163 British pounds sterling = 9.65 Dutch guilders. Total emission represented by 200,000 bonds. Signed by the Bulgarian Finance Ministry. The 1909 loan was issued shortly after Bulgaria declared full independence from the Ottoman Empire (October 5, 1908), requiring repayment of an Ottoman tribute and new international capital for state development. The multiple-currency denomination equivalences reflect the gold-standard international bond market of the Edwardian era.',
  {
    type: 'Bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Kingdom of Bulgaria, Finance Ministry',
    issueDate: '1909-01-01',
    currency: 'BGN',
    language: 'Bulgarian, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Royaume de Bulgarie 4½% Or State Loan 1909. 500 gold francs = 500 leva zlatni. No. 055,199. 200,000 bonds total. Multiple-currency equivalences stated.',
  }
);

// --- Row 432: Russian General Oil Corporation, Share Warrant No. 1625, London, 1913 ---
setDoc(432,
  'Russian General Oil Corporation (Société Générale Naphthifère Russe) Limited: Share Warrant to Bearer, £1 (No. 1625, London, 1913)',
  'This trilingual (English/French/Russian) printed share warrant to bearer entitles the holder to one fully paid share of £1 sterling in the Russian General Oil Corporation (Société Générale Naphthifère Russe) Limited. Capital $2,500,000 in 2,500,000 shares of £1 each. No. 1625. Given under the Common Seal of the Company. London, [1913]. An annotation dated April 23, 1954 (possibly a later registration or claim by émigré owners) appears at the bottom. The company was incorporated in England to exploit oil properties in the Russian Empire, part of the wave of British-registered companies active in the Baku and Grozny oil regions of the Caucasus before World War I. The trilingual format reflects the international nature of the ownership and investor base.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'United Kingdom',
    creator: 'Russian General Oil Corporation (Société Générale Naphthifère Russe) Limited',
    issueDate: '1913-01-01',
    currency: 'GBP',
    language: 'English, French, Russian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Russian General Oil Corporation. Share warrant to bearer, £1. No. 1625. Capital $2,500,000 / 2,500,000 shares. London, 1913. Trilingual English/French/Russian. Annotation dated April 23, 1954.',
  }
);

// --- Row 433: Imperial Russian Government State 4% Rente, 1,000 Rubles ---
setDoc(433,
  'Imperial Russian Government: State 4% Rente (Государственная 4% Рента), Certificate No. 02756, 1,000 Rubles',
  'This printed certificate (свидетельство) No. 02756 represents 1,000 rubles face value in the Imperial Russian Government State 4% Rente (Государственная 4% Рента), yielding 40 rubles annual income. Series 191 (Сто девяносто первая / one hundred and ninety-first), total series value 10,000,000 rubles. Bearer certificate (на предъявителя). Decorated in a striking pink-and-brown Art Nouveau design with the Imperial Russian double-eagle. The State Rente was a perpetual or long-term government bond paying a fixed annual income (renta), modeled on French and British consolidated annuities. Russian state rentes were widely held domestically and abroad as conservative income securities.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1900-01-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Imperial Russian Government State 4% Rente (Государственная 4% Рента). Certificate No. 02756. 1,000 rubles (40 rubles annual income). Series 191. Bearer certificate. Art Nouveau design.',
  }
);

// --- Row 434: Russisk-Norsk Skogindustri A/S, Share No. 0349, 1,000 Rubles, Petrograd, 1917 ---
setDoc(434,
  'Russisk-Norsk Skogindustri A/S (Russian-Norwegian Timber Industry Company): Share No. 0349, 1,000 Rubles (Petrograd, 1917)',
  'This bilingual (Russian/Norwegian) printed bearer share No. 0349 of 1,000 rubles represents one share in the Русско-Норвежское Лесопромышленное Акционерное Общество (Russian-Norwegian Timber Industry Joint-Stock Company / Russisk-Norsk Skogindustri A/S). Capital 3,000,000 rubles. Charter (Устав) approved December 29, 1916. Issued to Herr Jorgen Bogaland. Petrograd, 1917. The company exploited timber resources in northern Russia and was a product of Norwegian-Russian commercial relations in the timber trade, which had been active since the nineteenth century. This share was issued in 1917—the year of the Russian Revolution—and the company was likely nationalized by the Bolsheviks shortly thereafter.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Russisk-Norsk Skogindustri A/S (Russian-Norwegian Timber Industry Company)',
    issueDate: '1917-01-01',
    currency: 'RUB',
    language: 'Russian, Norwegian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Russisk-Norsk Skogindustri A/S. Share No. 0349, 1,000 rubles. Capital 3,000,000 rubles. Charter approved December 29, 1916. Issued to Herr Jorgen Bogaland. Petrograd, 1917.',
  }
);

// --- Row 435: Stadtanleihe Kremnitz (Kremnica Municipal Bond), 5% Bond, 10,000 Kronen ---
setDoc(435,
  'Stadtanleihe Kremnitz (Kremnica Municipal Bond): 5% Bearer Bond, 10,000 Kronen (Serie 000513, No. 011)',
  'This printed bearer bond (5%ige Schuldverschreibung) of 10,000 kronen at 5% annual interest is a municipal bond issued by the town of Kremnitz (Kremnica, present-day Slovakia), then part of the Kingdom of Hungary within the Austro-Hungarian Empire. Serie 000513, No. 011. The bond features a decorative Art Nouveau design with a panoramic engraved view of the city. Kremnica (Kremnitz in German, Körmöcbánya in Hungarian) is a historic Slovak mining town, site of one of the oldest continuously operating mints in Europe (the Kremnica Mint, founded ca. 1328), and was an important economic center of Upper Hungary. Municipal bonds (Stadtanleihen) were a common form of financing for Central European towns during the late nineteenth and early twentieth centuries.',
  {
    type: 'Bond',
    subjectCountry: 'Slovakia',
    issuingCountry: 'Slovakia',
    creator: 'Stadtgemeinde Kremnitz (Municipality of Kremnica)',
    issueDate: '1910-01-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Stadtanleihe Kremnitz (Kremnica Municipal Bond). 5% bond, 10,000 kronen. Serie 000513, No. 011. Austria-Hungary / Kingdom of Hungary. Art Nouveau design with city panorama. Ca. 1910.',
  }
);

// --- Row 436: Walchensee-Anleihe (Walchenseewerk/Mittlere Isar/Bayernwerk), 10,000 Marks, 1922 ---
setDoc(436,
  'Walchensee-Anleihe: Schuldverschreibung, 10,000 Mark (Buchstabe E, No. 419133, Munich, February 28, 1922)',
  'This printed bearer bond (Schuldverschreibung) of 10,000 marks at minimum 7% annual interest per year was issued jointly by three Bavarian electricity companies—Walchenseewerk A.G., Mittlere Isar A.G., and Bayernwerk A.G.—as the Walchensee-Anleihe (Walchensee Loan). Buchstabe E, No. 419133. Munich, February 28, 1922. The loan financed the construction of the Walchensee hydroelectric power plant (Walchenseekraftwerk) in the Bavarian Alps, which when completed in 1924 was one of the most powerful hydroelectric plants in Europe. The three issuing companies were state-financed utilities instrumental in Bavaria\'s post-World War I industrialization. This bond was issued during the Weimar hyperinflation, which would eventually render the nominal 10,000-mark denomination nearly worthless.',
  {
    type: 'Bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Walchenseewerk A.G. / Mittlere Isar A.G. / Bayernwerk A.G.',
    issueDate: '1922-02-28',
    currency: 'DEM',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Walchensee-Anleihe. 10,000 marks, min. 7% p.a. Buchstabe E, No. 419133. Munich, February 28, 1922. Issued by Walchenseewerk A.G., Mittlere Isar A.G., Bayernwerk A.G. Finances Walchensee hydroelectric power plant, Bavaria.',
  }
);

// --- Row 437: Compañía del Ferro-Carril de Sevilla–Jerez–Cadiz, Action 500 Francs ---
setDoc(437,
  'Compañía del Ferro-Carril de Sevilla–Jerez–Cadiz (Séville–Xérès–Cadix): Action de 500 Francs / 1,900 Reales (No. 40,721)',
  'This bilingual (Spanish/French) printed bearer share (acción al portador / action au porteur) No. 40,721 of 500 francs (= 1,900 reales vellón) is a share in the Compañía Anónima de los Ferro-Carriles de Sevilla–Jerez–Cadiz (Compagnie Anonyme des Chemins de Fer de Séville–Xérès–Cadix). Authorized by Royal Decree (Spain) and by the French Crown, May 4, 1857. Capital: 266,000,000 reales = 70,000,000 francs, divided into 140,000 shares. Duration: 99 years. The Seville–Jerez–Cádiz railway was one of the first railways in Andalusia, connecting the great sherry wine-producing region of Jerez de la Frontera (known in French as Xérès, giving its name to the wine "sherry") to the major ports of Seville and Cádiz. The bilingual format reflects French investor participation. The certificate includes an attached coupon sheet at the bottom.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Spain',
    issuingCountry: 'Spain',
    creator: 'Compañía del Ferro-Carril de Sevilla–Jerez–Cadiz',
    issueDate: '1857-01-01',
    currency: 'FRF',
    language: 'Spanish, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Compañía del Ferro-Carril de Sevilla–Jerez–Cadiz. Bilingual Spanish/French. 500 francs = 1,900 reales. No. 40,721. Capital 70,000,000 francs / 140,000 shares. Royal Decree May 4, 1857. With coupon sheet.',
  }
);

// --- Row 438: Shanghai-Nanking Railway Net Profit Sub-Certificate, London, 1904 ---
setDoc(438,
  'Shanghai-Nanking Railway: Net Profit Sub-Certificate No. E3036 (British & Chinese Corporation Limited, London, December 2, 1904)',
  'This printed certificate is a Net Profit Sub-Certificate No. E3036 issued by the British & Chinese Corporation Limited, entitling the bearer to participate in the net profits of the Shanghai-Nanking Railway (今上海-南京铁路) as secured by a Declaration of Trust dated July 10, 1904, made between the British and Chinese Corporation Limited (of one part) and the Hongkong and Shanghai Banking Corporation (of the other part). One of not more than 32,500 such sub-certificates. London, December 2, 1904. The Shanghai-Nanking Railway was a key trunk line in the Yangtze Delta region of China, running between Shanghai and Nanjing; construction was largely British-financed. The British & Chinese Corporation Limited was a joint venture of the Hongkong and Shanghai Banking Corporation and Jardine, Matheson & Co. formed to underwrite Chinese railway and development loans.',
  {
    type: 'Certificate',
    subjectCountry: 'China',
    issuingCountry: 'United Kingdom',
    creator: 'British & Chinese Corporation Limited',
    issueDate: '1904-12-02',
    currency: '',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Shanghai-Nanking Railway Net Profit Sub-Certificate. No. E3036. British & Chinese Corporation Limited. Declaration of Trust, July 10, 1904 (with Hongkong and Shanghai Banking Corporation). London, December 2, 1904. One of max. 32,500 sub-certificates.',
  }
);

// --- Row 439: Sherman & Barnsdall Oil Company, New York, 1865 ---
setDoc(439,
  'Sherman & Barnsdall Oil Company: Stock Certificate (New York, May 1865)',
  'This engraved stock certificate in the Sherman & Barnsdall Oil Company, Capital $750,000, was issued in New York in May 1865. Bearing a U.S. Internal Revenue Civil War revenue stamp. The certificate records a transfer to James H. Arbona from James Irving, New York, June 27, 1868, by H.C. Brolasky. Sherman & Barnsdall was an early Pennsylvania petroleum company active in the oil regions of western Pennsylvania during the first decade of the oil industry\'s explosive growth following Edwin Drake\'s 1859 discovery at Titusville. The period 1863–69 saw enormous speculative activity in oil company stocks, with dozens of companies incorporated to exploit the Pennsylvania oil fields. This certificate, issued in the immediate aftermath of the Civil War, represents the intersection of oil speculation and wartime economic expansion.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Sherman & Barnsdall Oil Company',
    issueDate: '1865-05-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Sherman & Barnsdall Oil Company. Capital $750,000. New York, May 1865. With U.S. Civil War revenue stamp. Transfer from James Irving to James H. Arbona, June 27, 1868, by H.C. Brolasky. Pennsylvania oil boom era.',
  }
);

// --- Row 440: Сибирский Торговый Банк (Commercial Bank of Siberia), Share No. 10691, 1906 ---
setDoc(440,
  'Сибирский Торговый Банк (Commercial Bank of Siberia / Banque de Commerce de Sibérie): Share No. 10691, 250 Rubles (St. Petersburg, 1906)',
  'This printed quadrilingual (Russian/German/French/English) bearer share No. 10691 of 250 rubles is a share in the Сибирский Торговый Банк (Sibirische Handels-Bank / Banque de Commerce de Sibérie / Commercial Bank of Siberia). Charter (Устав) approved June 28, 1872. St. Petersburg, 1906. The four languages used in the border text—Russian, German, French, and English—reflect the international investor base this bank sought to attract. The Commercial Bank of Siberia was one of the major commercial banks of the Russian Empire, with its operational center in Siberia and Trans-Baikal region, and its administrative headquarters in St. Petersburg. It financed trade and industrial development in Siberia, including the growing Trans-Siberian Railway commerce.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Сибирский Торговый Банк (Commercial Bank of Siberia)',
    issueDate: '1906-01-01',
    currency: 'RUB',
    language: 'Russian, German, French, English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Сибирский Торговый Банк (Commercial Bank of Siberia / Banque de Commerce de Sibérie / Sibirische Handels-Bank). Share No. 10691, 250 rubles. Charter approved June 28, 1872. St. Petersburg, 1906. Quadrilingual Russian/German/French/English.',
  }
);

// --- Row 441: Société Anonyme des Verreries d'Extrême Orient, Japan, 1927 ---
setDoc(441,
  'Société Anonyme des Verreries d\'Extrême Orient: Action No. 8339 (Capital ¥1,500,000, Japan, 1927)',
  'This bilingual (French/Japanese) printed bearer share No. 8339 is a share in the Société Anonyme des Verreries d\'Extrême Orient (Far East Glassworks Company). Capital ¥1,500,000. Incorporated under Japanese law, the society was constituted in March 1927 before a Japanese commercial court as stated on the certificate. The company manufactured glass in Japan or the Japanese empire, and its French corporate form reflects the involvement of French investors and management. The Verreries d\'Extrême Orient operated in the Far East glass manufacturing market during the interwar period, supplying glass products to the rapidly industrializing economies of East Asia.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Japan',
    issuingCountry: 'Japan',
    creator: 'Société Anonyme des Verreries d\'Extrême Orient',
    issueDate: '1927-01-01',
    currency: 'JPY',
    language: 'French, Japanese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société Anonyme des Verreries d\'Extrême Orient. Action No. 8339. Capital ¥1,500,000. Incorporated under Japanese jurisdiction, March 1927. Bilingual French/Japanese. Glass manufacturing in the Far East.',
  }
);

// --- Row 442: Société Anonyme Minière des Aimaks de Touchetoukhan et de Tsetsenkhan en Mongolie, 1911 ---
setDoc(442,
  'Société Anonyme Minière des Aimaks de Touchetoukhan et de Tsetsenkhan en Mongolie: Action, 50 Rubles (No. 07520, St. Petersburg, 1911)',
  'This trilingual (Russian/French/Chinese) printed bearer share No. 07520 of 50 rubles is a share in the Акционерное Общество Рудного Дела Тушетухановского и Цэцэнхановского Аймаков в Монголии (Société Anonyme Minière des Aimaks de Touchetoukhan et de Tsetsenkhan en Mongolie / Mining Company of the Tushetu Khan and Tsetsenkhan Aimaks in Mongolia). The Chinese border text reads 蒙古郭爾河等處礦務股票 (Mongolia Goer River Area Mining Company Share). St. Petersburg, 1911. The company exploited mineral deposits in the Khalkha Mongolian aimaks (provinces) of Tushetu Khan and Tsetsenkhan, areas of Inner/Outer Mongolia then under nominal Qing Chinese suzerainty and Russian commercial penetration. The trilingual format reflects the overlapping Russian, French, and Chinese interests in Mongolian mining.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Mongolia',
    issuingCountry: 'Russia',
    creator: 'Société Anonyme Minière des Aimaks de Touchetoukhan et de Tsetsenkhan en Mongolie',
    issueDate: '1911-01-01',
    currency: 'RUB',
    language: 'Russian, French, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société Anonyme Minière des Aimaks de Touchetoukhan et de Tsetsenkhan en Mongolie. Action No. 07520, 50 rubles. Trilingual Russian/French/Chinese. St. Petersburg, 1911. Mining in Khalkha Mongolia.',
  }
);

// --- Row 443: Société Atlantique de Réassurances, Tanger, ca. 1950s ---
setDoc(443,
  'Société Atlantique de Réassurances: Action de 1,000 Francs au Porteur (No. 69141, Tanger, ca. 1952)',
  'This printed bearer share (action de mille francs au porteur, entièrement libérée) No. 69141 is a share in the Société Atlantique de Réassurances (Atlantic Reinsurance Society), Société Anonyme, registered at 17, Rue Goya, Tanger (the International Zone of Tangier, Morocco). Capital 270,000,000 francs divided into 270,000 shares of 1,000 francs each. With a large coupon sheet of 18 coupons attached for annual dividend payments. The International Zone of Tangier (1923–56) had a unique political status—governed jointly by several European powers and Morocco—and was a center of international business, finance, and speculation, particularly attractive for offshore companies and reinsurance operations.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Morocco',
    issuingCountry: 'Morocco',
    creator: 'Société Atlantique de Réassurances',
    issueDate: '1952-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société Atlantique de Réassurances. Action au porteur, 1,000 francs. No. 69141. Capital 270,000,000 francs / 270,000 shares. Siège social: 17 Rue Goya, Tanger (International Zone). With 18-coupon sheet. Ca. 1952.',
  }
);

// --- Row 444: Société Belge d'Entreprises en Chine, Action 500 Francs, Brussels, 1924 ---
setDoc(444,
  'Société Belge d\'Entreprises en Chine: Action de 500 Francs au Porteur (No. 0628, Brussels, March 19, 1924)',
  'This printed bearer share (action de 500 francs au porteur) No. 0628 is a share in the Société Belge d\'Entreprises en Chine (Belgian Company for Enterprises in China), Société Anonyme, with registered office in Brussels. Capital social: 3,000,000 francs represented by 6,000 shares of 500 francs each. Brussels, March 19, 1924. Belgium had significant commercial and financial interests in China during the early twentieth century, including concessions in Tianjin and involvement in Chinese railway finance. The Société Belge d\'Entreprises en Chine was a holding or operating company engaged in Belgian commercial activities in China during the Republican period.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'China',
    issuingCountry: 'Belgium',
    creator: 'Société Belge d\'Entreprises en Chine',
    issueDate: '1924-03-19',
    currency: 'BEF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société Belge d\'Entreprises en Chine. Action au porteur, 500 francs. No. 0628. Capital 3,000,000 francs / 6,000 shares. Brussels, March 19, 1924.',
  }
);

// --- Row 445: Bulgarian Red Cross Society Lottery Bond, 20 Gold Leva ---
setDoc(445,
  'Bulgarian Red Cross Society (Дружество "Червен Кръст"): Lottery Bond, 20 Gold Leva (No. 05331)',
  'This trilingual (Bulgarian/French/Russian) printed lottery bond (Облигация / Obligation / Облигация) No. 05331 of 20 gold leva was issued by the Българско дружество "Червен Кръст" (Société Bulgare de la Croix Rouge / Bulgarian Red Cross Society) as part of a 4,000,000 gold-leva prize loan (задоженъ заем по именни / emprunt à termes à lots). With coupon sheet on the right side. Red Cross and humanitarian organizations frequently issued prize lottery bonds in the Balkan states during the late nineteenth and early twentieth centuries as a fundraising mechanism, combining charitable appeal with the attraction of lottery prizes. The trilingual format (Bulgarian, French, Russian) reflects Bulgaria\'s diplomatic alignment and the international Red Cross movement.',
  {
    type: 'Bond',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: 'Българско дружество "Червен Кръст" (Bulgarian Red Cross Society)',
    issueDate: '1910-01-01',
    currency: 'BGN',
    language: 'Bulgarian, French, Russian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Bulgarian Red Cross Society (Дружество "Червен Кръст"). Lottery bond, 20 gold leva. No. 05331. Total loan 4,000,000 gold leva. Trilingual Bulgarian/French/Russian. With coupon sheet.',
  }
);

// --- Row 446: Société de Lovitch des Produits et Engrais Chimiques, 1st Emission, 250 Rubles, 1895 ---
setDoc(446,
  'Société de Lovitch des Produits et Engrais Chimiques: Action, 250 Rubles (No. 2106, Warsaw, 1895) [1st Emission]',
  'This bilingual (Russian/French) printed bearer share No. 2106 of 250 rubles represents one share in the Ловичское Общество Химических Продуктов и Землеудобрительных Веществ (Société de Lovitch des Produits et Engrais Chimiques / Łowicz Chemical Products and Fertilizer Company). Capital social: 600,000 rubles. 1st emission. Charter approved by H.I.M. (His Imperial Majesty, Tsar Alexander III) on June 9, 1895. Warsaw (Варшава), 1895. Łowicz (Russian: Lovitch) is a town in central Poland, then part of Russian-controlled Congress Poland. The company produced chemical fertilizers and agricultural chemicals, reflecting the late nineteenth-century expansion of industrial agriculture in Central Europe.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Poland',
    issuingCountry: 'Poland',
    creator: 'Société de Lovitch des Produits et Engrais Chimiques',
    issueDate: '1895-01-01',
    currency: 'RUB',
    language: 'Russian, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Société de Lovitch des Produits et Engrais Chimiques. 1st emission. Action No. 2106, 250 rubles. Capital 600,000 rubles. Charter approved June 9, 1895. Warsaw. Chemical fertilizers company.',
  }
);

// --- Row 447: Société de Lovitch, 2nd Emission, 250 Rubles, 1895 ---
setDoc(447,
  'Société de Lovitch des Produits et Engrais Chimiques: Action, 250 Rubles (No. 3276, Warsaw, 1895) [2nd Emission]',
  'This bilingual (Russian/French) printed bearer share No. 3276 of 250 rubles represents one share in the 2nd emission (2-й выпуск / 2ème émission) of the Ловичское Общество Химических Продуктов и Землеудобрительных Веществ (Société de Lovitch des Produits et Engrais Chimiques / Łowicz Chemical Products and Fertilizer Company). Capital increased to 1,000,000 rubles for this second issue. Charter approved June 9, 1895. Warsaw, 1895. Companion piece to the 1st emission share No. 2106 (cf. row 446). The capital increase from 600,000 to 1,000,000 rubles between the two emissions indicates the company expanded its operations following its initial establishment.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Poland',
    issuingCountry: 'Poland',
    creator: 'Société de Lovitch des Produits et Engrais Chimiques',
    issueDate: '1895-01-01',
    currency: 'RUB',
    language: 'Russian, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Société de Lovitch des Produits et Engrais Chimiques. 2nd emission. Action No. 3276, 250 rubles. Capital 1,000,000 rubles. Charter approved June 9, 1895. Warsaw. Companion to 1st emission No. 2106 (row 446).',
  }
);

// --- Row 448: Общество Столичнаго Освещения (Capital Lighting Company), Share No. 39419, 1858 ---
setDoc(448,
  'Общество Столичнаго Освещения (Société d\'Eclairage de la Capitale): Share No. 39419, 100 Silver Rubles (St. Petersburg, ca. 1858)',
  'This bilingual (Russian/French) printed bearer share No. 39419 of 100 silver rubles is a share in the Общество Столичнаго Освещения (Société d\'Eclairage de la Capitale / Capital Lighting Company). Capital: 4,000,000 silver rubles divided into 40,000 shares of 100 silver rubles each. Charter (Высочайше утвержденное) highest-approved October 10, 1858. Issued to Byron Williams John Plytho (?) / Johnston Robinson Esqre. The Общество Столичнаго Освещения was chartered in 1858 to provide gas lighting for St. Petersburg, marking the beginning of organized municipal gas lighting in the Russian imperial capital. It was among the pioneering utility companies in Russia and one of the first public utilities to be organized as a joint-stock company in the Empire.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Общество Столичнаго Освещения (Capital Lighting Company)',
    issueDate: '1858-01-01',
    currency: 'RUB',
    language: 'Russian, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Общество Столичнаго Освещения (Société d\'Eclairage de la Capitale). Share No. 39419, 100 silver rubles. Capital 4,000,000 silver rubles / 40,000 shares. Charter approved October 10, 1858. St. Petersburg gas lighting company.',
  }
);

// --- Row 449: Société des Mines d'Or de Kilo-Moto, Part Bénéficiaire, Belgian Congo, ca. 1944 ---
setDoc(449,
  'Société des Mines d\'Or de Kilo-Moto: Part Bénéficiaire No. 1191340 (Belgian Congo, ca. 1944)',
  'This printed profit-participation share (part bénéficiaire sans désignation de valeur, "without designation of value") No. 1191340 was issued by the Société des Mines d\'Or de Kilo-Moto, Société Congolaise à Responsabilité Limitée (Belgian Congo limited liability company). Siège Social: Kilo, Belgian Congo; Siège Administratif: Brussels. Capital social: 230,000,000 francs. "Titre créé après le 6-10-1944" (certificate created after October 6, 1944). The Kilo-Moto mines are located in what is now the Ituri and Haut-Uele provinces of the northeastern Democratic Republic of Congo and were among the most productive gold mines in sub-Saharan Africa throughout the twentieth century. The mines were operated under Belgian colonial concession until independence in 1960 and subsequently under Zairian/Congolese state control.',
  {
    type: 'Certificate',
    subjectCountry: 'Democratic Republic of Congo',
    issuingCountry: 'Belgium',
    creator: 'Société des Mines d\'Or de Kilo-Moto',
    issueDate: '1944-10-06',
    currency: 'BEF',
    language: 'French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Société des Mines d\'Or de Kilo-Moto. Part bénéficiaire sans désignation de valeur. No. 1191340. Capital 230,000,000 francs. Belgian Congo. "Titre créé après le 6-10-1944." Major gold mining concession, Ituri/Haut-Uele.',
  }
);

// --- Row 450: Société Franco-Péruvienne des Mines de Castro-Virreyna, 500 Francs, 1852 (Liquidated) ---
setDoc(450,
  'Société Franco-Péruvienne des Mines de la Province de Castro-Virreyna: Action de 500 Francs (No. 1,735, Paris, January 9, 1852) [Liquidated]',
  'This printed share (action) of 500 francs is No. 1,735 of the Société Franco-Péruvienne des Mines de la Province de Castro-Virreyna, en commandite sous la raison sociale Crosnier & Cie. Capital: 1,000,000 francs (2,000 shares of 500 francs). Issued to Monsieur Blaryez François Charles Bouguier. Paris, January 9, 1852. Stamped "Remboursé par le Liquidateur" (repaid through liquidation). The Castro-Virreyna region of central Peru was historically one of the richest silver and gold mining areas in South America, home to mines exploited since the early colonial period. This French mining venture in Peru represented the broader wave of European mining investment in Latin America following the independence movements. The liquidation stamp indicates the company was wound up—a common fate for many foreign mining ventures in Peru, where complex land rights, geological difficulties, and political instability frustrated investors.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Peru',
    issuingCountry: 'France',
    creator: 'Société Franco-Péruvienne des Mines de la Province de Castro-Virreyna (Crosnier & Cie.)',
    issueDate: '1852-01-09',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Société Franco-Péruvienne des Mines de la Province de Castro-Virreyna (Crosnier & Cie.). Action No. 1,735, 500 francs. Capital 1,000,000 francs. Paris, January 9, 1852. Issued to Blaryez François Charles Bouguier. Stamped REMBOURSÉ PAR LE LIQUIDATEUR.',
  }
);

// --- Row 451: Société Générale Égyptienne, Dixième de Part de Fondateur, No. 02524 ---
setDoc(451,
  'Société Générale Égyptienne pour l\'Agriculture & le Commerce: Dixième de Part de Fondateur, No. 02524',
  'This printed bearer certificate represents one-tenth of a founder\'s share (dixième de part de fondateur au porteur, sans désignation de valeur) No. 02524 in the Société Générale Égyptienne pour l\'Agriculture & le Commerce (Egyptian General Society for Agriculture and Commerce), Société Anonyme. Capital initially 12,500,000 francs, subsequently increased to 15,000,000 francs (60,000 shares of 250 francs each, plus founder\'s parts divided into tenths). The company was constituted by notarial act before M. Amenosse-Louis-Jean-Collin, notary in Antwerp, Belgium. The Société Générale Égyptienne operated in Egypt during the Khedivial period (likely 1860s–1870s) in agricultural development and commerce. "Parts de fondateur" (founder\'s parts) without par value entitled holders to a share of surplus profits above the ordinary dividend.',
  {
    type: 'Certificate',
    subjectCountry: 'Egypt',
    issuingCountry: 'Belgium',
    creator: 'Société Générale Égyptienne pour l\'Agriculture & le Commerce',
    issueDate: '1870-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Société Générale Égyptienne pour l\'Agriculture & le Commerce. Dixième de part de fondateur. No. 02524. Capital 15,000,000 francs. Constituted before notary in Antwerp. Khedivial Egypt, ca. 1870.',
  }
);

// --- Row 452: The Spassky Copper Mine Limited, Share Warrant, London, 1917 ---
setDoc(452,
  'The Spassky Copper Mine Limited: Share Warrant to Bearer, £1 (No. 68868, London, May 10, 1917)',
  'This bilingual (English/French) printed share warrant to bearer entitles the holder to one fully paid share of £1 in The Spassky Copper Mine Limited. Capital £1,250,000 in 1,250,000 shares of £1 each. No. 68868. Given under the Common Seal of the Company. London, May 10, 1917. The Spassky copper mine was located in the Karaganda region of the Kazakh steppe (present-day central Kazakhstan), one of the richest copper deposits in Central Asia. Incorporated in the United Kingdom, the company was part of the wave of British investment in Russian and Central Asian mining ventures during the late imperial period. Issued in May 1917—two months after the February Revolution that overthrew the Tsar—the company was subsequently nationalized by the Bolsheviks.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Kazakhstan',
    issuingCountry: 'United Kingdom',
    creator: 'The Spassky Copper Mine Limited',
    issueDate: '1917-05-10',
    currency: 'GBP',
    language: 'English, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'The Spassky Copper Mine Limited. Share warrant to bearer, £1. No. 68868. Capital £1,250,000. London, May 10, 1917. Bilingual English/French. Spassky mine, Kazakhstan.',
  }
);

// --- Row 453: Sechste Österreichische Kriegsanleihe (6th Austrian War Loan), 1,000 Kronen, 1917 ---
setDoc(453,
  'Sechste Österreichische Kriegsanleihe (Sixth Austrian War Loan): 5½% Staatsanleihe, 1,000 Kronen (Serie 247, No. 007886, Vienna, April 1, 1917)',
  'This printed bearer bond of 1,000 kronen at 5½% annual tax-free (steuerfreie) amortizable interest is one of the denominations of the Sixth Austrian War Loan (Sechste Österreichische Kriegsanleihe). Serie 247, No. 007,886. Vienna, April 1, 1917. Art Nouveau decorative design with flanking figures, cityscape oval, and Austrian imperial arms. This was the sixth of eight successive war loans issued by Austria-Hungary to finance World War I. By 1917 the financial situation was increasingly desperate, and these loans were heavily promoted through patriotic advertising campaigns. Along with the seventh and eighth loans, the sixth issue was subscribed largely by patriotic compulsion rather than investor confidence, and the collapse of the Habsburg Empire in 1918 rendered these bonds worthless.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'Austro-Hungarian Government',
    issueDate: '1917-04-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Sechste Österreichische Kriegsanleihe (6th Austrian War Loan). 5½% tax-free amortizable Staatsanleihe. 1,000 kronen. Serie 247, No. 007,886. Vienna, April 1, 1917. Art Nouveau design.',
  }
);

// --- Row 454: Achte Österreichische Kriegsanleihe (8th Austrian War Loan), 100 Kronen, 1918 ---
setDoc(454,
  'Achte Österreichische Kriegsanleihe (Eighth Austrian War Loan): 5½% Staatsanleihe, 100 Kronen (Serie 117, No. 128924, Vienna, June 1, 1918)',
  'This printed bearer bond of 100 kronen at 5½% annual tax-free amortizable interest is a small-denomination note of the Eighth Austrian War Loan (Achte Österreichische Kriegsanleihe). Serie 117, No. 128,924. Vienna, June 1, 1918. Art Nouveau decorative design with the Austrian imperial coat of arms flanked by allegorical figures. The eighth and final Austrian war loan was issued in June 1918, just five months before the armistice that ended World War I and the simultaneous collapse of the Austro-Hungarian Empire. By this point it was clear to most observers that Austria-Hungary was heading for defeat, making this among the most ill-fated of all sovereign bond issues in European history.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'Austro-Hungarian Government',
    issueDate: '1918-06-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Achte Österreichische Kriegsanleihe (8th Austrian War Loan). 5½% tax-free amortizable. 100 kronen. Serie 117, No. 128,924. Vienna, June 1, 1918. Art Nouveau design. Final Austrian war loan, issued five months before armistice.',
  }
);

// --- Row 455: k.k. Staatsschuldverschreibung (Austrian Imperial State Bond), 100 Gulden, 1868 ---
setDoc(455,
  'k.k. Staatsschuldverschreibung (Austrian Imperial State Bond): 100 Gulden Österreichische Währung (No. 207,407, Vienna, July 1, 1868)',
  'This printed Imperial-Royal (k.k.) State Debt Certificate (Staatsschuldverschreibung) represents 100 gulden (Österreichische Währung / Austrian Currency) in the 4½% transfunded Austrian state debt. No. 207,407. Vienna, July 1, 1868. Issued by the k.k. Direction der Staatschulden (Imperial-Royal Directorate of State Debts). The document certifies the bearer\'s participation in the consolidated state debt resulting from the debt arrangements of 1867 (following Austria\'s defeat in the Austro-Prussian War of 1866 and the Compromise establishing the Austro-Hungarian dual monarchy). The Austrian gulden (florin) was the standard currency until replaced by the krone in 1892. Bears an Antwerpen (Antwerp) stamp, indicating this bond was registered or deposited there, consistent with Dutch/Belgian investor holding.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'k.k. Direction der Staatschulden (Austrian Imperial-Royal Directorate of State Debts)',
    issueDate: '1868-07-01',
    currency: 'ATS',
    language: 'German',
    numberPages: 1,
    period: '19th Century',
    notes: 'k.k. Staatsschuldverschreibung (Austrian Imperial State Bond). 100 gulden (Österreichische Währung). No. 207,407. Vienna, July 1, 1868. 4½% transfunded state debt. k.k. Direction der Staatschulden. Antwerpen deposit stamp.',
  }
);

// --- Row 456: Two Massachusetts financial documents, 1785-1786 ---
setDoc(456,
  'Massachusetts Commonwealth Tax Certificate (£3, 1786) and State of Massachusetts-Bay Bill ($5, ca. 1785): Two Documents',
  'This image documents two early American financial instruments on a single sheet. The upper document is a Massachusetts Commonwealth Tax Certificate, No. 1582, issued by the Treasurer\'s Office in Boston, April 1, 1786. Signed by Alex Hodgdon, Treasurer. It certifies that £3.13s.0 is due to P. Seek or Bearer, receivable in payment of one-third of Tax No. 5 (granted by the General Court of Massachusetts in March 1786), equivalent in value to gold and silver—a noteworthy inflation-adjustment clause. The lower document is a State of Massachusetts-Bay Bill, No. 1541 (or similar), for Five Dollars ($5), issued pursuant to an Act of the Legislature of the State of Massachusetts Bay dated the Fifth Day of May, 1785. Payable December 31, 1786, with 5% annual interest, paid by the State of Massachusetts-Bay in Spanish milled dollars. Signed by A. Cranch. These two documents represent the emergency finance instruments of the early American republic in the financially troubled years after the Revolution, when many states issued paper money and short-term certificates to manage fiscal obligations.',
  {
    type: 'Certificate, Promissory Note',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Commonwealth of Massachusetts (Treasurer\'s Office)',
    issueDate: '1786-04-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Two documents: (1) Massachusetts Commonwealth Tax Certificate No. 1582, £3.13s.0, Boston, April 1, 1786. Signed by Alex Hodgdon, Treasurer. (2) State of Massachusetts-Bay Bill No. 1541(?), $5, ca. 1785. 5% interest; payable Dec 31, 1786. Signed by A. Cranch.',
  }
);

// --- Row 457: State of New York Canal Stock, $1,000, Albany, 1842 ---
setDoc(457,
  'State of New York Canal Stock: $1,000 at 3% Per Annum (Henry Whiting, Albany, May 31, 1842)',
  'This printed stock certificate from the New York State Comptroller\'s Office certifies that the People of the State of New York owe to Henry Whiting or his assigns the sum of $1,000 ("New York State Canal Stock") bearing interest at 3% per annum from May 31, 1842, payable quarterly on the first days of February, April, July, and October. Reimbursable at the pleasure of the State at any time after one year\'s advance notice. Albany, May 31, 1842. Signed by N.E. [?], Comptroller. Issued in pursuance of Chapter [?] of the Laws of New York, 1842. New York State canal stock was issued to fund the construction and expansion of the Erie Canal and related waterways—the most transformative public works project in American history, which opened the interior to commerce and made New York City the nation\'s leading commercial center.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of New York, Comptroller\'s Office',
    issueDate: '1842-05-31',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'State of New York Canal Stock. $1,000 at 3% p.a. Issued to Henry Whiting. Albany, May 31, 1842. Comptroller\'s Office. Quarterly interest; reimbursable at State\'s pleasure after 1 year notice. Erie Canal financing.',
  }
);

// --- Row 458: "Св. Георги" (Saint George) Textile Company, 10 shares, Sofia, 1929 ---
setDoc(458,
  '"Св. Георги" ("Sv. Guéorgu" / Saint George) Textile Company: 10 Shares at 100 Leva Each (No. 79881–79890, Sofia, 1929)',
  'This bilingual (Bulgarian/French) printed block certificate represents 10 shares at 100 leva each (total 1,000 leva) in the Акционерно дружество за текстилна индустрия и търговия с текстилни материали "Св. Георги" (Société Anonyme pour l\'Industrie Textile et le Commerce des Matériaux Textiles "Sv. Guéorgu" / Saint George Textile Company). Shares No. 79881–79890. Sofia, 1929. With Sofia chamber of commerce and fiscal stamps. The company was engaged in textile manufacturing and the textile materials trade in interwar Bulgaria, which was building its industrial base during this period. The bilingual Bulgarian/French format reflects the French legal forms commonly adopted by Bulgarian commercial companies in the early twentieth century.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'Bulgaria',
    issuingCountry: 'Bulgaria',
    creator: '"Св. Георги" (Saint George) Textile Company',
    issueDate: '1929-01-01',
    currency: 'BGN',
    language: 'Bulgarian, French',
    numberPages: 1,
    period: '20th Century',
    notes: '"Св. Георги" Textile Company (Société Anonyme pour l\'Industrie Textile). 10 shares at 100 leva each. Shares No. 79881–79890. Sofia, 1929. Bilingual Bulgarian/French.',
  }
);

// --- Row 459: Texas and German Emigration Company, Certificate of Stock No. 572, Houston, 1852 ---
setDoc(459,
  'Texas and German Emigration Company: Certificate of Stock No. 572 (Houston, June 15, 1852)',
  'This printed stock certificate, No. 572, Folio 92, certifies that the German Emigration Company is the proprietor of 2 shares (No. 2775 and 2776) at $100 each in the Texas and German Emigration Company, constituted by Indentures dated September 15, 1843. Given at the City of Houston, Texas, this fifteenth day of June, 1852. Issued by authority of Indenture and signed by Henry F. Fisher, Agent of the German Emigration Company. The Memorandum on the certificate describes the joint stock as consisting of lands and tenements in Texas (about 1,200,000 acres plus other Texas properties), all rights and claims under the colonization contract with the Republic/State of Texas, and all debts and claims in favor of the Company. The Texas and German Emigration Company was associated with the Adelsverein (Society for the Protection of German Immigrants in Texas), organized by Prince Carl of Solms-Braunfels and other German noblemen to colonize Texas. Its agents—especially Henry Francis Fisher—negotiated land grants with the Republic of Texas, though the colonization enterprise ultimately failed financially. This certificate and the Adelsverein priority bond (cf. row 424) are companion pieces from the same colonization venture.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Texas and German Emigration Company',
    issueDate: '1852-06-15',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Texas and German Emigration Company. Certificate of Stock No. 572, Folio 92. 2 shares (No. 2775–2776) at $100 each. Houston, June 15, 1852. Agent: Henry F. Fisher. German immigrant colonization of Texas (Adelsverein). Related to Adelsverein bond (row 424).',
  }
);

// --- Row 460: The Common Fund Company Limited, Scrip H170, London, 1869 ---
setDoc(460,
  'The Common Fund Company Limited: Scrip H170, Right to Call for 100 Shares at £20 Each (London, December 20, 1869)',
  'This engraved scrip certificate H170 of The Common Fund Company Limited certifies that the holder has the right to call for 100 shares of £20 each in the Capital Stock of the company, and to receive Certificates of Full Stock upon paying the par cost of £20 per share in one sum. Signed by R.W. Millz, Registrar. London, December 20, 1869. "Convertible into stock at par by all the Company\'s bankers in Europe and America." Offices in Washington, Berlin, Paris, and London. Incorporated under the Companies\' Acts of 1862–67. Features a decorative globe showing North and South America. The Common Fund Company was a transatlantic investment company of the late Victorian era, with an unusually broad international corporate presence (offices in four countries) and pan-European convertibility for its scrip. Scrip certificates (as opposed to full shares) were interim instruments representing the right to acquire fully paid shares upon payment of the subscription price.',
  {
    type: 'Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United Kingdom',
    creator: 'The Common Fund Company Limited',
    issueDate: '1869-12-20',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'The Common Fund Company Limited. Scrip H170, right to call for 100 shares at £20 each. London, December 20, 1869. Registrar: R.W. Millz. Offices: Washington, Berlin, Paris, London. Incorporated under Companies\' Acts 1862–67.',
  }
);

// --- Row 461: Commonwealth of Pennsylvania Land Warrant, 1792 ---
setDoc(461,
  'Commonwealth of Pennsylvania: Land Warrant for Survey, Allegheny/Conewango Purchase (Governor Thomas Mifflin, 1792)',
  'This printed and handwritten land warrant is issued by the Commonwealth of Pennsylvania, signed by Governor Thomas Mifflin, directing the Surveyor-General, Daniel Brodhead, Esquire, to survey a tract of land in the late purchase made from the Indians, East of the Allegheny River and Conewango Creek, for Robert A. Arres(?), who had paid the full purchase price to the Receiver-General of the Land-Office. The warrant references several statutes: An Act of April 7, 1792, for the purchase and regulation of unappropriated lands; An Act of October 3, 1788, for surveying and disposing of unappropriated lands; and An Act of April 19, 1792, for sale of vacant lands. The document instructs that four hundred [?] acres be surveyed. 1792. This land warrant represents a foundational document of American frontier real estate finance—the mechanism by which the Commonwealth of Pennsylvania allocated and monetized its western lands in the years immediately following independence.',
  {
    type: 'Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Commonwealth of Pennsylvania (Governor Thomas Mifflin)',
    issueDate: '1792-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Commonwealth of Pennsylvania land warrant. Signed by Governor Thomas Mifflin. For survey of land in the Allegheny River/Conewango Creek purchase (Act of April 7, 1792). Directed to Surveyor-General Daniel Brodhead. For Robert A. Arres(?). 1792.',
  }
);

// --- Row 462: The Egyptian Enterprise and Development Company, Deferred Share Warrant, Cairo, 1906 ---
setDoc(462,
  'The Egyptian Enterprise and Development Company: Deferred Share Warrant to Bearer No. 08527 (Cairo, March 5, 1906)',
  'This bilingual (English/French) printed deferred share warrant to bearer (Certificat de Part de Dividende au Porteur) No. 08527 entitles the holder to one deferred share in The Egyptian Enterprise and Development Company. Capital: £E. 160,000, increased to £E. 400,000 by resolution of the General Meeting of Shareholders on December 15, 1905, divided into 40,000 ordinary shares of £E. 10 each. Empowered by Khedival Decree of November 26, 1904. Cairo, March 5, 1906. Signed by two Administrators. The Egyptian Enterprise and Development Company was an Edwardian-era commercial enterprise operating in Egypt under Khedivial charter, engaged in development and business activities in the country during the period of British occupation (1882–1952). Deferred shares received dividends only after ordinary shares had received their fixed dividend, making them highly leveraged profit-participation instruments.',
  {
    type: 'Certificate',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'The Egyptian Enterprise and Development Company',
    issueDate: '1906-03-05',
    currency: 'EGP',
    language: 'English, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'The Egyptian Enterprise and Development Company. Deferred share warrant to bearer No. 08527. Capital £E. 400,000. Khedivial Decree November 26, 1904. Cairo, March 5, 1906. Bilingual English/French.',
  }
);

// --- Row 463: The New Four per Cent. Annuities (Bank of England), transfer receipt, 1822 ---
setDoc(463,
  'The New Four per Cent. Annuities (Bank of England): Stock Transfer Receipt for £6 6s. 2d. (London, August 2, 1822)',
  'This printed and handwritten transfer receipt records the transfer of £6 6s. 2d. (six pounds, six shillings, two pence) interest (or share) in the Capital Stock of The New Four per Cent. Annuities—the British Government Consols restructured from Five Per Cent. Annuities under an Act of Parliament of the 3rd Year of His Majesty King George IV, entitled "An Act for transferring several Annuities of Five Pounds per Cent. into Annuities of Four Pounds per Cent." Transferable at the Bank of England. Consideration paid: £6 5s. 7d. Transferred to Philip Davis, Abigail Davis, and Jane Catherine Davis. Received August 2, 1822. Witnessed by J. Robinson (1085) and H.J. Cosper. "Account Closed 25th November 1836." Transfer Days: Tuesday, Wednesday, Thursday, Friday (Holidays excepted). Dividends due January 5 and July 5 annually. This document illustrates the secondary market in British government perpetual annuities at the Bank of England—the foundation of the British national debt market that developed from the late seventeenth century and became the model for sovereign bond markets worldwide.',
  {
    type: 'Receipt',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'Bank of England (New Four per Cent. Annuities)',
    issueDate: '1822-08-02',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'The New Four per Cent. Annuities (Bank of England). Stock transfer receipt for £6 6s. 2d. August 2, 1822. Transfer to Philip Davis, Abigail Davis & Jane Catherine Davis. Witnesses: J. Robinson (1085), H.J. Cosper. Account closed November 25, 1836. Act of 3 Geo. IV.',
  }
);

// --- Row 464: Thomas C. Jenkins Wholesale Grocery Invoice, Pittsburgh, 1902 ---
setDoc(464,
  'Thomas C. Jenkins, Wholesale Dealer in Flour & Groceries: Invoice to Offutt & Son (Pittsburgh, Pennsylvania, December 19, 1902)',
  'This printed commercial invoice is from Thomas C. Jenkins, described as "The Largest Flour and Grocery House in the World," located at Nos. 508–522 Penn Avenue and Nos. 509–523 Liberty Street, Pittsburgh, Pennsylvania, at the Foot of Fifth Avenue near the Pittsburgh Market. The invoice is dated December 19, 1902, and addressed to Offutt & Son. Items listed include tobacco products: "J.S. Toba[cco] 63" (36 units at $2.68) and "Island Navy [tobacco] 56" (36 units at $2.016). Note: "No Goods Sold Families. No Retail Store or any Connection with any Retail House in any Shape or Form. All Electric Line Cars run to or pass near the New Building." SHIPPED DIRECT stamp. Jenkins\' establishment handled flour, groceries, teas, roast coffees, and tobaccos on a wholesale basis, and was a major commercial presence in late Gilded Age Pittsburgh.',
  {
    type: 'Invoice',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Thomas C. Jenkins (Wholesale Dealer in Flour & Groceries)',
    issueDate: '1902-12-19',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Thomas C. Jenkins, Pittsburgh, PA. Wholesale Flour & Groceries. Invoice to Offutt & Son, December 19, 1902. Items: tobacco products. Nos. 508–522 Penn Ave. / 509–523 Liberty St. "The Largest Flour and Grocery House in the World."',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 421–464 (44 documents, batch14).');
