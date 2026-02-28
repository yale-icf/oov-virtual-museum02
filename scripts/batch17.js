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

// --- Row 505: United States of America, Bill of Exchange No. 43, $120 / 600 Livres Tournois, October 1, 1780 ---
setDoc(505,
  'United States of America: Bill of Exchange No. 43, $120 (600 Livres Tournois) for Interest on Loan Office Certificates (October 1, 1780)',
  'A printed and handwritten bill of exchange No. 43, issued by the United States of America on October 1, 1780. At Thirty Days Sight, payable to Jonathan D. Sergeant (or Order) in Six Hundred Livres Tournois, equivalent to One Hundred and Twenty Dollars at five Livres Tournois per Dollar, for Interest due on Money borrowed by the United States. Countersigned by the Commissioner of the Continental Loan Office in the State of Pennsylvania. Signed by H. Atkinson, Treasurer of Loans. This is one of a series of bills of exchange issued by the Continental Congress to pay interest on Loan Office Certificates by drawing on French credit, reflecting the United States\' dependence on French currency networks during the Revolutionary War.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Pennsylvania',
    issueDate: '1780-10-01',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 43, $120 / 600 Livres Tournois. October 1, 1780. Payable to Jonathan D. Sergeant. Interest on Loan Office Certificates. Pennsylvania Loan Office; signed by H. Atkinson, Treasurer of Loans. Continental Revolutionary War finance.',
  }
);

// --- Row 506: Potosian Land Grant, Class C, 400 Acres, London, April 1825 ---
setDoc(506,
  'Potosian Land Grant: Class C, 400 Acres / Concesión de 400 Acres (London, April 1825)',
  'A large bilingual English/Spanish land grant certificate (Class C) of 400 Acres in the Potosí region, dated London, April 1825. The document is headed "POTOSIAN LAND GRANT" and includes the full grant text in English as well as a Spanish translation (Concesión de 400 Acres), with signatures from London confirming the Exchange and Redemption of the outstanding claims. The Potosian Land Grant scheme was associated with early 19th-century British speculative investment in Bolivian mining and land ventures following South American independence, when British investors rushed to capitalize on newly opened Latin American markets. Potosí, historically famous for its vast silver mines since the 16th century, attracted significant British capital in the 1820s as part of the broader Latin American investment mania.',
  {
    type: 'Land Grant',
    subjectCountry: 'Bolivia',
    issuingCountry: 'United Kingdom',
    creator: 'Potosian Land Grant Company (London)',
    issueDate: '1825-04-01',
    currency: 'GBP',
    language: 'English, Spanish',
    numberPages: 1,
    period: '19th Century',
    notes: 'Potosian Land Grant, Class C, 400 Acres / Concesión de 400 Acres. London, April 1825. Bilingual English/Spanish. British speculative investment in Bolivian (Potosí) mining/land post-independence. Large format with ornate header and two-column bilingual text.',
  }
);

// --- Row 507: Prys-Courant der Effecten, Amsterdam Securities Price List, No. 39, 1822 ---
setDoc(507,
  'Prys-Courant der Effecten: Amsterdam Securities and Exchange Rate Price List, No. 39 (Amsterdam, 1822)',
  'A printed weekly securities and exchange rate price list (Prys-Courant der Effecten) from Amsterdam, Thursday, 1822, No. 39. The broadsheet lists market prices for a wide range of securities: Dutch government bonds (Vereenigde Nederlanden), various Dutch negotiaties (investment pools), and a broad array of foreign government bonds and securities including Spanish, French, Austrian, Russian, and other European and American obligations. The lower portion gives Wissels en Specie Cours (exchange rates for bills of exchange and coin prices) for London, Paris, Hamburg, Frankfurt, Berlin, Vienna, and other centers. The Prys-Courant der Effecten was Amsterdam\'s authoritative weekly securities price publication, a primary record of early 19th-century European financial market activity and international capital flows.',
  {
    type: 'Price List',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Amsterdam Stock Exchange',
    issueDate: '1822-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Amsterdam Prys-Courant der Effecten, No. 39, 1822. Lists prices for Dutch government bonds, negotiaties, Spanish, French, Austrian, Russian, and other foreign securities. Also: Wissels en Specie Cours (exchange rates and coin prices) for major European financial centers.',
  }
);

// --- Row 508: Regeering van de Chineesche Republiek, 8% Schatkistbiljet, Lung-Tsing-U-Hai Railway, 1923, f.1000, No. 11639 ---
setDoc(508,
  'Regeering van de Chineesche Republiek: 8% Schatkistbiljet, Lung-Tsing-U-Hai Spoorweg (1923), f. 1000 (No. 11639)',
  'An 8% treasury bill (Schatkistbiljet aan Toonder) No. 11639, denomination f. 1,000 Dutch guilders, issued by the Government of the Chinese Republic (Regeering van de Chineesche Republiek) for the Lung-Tsing-U-Hai Railway (Lung-Tsing-U-Hai-Spoorweg), total issue f. 16,667,000 divided into 16,667 notes of f. 1,000 each. The bill features an ornate multicolor printed design with a Chinese bridge or architectural vignette and decorative border. Signed in Chinese characters and Dutch by Chinese authorities. The Lung-Tsing-U-Hai (Longhai) Railway was a major east-west Chinese railway linking the coast to the interior. Dutch guilder-denominated Chinese railway bonds were a significant form of European financing for Chinese infrastructure development during the early Republic period.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'Netherlands',
    creator: 'Government of the Chinese Republic',
    issueDate: '1923-01-01',
    currency: 'NLG',
    language: 'Dutch, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: '8% Schatkistbiljet, Lung-Tsing-U-Hai Spoorweg (Longhai Railway), 1923. No. 11639, f. 1,000 Dutch guilders. Total issue f. 16,667,000 / 16,667 notes. Regeering van de Chineesche Republiek. Ornate multicolor certificate. Dutch-guilder denominated Chinese railway finance.',
  }
);

// --- Row 509: Kornelis van den Helm Boddaert, Middelburg Plantation Bond No. 99, Essequibo and Demerara, January 1760 ---
setDoc(509,
  'Kornelis van den Helm Boddaert: Middelburg Plantation Bond No. 99, Essequibo and Demerara (January 1760)',
  'A handwritten mortgage bond/plantation loan document No. 99 from the Register of Middelburg (Zeeland), signed by Kornelis van den Helm Boddaert as Directeur and qualified secretary for the money-lenders (geld-Opschieters) of various Plantations in the Colony of Essequibo and Demerara (present-day Guyana), acting according to the General Plan of Hypotheken (mortgages) secured by the colonial plantations. Dated January 1760. Capital of Five Hundred Livres at 5% per cent per annum, interest from January 1, 1760. Authenticated with the Compagnie seal. Multiple endorsements and subsequent annotations in later hands (1772, 1793, 1817). This document exemplifies the Dutch colonial plantation mortgage system (plantageleningen) of Essequibo and Demerara, which bundled West Indian plantation mortgages for sale to Dutch investors — among the earliest securitization schemes in financial history.',
  {
    type: 'Plantation Bond',
    subjectCountry: 'Guyana',
    issuingCountry: 'Netherlands',
    creator: 'Kornelis van den Helm Boddaert; Middelburg Administration',
    issueDate: '1760-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Middelburg plantation bond No. 99. Kornelis van den Helm Boddaert, Director/Secretary. Plantations in Essequibo and Demerara (Guyana). Capital ~500 Livres at 5% p.a. January 1760. Dutch plantageleningen system. Endorsements dated 1772, 1793, 1817.',
  }
);

// --- Row 510: French Royal Life Annuity (Rentes Viagères), Édit Juillet 1723, No. 21839, Simonne Jenaille ---
setDoc(510,
  'Rentes Viagères au Denier Vingt-cinq sur les Quatre Millions de Livres (Édit de Juillet 1723): Life Annuity No. 21839 for Simonne Jenaille (Paris, 1723)',
  'A printed and handwritten royal life annuity contract (Rente Viagère au Denier Vingt-cinq) No. 21839, issued by Jean Antoine Paris, Conseiller du Roy en Exercice en cette Ville de Paris, for Simonne Jenaille (and family). The annuity is drawn on the Four Millions of Livres raised under the Édit du Mois de Juillet 1723 on the Tailles and other Impositions. The annual payment is twenty-three livres and approximately seventy-four sous. The rente viagère was tied to the life of the nominee and ceased upon their death. Issued by the Quinzième du Garde du Trésor Royal, Paris, 1723. Signed by M. Lacromont [?]. An important instrument of French royal finance during the Regency and early Louis XV period, when life annuities were a principal mechanism for raising state revenue.',
  {
    type: 'Life Annuity',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Jean Antoine Paris, Conseiller du Roy',
    issueDate: '1723-07-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century',
    notes: 'Rente Viagère au Denier Vingt-cinq, No. 21839. Quatre Millions de Livres, Édit de Juillet 1723, Tailles & autres Impositions. Nominee: Simonne Jenaille. Jean Antoine Paris, Conseiller du Roy. Paris, 1723. Annual payment ~23 livres 74 sous. French royal life annuity.',
  }
);

// --- Row 511: French Royal Life Annuity (Rentes Viagères Sur une Tête), Édit Novembre 1738, Comtesse de Saint-Hermaurice, Paris, July 26, 1739 ---
setDoc(511,
  'Rentes Viagères Sur une Tête (Édit de Novembre 1738): Life Annuity, Première Classe, for the Comtesse de Saint-Hermaurice (Paris, July 26, 1739)',
  'A printed and handwritten royal life annuity contract (Rente Viagère Sur une Tête) for the Comtesse de Saint-Hermaurice (Madame La Comtesse de Saint-Hermaurice), drawn under the Édit de Novembre 1738, Première Classe. Issued July 26, 1739, Paris, before Jean-Baptiste-Élie Canus de Pontcaillé, seigneur de Vuarnet, Sirogu, Briey, and others, Notaires, Gardes-note et Garde du Seau de Sa Majesté au Châtelet de Paris, and other Parisian notaries. The annuity is for acquéreurs of annuities from the Marchands et Echevins (merchant aldermen) of Paris, with four thousand livres of capital each, at a fixed annual rente. The life annuity runs for the lifetime of the nominated individual and provides for specific conditions on re-registration and transfer. An important document from French royal life annuity finance during the reign of Louis XV.',
  {
    type: 'Life Annuity',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Notaires du Châtelet de Paris; Jean-Baptiste-Élie Canus de Pontcaillé',
    issueDate: '1739-07-26',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century',
    notes: 'Rente Viagère Sur une Tête, Édit de Novembre 1738, Première Classe. Nominee: Comtesse de Saint-Hermaurice. Issued July 26, 1739, Paris. Notaires du Châtelet. Marchands et Echevins financing (4,000 livres capital each). French royal life annuity under Louis XV.',
  }
);

// --- Row 512: Republic of New Granada, Deferred Bond, Letter A, No. 7376, £100 Sterling, London, January 1, 1845 ---
setDoc(512,
  'Republic of New Granada: Deferred Bond, Letter A, No. 7376, £100 Sterling (London, January 1, 1845)',
  'A bilingual English/Spanish deferred bearer bond, Letter A, No. 7376, for £100 Sterling, issued by the Republic of New Granada (predecessor to Colombia), dated London, January 1, 1845. Signed by Juan Climaco Ordóñez as Chargé d\'Affaires of the Republic of New Granada at the Court of London, and by Powles, Illingworth, Willson & Co. as agents and commissioners. The bond acknowledges £100 of the Foreign Debt of New Granada. Interest is deferred at Four Per Cent per annum; dividends to be paid half-yearly from December 1 each year at the counting house of Messrs. Runneys, Barings & Co. in London, upon delivery of Warrants or Coupons. The Government of New Granada engages to deliver additional bonds as payment for accumulated arrears of interest. Signed in Bogotá, August 15, 1843. The New Granadan external debt, originating in the independence period, underwent multiple restructurings throughout the 19th century.',
  {
    type: 'Bond',
    subjectCountry: 'Colombia',
    issuingCountry: 'United Kingdom',
    creator: 'Republic of New Granada; Powles, Illingworth, Willson & Co.',
    issueDate: '1845-01-01',
    currency: 'GBP',
    language: 'English, Spanish',
    numberPages: 1,
    period: '19th Century',
    notes: 'Republic of New Granada Deferred Bond, Letter A, No. 7376, £100 Sterling. London, January 1, 1845. Chargé d\'Affaires: Juan Climaco Ordóñez. Agents: Powles, Illingworth, Willson & Co. 4% interest deferred; dividends at Runneys, Barings & Co. Signed Bogotá August 15, 1843. Bilingual English/Spanish.',
  }
);

// --- Row 513: Rjasan Uralsk Spoorweg Maatschappij, 4½% Certificaat A No. 1977, f.1200, Amsterdam, July 14, 1903 ---
setDoc(513,
  'Rjasan Uralsk Spoorweg Maatschappij: 4½% Certificaat A No. 1977, f. 1,200 (Amsterdam, July 14, 1903)',
  'A 4½% Certificaat A No. 1977 of the Rjasan Uralsk Spoorweg Maatschappij (Ryazan-Ural Railway Company), denomination f. 1,200 Dutch guilders (Twaalf Honderd Gulden). The Ryazan-Koslov (Rjasan Koslov) Railway Company was established in Moscow under the Imperial Ukase of January 11, 1897. Issued by the Algemeene Trust Maatschappij (General Trust Company) in Amsterdam, July 14, 1903. This consolidated certificate replaces twelve prior Certificaten A No. 2722, each of f. 100. All obligations, rights, and conditions of the Ryazan-Ural Railway Maatschappij company bond remain attached, and the certificate is subject to Article 6 of the relevant terms. The Ryazan-Ural Railway linked the Ryazan region to the Ural industrial and agricultural zone, playing a key role in Russian grain export logistics. Dutch certificates for Russian railway bonds were a staple of Amsterdam investment portfolios in the early 20th century.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Rjasan Uralsk Spoorweg Maatschappij; Algemeene Trust Maatschappij',
    issueDate: '1903-07-14',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: '4½% Certificaat A No. 1977, f. 1,200. Rjasan Uralsk Spoorweg Maatschappij (Ryazan-Ural Railway). Algemeene Trust Maatschappij, Amsterdam, July 14, 1903. Replaces 12× Certificaat A No. 2722 at f. 100 each. Imperial Ukase January 11, 1897.',
  }
);

// --- Row 514: The Rock Island Company, Dutch Certificaat, $1,000 Common Stock, No. 52108, Amsterdam ---
setDoc(514,
  'The Rock Island Company: Dutch Certificaat, $1,000 Common Stock (No. 52108, Amsterdam)',
  'A Dutch Certificaat No. 52108 of $1,000 Common Stock of The Rock Island Company (Chicago, Rock Island and Pacific Railway), administered by the Amsterdam Administratie-Kantoor of Broes en Goosman, Ten Have en van Essen, Jarman en Zoonen, and Plantenga. The holder is entitled to the benefits of one share of $1,000 Common Stock in The Rock Island Company under the conditions of the Dutch administration agreement. Printed in pink/red on cream paper. Last dividend date noted. The Rock Island Company was one of the major American Midwestern railroads, operating routes from Chicago westward. Amsterdam-listed certificates for American railroad shares were widely held by Dutch retail investors seeking exposure to the booming American railway sector in the late 19th and early 20th centuries.',
  {
    type: 'Stock',
    subjectCountry: 'United States',
    issuingCountry: 'Netherlands',
    creator: 'The Rock Island Company; Broes en Goosman; Ten Have en van Essen; Jarman en Zoonen; Plantenga (Amsterdam)',
    issueDate: '1900-01-01',
    currency: 'USD',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Dutch Certificaat No. 52108, $1,000 Common Stock, The Rock Island Company (Chicago, Rock Island and Pacific Railway). Administered by Broes & Goosman, Ten Have & van Essen, Jarman & Zoonen, Plantenga, Amsterdam. Pink-printed certificate.',
  }
);

// --- Row 515: Vereeniging tot Bevordering van 's Lands Weerbaarheid, Loterij-Geldleening f.1,000,000, Series 7734 No. 38, Rotterdam, May 1877 ---
setDoc(515,
  'Vereeniging tot Bevordering van \'s Lands Weerbaarheid: Loterij-Geldleening f. 1,000,000, Series 7734 No. 38 (Rotterdam, May 1877)',
  'A lottery bond (Loterij-Geldleening) of the Dutch Association for the Promotion of National Defense (Vereeniging tot Bevordering van \'s Lands Weerbaarheid), Series 7734, No. Acht en Dertig (38). Total loan: f. 1,000,000, divided into 400,000 shares at f. 2.50 each, approved by Ministerial Resolution of July 8, 1878, No. 28. Issued in Rotterdam, May 1877. On July 31 and January 15 annually, 12,000 shares are drawn for prizes, with top prizes of f. 10,000, f. 5,000, f. 2,000, f. 1,000, f. 500, f. 200, f. 100, f. 75, f. 50, f. 20, and smaller amounts. The certificate lists the full prize table and redemption conditions. The Vereeniging tot Bevordering van \'s Lands Weerbaarheid (Association for National Defense Readiness) was a Dutch patriotic association that used lottery bonds to raise funds for Dutch military preparedness in the post-1870 period.',
  {
    type: 'Lottery Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Vereeniging tot Bevordering van \'s Lands Weerbaarheid',
    issueDate: '1877-05-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Loterij-Geldleening, f. 1,000,000 / 400,000 shares at f. 2.50. Series 7734, No. 38. Rotterdam, May 1877. Draws: July 31 and January 15 annually; 12,000 shares. Top prize f. 10,000. Vereeniging tot Bevordering van \'s Lands Weerbaarheid (national defense association).',
  }
);

// --- Row 516: Imperial Russian State Debt Commission, 5% Perpetual Income Certificate, 960 Roubles, 1822 ---
setDoc(516,
  'Imperial Russian State Debt Commission: 5% Perpetual Income Certificate (Непрерывный Доход), 960 Roubles (1822)',
  'A printed 5% perpetual income certificate (Свидетельство о непрерывном доходе по 5 на сто / "Nepreryvny Dokhod") issued by the Imperial Russian State Commission for the Repayment of Debts (Государственная Комиссия погашения долгов). Denomination: 960 Roubles, registered in the State Debt Book (Государственная долговая книга) on March 1, 1822. The Imperial Russian double-headed eagle is displayed prominently at top. The certificate entitles the holder to perpetual income at 5% per annum, payable semi-annually from the State Treasury. Conditions regarding transfer, inheritance, and interest payment are printed below. The "Nepreryvny Dokhod" (perpetual income) certificates were the Russian equivalent of Western government consols, representing a permanent claim on state revenues, and were among the principal instruments of 19th-century Russian public finance.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian State Commission for the Repayment of Debts',
    issueDate: '1822-03-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Russian 5% Perpetual Income Certificate (Непрерывный Доход по 5 на сто), 960 Roubles. Registered State Debt Book, March 1, 1822. Imperial double-headed eagle. Государственная Комиссия погашения долгов. Perpetual 5% income, semi-annual payments. Russian equivalent of consols.',
  }
);

// --- Row 517: Russian Public Debt, Inscription au Grand Livre, Commission Impériale d'amortissement, 1876, Première Série ---
setDoc(517,
  'Inscription au Grand Livre de la Dette Publique de Russie: Commission Impériale d\'Amortissement, Première Série (1876)',
  'A French-language certificate of inscription in the Grand Livre de la dette publique de Russie (Grand Ledger of Russian Public Debt), administered by the Commission Impériale d\'amortissement (Imperial Amortization Commission), 1876, Première Série (First Series). Inscribed for a capital of [?] Roubles in silver, in the names of Ketwich & Voombergh and Mme. W. Borski in Amsterdam, with an annual income (revenu annuel) of 500 Roubles. Payable on the 1st and 15th of March and the 1st and 15th of September (four times yearly). Signed by the Director of the Commission. Includes an Extrait du Règlement de la Commission, Chap. II, describing the conditions of the inscription. This certificate documents Dutch investment in Russian sovereign debt via the Commission Impériale d\'amortissement\'s registration system, a key channel for foreign capital into Russia.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Commission Impériale d\'amortissement de la dette publique de Russie',
    issueDate: '1876-01-01',
    currency: 'RUB',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Inscription au Grand Livre, dette publique de Russie. Commission Impériale d\'amortissement. 1876, Première Série. Capital [?] Roubles silver; annual income 500 Roubles. Amsterdam: Ketwich & Voombergh; Mme. W. Borski. Payable March 1/15 and September 1/15. French-language certificate.',
  }
);

// --- Row 518: Russian State Debt Commission, Perpetual Capital Certificate, 300 Roubles, 1839 ---
setDoc(518,
  'Билет Государственной Комиссии Погашения Долгов: Perpetual Capital Certificate (Капитал Безсрочный), 300 Roubles (1839)',
  'A printed certificate (Билет) of the Russian State Commission for the Repayment of Debts (Государственная Комиссия погашения долгов), issued 1839, for a Perpetual Capital (Капитал Безсрочный) of 300 Roubles. The Imperial Russian double-headed eagle is prominently displayed. The certificate confirms the holder\'s perpetual deposit in the State Debt Book (Государственная Долговая Книга), with income at 6% per annum, paid semi-annually on January 15 and July 15 (по Шести на Сто). Signed by the Administrator and Bookkeeper of the Commission. An important early 19th-century Russian state finance instrument representing a perpetual income claim on the Russian Treasury, functionally equivalent to Western consols or rentes perpétuelles.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Russian State Commission for the Repayment of Debts',
    issueDate: '1839-01-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Билет Государственной Комиссии погашения долгов. Perpetual Capital (Капитал Безсрочный), 300 Roubles. 1839. Income 6% p.a., paid January 15 and July 15. Imperial double-headed eagle. Signed by Administrator and Bookkeeper.',
  }
);

// --- Row 519: Russian State Debt Commission, Perpetual Capital Certificate, 50 Roubles in Silver, 1854 ---
setDoc(519,
  'Билет Государственной Комиссии Погашения Долгов: Perpetual Capital Certificate (Капитал Безсрочный), 50 Roubles in Silver (1854)',
  'A printed certificate (Билет) of the Russian State Commission for the Repayment of Debts, issued 1854, for a Perpetual Capital (Капитал Безсрочный) of 50 Roubles in Russian silver currency (Российским серебряным монетою). The Imperial Russian double-headed eagle is displayed at top. Income at 6% per annum (по Шести на Сто), payable semi-annually on January 15 and July 15 (по 15 Генваря и по 15-го Июля). Signed by the Administrator and Bookkeeper of the Commission. The silver denomination reflects Russia\'s monetary reform of 1839–1843 (the Kankrin Reform), which stabilized the ruble on a silver standard. This certificate is from the same series as the 1839 certificate (goetzmann0518) but represents a later issuance under Tsar Nicholas I, with the specific designation of silver currency.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Russian State Commission for the Repayment of Debts',
    issueDate: '1854-01-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Билет Государственной Комиссии погашения долгов. Perpetual Capital (Капитал Безсрочный), 50 Roubles in Russian silver (серебряным монетою). 1854. Income 6% p.a., paid January 15 and July 15. Imperial double-headed eagle. Post-Kankrin monetary reform denomination.',
  }
);

// --- Row 520: Hope en Comp. et al., 6% Certificaat / Certificat d'Inscription Russe, in Zilver / en Argent, 500 Roubles, No. 670, Amsterdam, November 19, 1857 ---
setDoc(520,
  'Hope en Comp., Ketwich en Voombergh, en Wed. W. Borski: 6% Certificaat / Certificat d\'Inscription Russe, Russische Fondsen in Zilver, 500 Roubles (No. 670, Amsterdam, November 19, 1857)',
  'A bilingual Dutch/French certificate No. 670 (R° 500 Roubles) for 6% Russian Funds in Silver (Russische Fondsen in Zilver / Inscription Russe en Argent), issued in Amsterdam on November 19, 1857, by Hope en Comp., Ketwich en Voombergh, and Weduwe W. Borski. The certificate represents a 500-ruble inscription in the Russian State Debt Book (Grand Livre de St. Pétersbourg / Groot Boek te St. Petersburg), held by the Bureau d\'Administration established in Amsterdam under direction of the issuing firms. Interest at 6% per annum, payable for the account and risk of the holder on Coupons (Art. 6 of the published notice). The holder may at any time reclaim the original inscription against return of the certificate and unused coupons. Includes ten half-yearly coupon tickets. A standard Amsterdam intermediary instrument for Dutch retail investment in Russian government silver bonds.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Hope en Comp.; Ketwich en Voombergh; Weduwe W. Borski',
    issueDate: '1857-11-19',
    currency: 'RUB',
    language: 'Dutch, French',
    numberPages: 1,
    period: '19th Century',
    notes: '6% Certificaat/Certificat d\'Inscription Russe, Russische Fondsen in Zilver/en Argent. No. 670, R° 500 Roubles. Amsterdam, November 19, 1857. Hope en Comp.; Ketwich en Voombergh; Wed. W. Borski. Interest 6% p.a. Ten half-yearly coupon tickets.',
  }
);

// --- Row 521: Hope en Comp. et al., 6% Certificaat / Certificat d'Inscription Russe, in Bank-Assignatien, 1000 Roubles, No. 4031, Amsterdam, 1822 ---
setDoc(521,
  'Hope en Comp., Ketwich en Voombergh, Wed. W. Borski: 6% Certificaat / Certificat d\'Inscription Russe, Russische Fondsen in Bank-Assignatien, 1000 Roubles (No. 4031, Amsterdam, 1822)',
  'A bilingual Dutch/French certificate No. 4031 (R° 1,000 Roubles) for 6% Russian Funds in Bank Assignations (Russische Fondsen in Bank-Assignatien / Inscription Russe en Assignations de Banque), issued in Amsterdam on August 16, 1822, by Hope en Comp., Ketwich en Voombergh, and Weduwe W. Borski. The certificate represents a 1,000-ruble inscription in the Russian State Debt Book in bank assignations (Russian paper currency of the era). Interest at 6% per annum. Includes ten half-yearly coupon tickets (deliverable from July 1, 1829). Signed by Ketwich Voombergh and partner; witnessed by a notary. The denomination in Bank Assignations (paper rubles) predates Russia\'s 1839 Kankrin Reform, which established a silver standard. An earlier example in the same series as No. 670 (goetzmann0520), illustrating the evolution of Dutch-intermediated Russian debt instruments.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Hope en Comp.; Ketwich en Voombergh; Weduwe W. Borski',
    issueDate: '1822-08-16',
    currency: 'RUB',
    language: 'Dutch, French',
    numberPages: 1,
    period: '19th Century',
    notes: '6% Certificaat/Certificat d\'Inscription Russe, in Bank-Assignatien/Assignations de Banque. No. 4031, R° 1,000 Roubles. Amsterdam, August 16, 1822. Hope en Comp.; Ketwich en Voombergh; Wed. W. Borski. Ten half-yearly coupons (from July 1, 1829). Denominated in Bank Assignations (pre-Kankrin Reform Russian paper currency).',
  }
);

// --- Row 522: Sillem, Benecke & Comp. and H.J. Stresow, Certificaat No. 13549, 500 Roubles 5% Russian Funds in Metalick, Hamburg, March 13, 1821 ---
setDoc(522,
  'Sillem, Benecke & Comp. and H.J. Stresow: Certificaat No. 13549, 500 Roubles 5% Russische Fondsen betaalbaar in Metalick (Hamburg, March 13, 1821)',
  'A Dutch certificaat No. 13549 issued in Hamburg on March 13, 1821, for 500 Roubles of 5% Russian Funds payable in metal (Metalick / specie). Administered by the Heeren Sillem, Benecke & Comp. and Heer H.J. Stresow, who established an administrative bureau in Hamburg for Russian State Inscriptions (Russische Inschrijvingen) registered in the Grand Ledger in the name of the Hamburg administration. Interest at 5% per annum, received for the account and risk of the holder, paid on Coupons per Art. 5 of the published notice. The holder may at any time reclaim the original inscription of 500 Roubles against return of the certificate and unused coupons. Signed by Sillem Benecke & Co. and H.J. Stren [?]. Notarized by N. Dongkants. Side stamp: Kantoor van Administratie. Illustrates the Hamburg market\'s significant role in distributing Russian government bonds to northern European investors in the early 19th century.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Germany',
    creator: 'Sillem, Benecke & Comp.; H.J. Stresow',
    issueDate: '1821-03-13',
    currency: 'RUB',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Certificaat No. 13549, 500 Roubles 5% Russian Funds payable in Metalick (specie). Sillem, Benecke & Comp. and H.J. Stresow, Hamburg. Grand Ledger inscription. Hamburg, March 13, 1821. Notary: N. Dongkants. Hamburg-issued Dutch certificate for Russian government debt.',
  }
);

// --- Row 523: Volksstaat Hessen, Schuldverschreibung, 6% Braunkohle-Roggen-Anleihe von 1923, Reihe 37 No. 20002, Darmstadt, April 5, 1923 ---
setDoc(523,
  'Volksstaat Hessen: Schuldverschreibung, Eine Einheit der 6% Braunkohle-Roggen-Anleihe von 1923 (Reihe 37, No. 20002, Darmstadt, April 5, 1923)',
  'A Schuldverschreibung (bond certificate) of the Volksstaat Hessen (People\'s State of Hesse), one unit (Eine Einheit) of the 6% Braunkohle-Roggen-Anleihe (Brown Coal-Rye Loan) of 1923. Reihe (Series) 37, No. 20002. Issued in Darmstadt, April 5, 1923, by the Hessische Staatsschuldenverwaltung (Hessian State Debt Administration). The bond carries 6% annual interest, payable on May 1 and November 1. The unit\'s monetary value (Geldwert) is indexed to the market price of brown coal and rye — a commodity-linked bond designed to preserve real value against the hyperinflationary collapse of the Reichsmark then underway. The loan runs from May 1, 1923 to April 30, 1924, with early redemption possible from May 1, 1928 at 2% annually. Attached are 30 Zinscheine (interest coupons) and 1 Erneuerungsschein (renewal coupon). The commodity-indexed "Sachwertanleihe" (real-value bond) was a widespread Weimar-era innovation.',
  {
    type: 'Bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Hessische Staatsschuldenverwaltung (Volksstaat Hessen)',
    issueDate: '1923-04-05',
    currency: 'German Marks',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Schuldverschreibung, Volksstaat Hessen. Eine Einheit, 6% Braunkohle-Roggen-Anleihe 1923. Reihe 37, No. 20002. Darmstadt, April 5, 1923. Value indexed to brown coal and rye prices. 30 interest + 1 renewal coupon. Weimar hyperinflation commodity-indexed bond (Sachwertanleihe).',
  }
);

// --- Row 524: Conversion of Spanish Debt, Passive Stock Certificate No. 457, £14 3s 4d, London, May 4, 1835 ---
setDoc(524,
  'Conversion of Spanish Debt: Passive Stock Certificate No. 457, £14 3s 4d (London, May 4, 1835)',
  'A certificate No. 457 for Fourteen Pounds Three Shillings and Four Pence (£14.3.4) of Spanish Five Per Cent Passive Stock, issued as part of the Conversion of Spanish Debt in London, May 4, 1835. The bearer is entitled to this sum of Passive Stock, and upon presentation of sufficient certificates forming together any of the listed standard amounts (£42.10, £85, £170, £255, £510, £1,020), the undersigned engages to deliver in exchange a Bond of the Spanish Government. Signed by [?] Ricardo. The Spanish Passive Stock represented capitalized arrears of interest on Spanish government bonds, converted into a new class of deferred-interest securities as part of the Martínez de la Rosa government\'s 1834–35 restructuring of Spain\'s troubled external debt. These fractional certificates circulated on the London market as interim claims pending conversion into full bonds.',
  {
    type: 'Bond',
    subjectCountry: 'Spain',
    issuingCountry: 'United Kingdom',
    creator: 'Spanish Government (London agent: Ricardo)',
    issueDate: '1835-05-04',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Conversion of Spanish Debt, Passive Stock Certificate No. 457, £14 3s 4d. Spanish 5% Passive Stock. Bearer exchanges certificate at standard amounts: £42.10, £85, £170, £255, £510, £1,020 for Spanish government bond. Signed by [Ricardo]. London, May 4, 1835. 1834–35 Spanish debt restructuring.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 505–524 (20 documents, batch17).');
