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

// --- Row 465: French Royal Tontine, August 1734 ---
setDoc(465,
  'Tontine Royale: Subscription Contract, Second Class, First Subdivision (Paris, August 1734)',
  'A printed subscription contract for the French royal tontine established by the King\'s edict of August 1734, drawn up before Notaires du Châtelet de Paris, including Michel Etienne Turgot (Chevalier, Seigneur de Soufflinois), François Poulain de Soissons, and others. The document records participation in the royal tontine under which 14,063,000 livres of life annuities (rentes viagères) were to be sold, organized in divisions (classes and subdivisions) of three livres de capital each, with annuities charged on the revenues of the Aides et Gabelles. Survivors of each subdivision were to receive the annuities of deceased members, concentrating income among the longest-lived participants. The tontine was a leading instrument of 18th-century French royal finance, combining elements of life insurance, speculation, and sovereign borrowing.',
  {
    type: 'Tontine',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Notaires du Châtelet de Paris',
    issueDate: '1734-08-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century',
    notes: 'French royal tontine, August 1734. Second Class, First Subdivision. 14,063,000 livres of life annuities. Drawn up before Notaires du Châtelet de Paris including Michel Etienne Turgot. Annuities charged on Aides et Gabelles revenues.',
  }
);

// --- Row 466: Compagnie Générale des Tramways de Moscou & de Russie, Action Privilégiée, 250 Francs, 1885 ---
setDoc(466,
  'Compagnie Générale des Tramways de Moscou & de Russie: Action Privilégiée, 250 Francs (No. 00274, Brussels, 1885)',
  'Preferred share (Action Privilégiée) No. 00274 of the Compagnie Générale des Tramways de Moscou & de Russie (General Company of Moscow and Russia Tramways), a Belgian société anonyme with registered offices in Brussels. The capital of 250 francs is fully paid (au porteur entièrement libérée). The company was established under contracts dated January 17, 1885, confirmed by M. Van Halteren, notary, and published in the Moniteur belge on February 4 and 7, 1885. Signed by two administrators. The company financed and operated urban tramway infrastructure in Moscow and other Russian cities, representing a significant Belgian capital investment in Russian transport.',
  {
    type: 'Stock',
    subjectCountry: 'Russia',
    issuingCountry: 'Belgium',
    creator: 'Compagnie Générale des Tramways de Moscou & de Russie',
    issueDate: '1885-02-08',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Action Privilégiée No. 00274, 250 francs fully paid. Belgian SA. Registered in Brussels. Contracts dated January 17, 1885, confirmed by M. Van Halteren, notary; published Moniteur belge February 4–7, 1885.',
  }
);

// --- Row 467: Egyptian Government Irrigation Works, Mandat de Paiement, £500, 1898 ---
setDoc(467,
  'Travaux d\'Irrigation du Gouvernement Égyptien (Assouân et Assiout 1898): Mandat de Paiement, £500 Sterling (No. 58/111)',
  'A payment warrant (Mandat de Paiement) No. 58 / No. 111, for £500 Sterling, issued by the Egyptian Government for the Aswan and Assiut Irrigation Works of 1898. The Egyptian Government acknowledges owing to MM. John Aird & Cie., or the bearer, the sum of Five Hundred Pounds Sterling, payable in London on January 1, 1932, at the offices of the Banque de l\'Afrique, with interest at 3½% per annum. The document relates to the major British-engineered irrigation projects at Aswan (the first Aswan Dam, completed 1902) and Assiut (Asyut Barrage), financed through payment warrants (mandats) issued against future revenue. John Aird & Co. was the principal British contractor for the Aswan Dam.',
  {
    type: 'Payment Warrant',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Egyptian Government, Ministry of Public Works',
    issueDate: '1898-01-01',
    currency: 'GBP',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Mandat de Paiement No. 58/111, £500 Sterling. Travaux d\'Irrigation du Gouvernement Égyptien, Assouân et Assiout 1898. Payable to John Aird & Cie. Due London, January 1, 1932, at Banque de l\'Afrique. Interest 3½% p.a.',
  }
);

// --- Row 468: Tennessee Chancery Court Injunction Bond, 1852 ---
setDoc(468,
  'Court of Chancery Injunction Bond, Madison County, Tennessee: Nathan H. Ballon and John E. Bostick, $1,900 (April 10, 1852)',
  'A handwritten surety bond (obligation) for $1,900 Lawful Money, filed in the Court of Chancery at Jackson, for the district of Madison County, Western Division of the State of Tennessee. Nathan H. Ballon and John E. Bostick bind themselves jointly and severally to pay the sum of nineteen hundred dollars, conditioned on Ballon prosecuting his bill of complaint against Thomas Farnell and others with full legal effect; the bond to be void if the suit succeeds, otherwise to remain in force. Dated April 10th, 1852. Security for the bond: property valued over $900, age 50 and 41 respectively. Approved by the Court. An example of 19th-century American legal surety practice in chancery proceedings.',
  {
    type: 'Legal Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Nathan H. Ballon; John E. Bostick',
    issueDate: '1852-04-10',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Chancery Court injunction bond, $1,900. Court of Chancery, Jackson, Madison County, Western Division, Tennessee. Nathan H. Ballon plaintiff; John E. Bostick surety. Dated April 10, 1852. Property security approved by Court.',
  }
);

// --- Row 469: Unilever N.V. Optiebewijs (Option Certificate), No. 97790, Rotterdam, 1937 ---
setDoc(469,
  'Unilever N.V.: Optiebewijs (Option Certificate) No. 97790 (Rotterdam, July 1937)',
  'Option certificate (Optiebewijs) No. 97790 of Unilever N.V., established in Rotterdam. The holder of eight such option certificates has the right to subscribe to eight new Unilever N.V. shares of f. 20 nominal value each (totaling f. 160). Valid until June 10, 1942 (van onwaarde na 10 June 1942). The certificate is linked to the conditions of Van den Bergh\'s Fabrieken N.V. 3½% Obligatielening (bond loan) of 1937, under which the option right was granted to bondholders. Dated Rotterdam, July 1937. Signed by the Secretary and a member of the Board of Directors. Represents a significant corporate finance instrument from one of the world\'s earliest modern consumer goods multinationals.',
  {
    type: 'Option',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Unilever N.V.',
    issueDate: '1937-07-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Optiebewijs No. 97790. Unilever N.V., Rotterdam. Eight certificates entitle holder to subscribe to 8 new shares at f.20 each (f.160 total). Valid until June 10, 1942. Related to Van den Bergh\'s Fabrieken N.V. 3½% Obligatielening, 1937.',
  }
);

// --- Row 470: United States of Brazil, 5% 20-Year Funding Bond, $100, 1931 ---
setDoc(470,
  'United States of Brazil: 5% 20-Year Funding Bond of 1931, $100 (No. S2940, Due October 1, 1951)',
  'A $100 bearer bond issued by the United States of Brazil under its 5% 20-Year Funding Loan of 1931, due October 1, 1951. Serial No. S2940. Printed in an ornate orange border with allegorical vignettes. Brazil promises to pay the bearer $100 on October 1, 1951, with semi-annual interest coupons at 5% per annum, payable at the City Bank Farmers Trust Company in New York. Authorized by Decree No. 20,251 of August 1931, representing a major Brazilian sovereign debt restructuring during the Great Depression. The Funding Loan consolidated Brazil\'s short-term obligations and was a key instrument of the country\'s efforts to maintain international creditworthiness during the global financial crisis.',
  {
    type: 'Bond',
    subjectCountry: 'Brazil',
    issuingCountry: 'Brazil',
    creator: 'United States of Brazil',
    issueDate: '1931-10-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Brazil 5% 20-Year Funding Bond of 1931, $100. No. S2940. Due October 1, 1951. Decree No. 20,251, August 1931. Semi-annual coupons. Payable at City Bank Farmers Trust Company, New York.',
  }
);

// --- Row 471: Holladay Overland Mail and Express Company, Blank Check, Virginia City, Montana, ca. 1866 ---
setDoc(471,
  'Holladay Overland Mail and Express Company: Blank Check, Virginia City, Montana Territory (No. 224S, ca. 1866)',
  'A blank (unused) check No. 224S from the Holladay Overland Mail and Express Company, drawn at Virginia City, Montana Territory. Payable to order in Dollars, signed by John Russell & Co., William, N.Y., acting as Treasurer and Agent of the Holladay Overland Mail and Express Company. The check features an engraved vignette of a frontier scene and a stagecoach. Ben Holladay\'s Overland Mail and Express Company was the dominant stagecoach and mail carrier in the American West in the 1860s, holding government mail contracts across transcontinental routes before Holladay sold the enterprise to Wells, Fargo & Co. in 1866. Unused financial paper from frontier-era western transportation finance.',
  {
    type: 'Check',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Holladay Overland Mail and Express Company',
    issueDate: '1866-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Blank check No. 224S. Holladay Overland Mail and Express Company. Virginia City, Montana Territory. Signed by John Russell & Co., William, N.Y., Treasurer. Ben Holladay\'s stagecoach enterprise; sold to Wells Fargo 1866.',
  }
);

// --- Row 472: Westchester & Philadelphia Rail Road Co., Scrip Certificate No. 262, $1,865, 1860 ---
setDoc(472,
  'Westchester & Philadelphia Rail Road Co.: Scrip Certificate No. 262, $1,865 at 6% (September 29, 1860)',
  'Scrip certificate No. 262 of the Westchester & Philadelphia Rail Road Co., authorized by resolution of the Board of Directors on September 29th, 1860. Certifies that Washington Hastings (or assigns) is entitled to receive the sum of One Thousand Eight Hundred and Sixty-Five Dollars ($1,865), or its equivalent in Railroad scrip, with interest at the rate of six percent per annum, on the first day of October. Amount of One Thousand Eight Hundred and Sixty-Five Dollars. Dated October 10, 1862. Signed by the Treasurer and bearing the company seal. The Westchester & Philadelphia Rail Road was a Pennsylvania short-line railroad serving the Philadelphia region. Railroad scrip was a common form of corporate obligation used in 19th-century American railroad finance.',
  {
    type: 'Scrip Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Westchester & Philadelphia Rail Road Co.',
    issueDate: '1860-09-29',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Scrip Certificate No. 262. Westchester & Philadelphia Rail Road Co. Washington Hastings, $1,865 at 6% interest. Board resolution September 29, 1860. Dated October 10, 1862.',
  }
);

// --- Row 473: Wisconsin Investment Company, Stock Warrant No. C8081, 100 Shares ---
setDoc(473,
  'Wisconsin Investment Company: Stock Warrant No. C8081, 100 Shares (ca. 1930s)',
  'An orange-printed stock warrant certificate No. C8081 for 100 shares of the Wisconsin Investment Company. Issued to Mary E. Carey. The reverse side of the certificate bears the company\'s articles or conditions of the warrant. The certificate features an allegorical vignette of a seated female figure with a globe, printed in orange ink on white paper, with an attached stub (No. C8081) and the company\'s seal. Signed by authorized officers. The Wisconsin Investment Company was a Midwestern investment holding company. The warrant structure was a common early-20th-century corporate finance mechanism granting holders the right to subscribe to shares at a specified price.',
  {
    type: 'Stock',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Wisconsin Investment Company',
    issueDate: '1930-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Stock Warrant No. C8081, 100 shares. Wisconsin Investment Company. Issued to Mary E. Carey. Orange-printed certificate with globe vignette and company seal. ca. 1930s.',
  }
);

// --- Row 474: Y. Rofé & Co. (Cairo), Acte Préliminaire de Vente, Suez Canal 3% Bond, 2nd Series ---
setDoc(474,
  'Y. Rofé & Co. (Cairo): Acte Préliminaire de Vente d\'Obligations, Canal de Suez 3%, 2ème Série',
  'A preliminary bond sale agreement (Acte Préliminaire de Vente d\'Obligations) issued by Y. Rofé & Co., bankers and brokers with offices in Cairo (Le Caire), Port Said, Alexandria (Alexandrie), and Mansoura. The document confirms that Y. Rofé & Co. has agreed to sell Suez Canal Company 3% bonds, 2nd Series (Oblig. Canal de Suez 3% 2ème Série) to the identified purchaser at the stated price and terms. The agreement specifies conditions for payment, delivery of definitive bond certificates, and recourse under Article 238 of the Code Civil Mixte in the event of failure to deliver. Issued from the Cairo head office (Rue d\'Ismaïlieh / Rue d\'Almaz, telephone 44-67). The Suez Canal Company\'s bonds were among the most internationally traded securities of the late 19th century.',
  {
    type: 'Contract',
    subjectCountry: 'Egypt',
    issuingCountry: 'Egypt',
    creator: 'Y. Rofé & Co., Le Caire',
    issueDate: '1880-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Acte Préliminaire de Vente d\'Obligations, Oblig. Canal de Suez 3%, 2ème Série. Y. Rofé & Co., Cairo; offices in Port-Said, Alexandrie, Mansoura. Reference to Code Civil Mixte, Art. 238. Ca. late 19th century.',
  }
);

// --- Row 475: German Reich, 8–15% Schatzanweisung, 5,000,000 Mark, Ausgabe II, 1923 ---
setDoc(475,
  'Deutsches Reich: 8–15% Schatzanweisung, 5,000,000 Mark Reichswährung, Ausgabe II (No. 73723, May 1923)',
  'An 8–15% variable-rate Treasury certificate (Schatzanweisung) of the German Reich, Issue II (Ausgabe II), Buchstabe A, No. 73723, for Five Million Mark Reichswährung (5,000,000 M). Issued May 20, 1923, at the height of the German hyperinflation. The certificate was authorized under the Reichsschuldenverwaltung (Imperial Debt Administration). It bears a strip of attached interest coupons, with payment dates marked from 1921 through the early 1930s. The variable interest rate (8–15%) reflected the extreme monetary instability of the Weimar Republic period. By the time of issuance in May 1923, the German mark had already lost the vast majority of its pre-war value; within months, 5,000,000 marks would become essentially worthless, culminating in the hyperinflationary collapse later that year.',
  {
    type: 'Treasury Certificate',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Reichsschuldenverwaltung',
    issueDate: '1923-05-20',
    currency: 'German Marks',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: '8–15% Schatzanweisung des Deutschen Reichs, Ausgabe II, Buchstabe A, No. 73723. 5,000,000 Mark Reichswährung. Issued May 20, 1923. Runs from September 1, 1921; matures December 31, 1930. Attached coupon strip. Weimar hyperinflation era.',
  }
);

// --- Row 476: Heeren van Zaamslag, Eendracht Polder Bond, £100 Flemish, No. 353, 18th century ---
setDoc(476,
  'Heeren van Zaamslag: Eendracht Polder Bond, £100 Flemish (No. 353, Island of Axel, Zeeland, ca. 18th Century)',
  'A handwritten and partly printed bond document No. 353 from the Heeren (Lords) of Zaamslag, as administrators of the newly completed Eendracht Polder reclamation on the island of Axel, Zeeland, Netherlands. The undersigned acknowledge being justly and truly indebted to Pieter Paul van Gelce, or the lawful bearer (wettigen Toonder), in the capital sum of One Hundred Pounds Flemish (£100 Ponden Vlaams), received as a loan for financing the polder and its operations. Interest at 3% per annum (à 3 per Cento vry Geld jaarlyks), payable on demand. Holders of the bond are entitled to participate by lot in the profits of the polder\'s agricultural produce. An early instrument of Dutch land reclamation finance, reflecting the sophisticated capital markets supporting 18th-century Zeeland water management.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Heeren van Zaamslag',
    issueDate: '1770-01-01',
    currency: 'Flemish Pounds',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Eendracht Polder bond, No. 353. Heeren van Zaamslag, island of Axel, Zeeland. Capital: £100 Flemish at 3% p.a. Issued to Pieter Paul van Gelce or bearer. Holders entitled to participate in polder produce profits distributed by lot. Dutch land reclamation finance, ca. 18th century.',
  }
);

// --- Row 477: Russian 4% Gold Loan, Sixth Issue 1894, Coupon Sheet (Russian), No. X035911 ---
setDoc(477,
  'Emprunt Russe 4% Or, Sixième Émission 1894: Coupon Sheet (Talon), 187 Roubles 50 Cop. (No. X035911)',
  'A large sheet of detachable interest coupons (talon) for the Russian Imperial Government 4% Gold Loan, Sixth Issue of 1894. The obligation denomination is 187 roubles 50 kopeks (equivalent to 500 French francs, £20, 405 Reichsmarks, or $96.25). Serial No. X035911. The sheet contains approximately 20 detachable coupons numbered in sequence (approximately coupons 97–116), each representing a semi-annual interest installment. Text is primarily in Russian, with French headers identifying the loan. Each coupon bears the coupon number, payment amount, and payment date. The 1894 Russian 4% Gold Loan was one of the largest sovereign bond issues of the late 19th century, placed across European financial markets.',
  {
    type: 'Coupon Sheet',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1894-01-01',
    currency: 'RUB',
    language: 'Russian, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Coupon sheet (talon), Russian 4% Gold Loan, Sixth Issue 1894. Denomination: 187 roubles 50 kopeks (= 500 fr. / £20 / 405 marks / $96.25). No. X035911. ~20 detachable coupons, ca. nos. 97–116. Russian text with French headers.',
  }
);

// --- Row 478: Russian 4% Gold Loan, Sixth Issue 1894, Bilingual Talon (German/English) ---
setDoc(478,
  'Russische 4% Gold-Anleihe, Sechste Emission 1894 / Russian 4% Gold Loan, Sixth Issue 1894: Bilingual Talon Coupon Sheet, 187 Rubel 50 Kop.',
  'A bilingual German/English talon coupon renewal sheet for the Russian Imperial Government 4% Gold Loan, Sixth Issue of 1894 (Russische 4% Gold-Anleihe, Sechste Emission, von 1894 / Russian 4% Gold Loan, Sixth Issue, 1894). Denomination: 187 Rubel 50 Kop. Gold (= 500 fr. / £20 / 405 Reichsmarks / $96.25). The sheet contains multiple rows of smaller detachable coupon stubs, each bilingual in German and English. The full bond terms are described in both languages, specifying interest payment locations: St. Petersburg, Moscow, Paris, London, Amsterdam, Berlin, Frankfurt, Hamburg, and Vienna. A key document illustrating the international distribution of Russian sovereign debt across European financial markets in the 1890s.',
  {
    type: 'Coupon Sheet',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1894-01-01',
    currency: 'RUB',
    language: 'German, English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Bilingual talon coupon sheet, Russische 4% Gold-Anleihe, Sechste Emission 1894 / Russian 4% Gold Loan, Sixth Issue 1894. Denomination: 187 Rubel 50 Kop. Gold = 500 fr. = £20 = 405 Marks = $96.25. German and English text with detachable coupon stubs.',
  }
);

// --- Row 479: Imperial Chinese Government, 5% Gold Loan 1903, Obligation 500 Francs, No. 30659, Brussels, 1907 ---
setDoc(479,
  'Gouvernement Impérial de Chine: Emprunt Chinois 5% Or 1903, Obligation de 500 Francs (No. 30659, Brussels, April 12, 1907)',
  'A 500-franc bearer bond (Obligation au Porteur) No. 30659 issued by the Imperial Chinese Government under its 5% Gold Loan of 1903 (Emprunt Chinois 5% Or 1903). Issued in Brussels on April 12, 1907. The total loan capital was 60,000,000 francs; interest at 5% per annum, payable semi-annually in gold. The bond bears Chinese imperial characters (seals and authorization signatures) alongside French text, with attached coupon stubs for future interest payments. The loan was placed through European banking syndicates and guaranteed by the revenues of Chinese Imperial Customs. Part of the series of late Qing dynasty foreign loans taken to finance government expenditures, the 1903 Chinese Gold Loan is an important example of early 20th-century Chinese sovereign borrowing in European capital markets.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'Belgium',
    creator: 'Imperial Chinese Government',
    issueDate: '1907-04-12',
    currency: 'FRF',
    language: 'French, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Gouvernement Impérial de Chine. Emprunt Chinois 5% Or 1903. Obligation de 500 Francs au Porteur, No. 30659. Brussels, April 12, 1907. Capital: 60,000,000 francs. Interest 5% p.a. in gold, semi-annual. Guaranteed by Chinese Imperial Customs revenues. Late Qing dynasty foreign loan.',
  }
);

// --- Row 480: Imperial Chinese Government, 5% Gold Loan 1903, Obligation Générale (General Terms and Amortization Table) ---
setDoc(480,
  'Gouvernement Impérial de Chine: Emprunt Chinois 5% Or 1903 – Obligation Générale (General Conditions and Amortization Table)',
  'The general conditions document (Obligation Générale) accompanying the Imperial Chinese Government 5% Gold Loan of 1903. The document reproduces the Édit de S.M. l\'Empereur de Chine (edict of His Majesty the Emperor of China) authorizing the loan, and provides full terms under the heading Création de l\'Emprunt (creation of the loan), including the total capital, interest terms, maturity, and security provisions. An amortization table at the bottom shows the annual schedule for redemption of bonds over the life of the loan (approximately 20 years). Signed by a Chinese minister. This document is either the reverse of a bond certificate or a separate general terms prospectus distributed with the 1903 Chinese Gold Loan, and provides essential context for the legal and financial structure of Qing-era sovereign borrowing.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'Belgium',
    creator: 'Imperial Chinese Government',
    issueDate: '1903-01-01',
    currency: 'FRF',
    language: 'French, Chinese',
    numberPages: 1,
    period: '20th Century',
    notes: 'Obligation Générale / general conditions for Emprunt Chinois 5% Or 1903. Includes Édit de S.M. l\'Empereur de Chine authorizing the loan, Création de l\'Emprunt terms, and amortization table. Signed by Chinese Minister of Finance. Likely reverse of bond certificate or separate prospectus sheet.',
  }
);

// --- Row 481: Imperial Russian Government, 4% Gold Loan, Sixth Issue 1894, Obligation 125 Rubles Gold, No. 59944 ---
setDoc(481,
  'Российское Императорское Правительство: 4% Золотой Заём, 6-й Выпуск 1894 г., Облигация 125 Рублей Золотом (No. 59944)',
  'An ornate Russian Imperial Government 4% Gold Loan bearer bond, Sixth Issue of 1894 (Российский Четырехпроцентный Золотой Заём, Шестой Выпуск 1894 года). Obligation No. 59944, denomination 125 Rubles in gold (equivalent to 500 French francs), 6th emission. The bond features the Imperial Russian double-headed eagle at center and an elaborate decorative printed border. Text in Russian. The loan was authorized under the Imperial ukase of February 3, 1894, for a total of 125,000,000 rubles at 4% interest, redeemable by annual drawings over 50 years. Interest payable at the major financial centers of Europe: St. Petersburg, Moscow, Paris (via Crédit Lyonnais), London (Hope and Baring), Amsterdam (Lippmann), Berlin, Frankfurt, Hamburg, and Vienna.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1894-02-03',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Russian Imperial 4% Gold Loan, Sixth Issue 1894. Obligation No. 59944, 125 Rubles Gold (= 500 fr.). Ornate design with Imperial double-headed eagle. Total loan: 125,000,000 rubles at 4%, redeemable over 50 years. Imperial ukase February 3, 1894.',
  }
);

// --- Row 482: Imperial Russian Government, 4% Gold Loan 1894, Multilingual General Terms Sheet ---
setDoc(482,
  'Gouvernement Impérial de Russie: Emprunt Russe 4% Or, Sixième Émission 1894 – General Terms Sheet (French, German, English, Russian)',
  'A multilingual general terms sheet for the Russian Imperial Government 4% Gold Loan, Sixth Issue of 1894, presenting the full bond conditions in four languages: French (Gouvernement Impérial de Russie, Emprunt Russe 4% Or, Sixième Émission, 1894), German (Kaiserlich Russische Regierung, Russische 4% Gold-Anleihe, Sechste Emission), English (Imperial Government of Russia, Russian Four Per Cent Gold Loan, Sixth Issue, 1894), and Russian. Denomination: 125 Roubles Gold = 500 fr. = £20 = 405 Reichsmarks = $96.25. Interest payable semi-annually at St. Petersburg, Moscow, Paris (Crédit Lyonnais), London (Hope, Baring), Amsterdam (Lippmann), Berlin, Frankfurt, Hamburg, and Vienna. Redeemable by annual lot drawings over 50 years from 1894.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian Government',
    issueDate: '1894-01-01',
    currency: 'RUB',
    language: 'French, German, English, Russian',
    numberPages: 1,
    period: '19th Century',
    notes: 'Quadrilingual (French/German/English/Russian) general terms sheet for Russian 4% Gold Loan, Sixth Issue 1894. Denomination: 125 Roubles Gold = 500 fr. = £20 = 405 Marks = $96.25. Payment at St. Petersburg, Moscow, Paris, London, Amsterdam, Berlin, Frankfurt, Hamburg, Vienna.',
  }
);

// --- Row 483: Leib-Renten-Lotterey der Stadt Hamburg, Lottery Ticket No. 7027, 200 Mark Banco, 1773 ---
setDoc(483,
  'Leib-Renten-Lotterie der Stadt Hamburg: Lottery Ticket No. 7027, 200 Mark Banco (Hamburg, April 23, 1773)',
  'A printed lottery ticket No. 7027 for the Life Annuity Lottery (Leib-Renten-Lotterey) of the City of Hamburg, published on April 23, 1773. The bearer of this ticket holds one lot in the life annuity lottery, credited under the noted number in the Kämmerei Lotterie-Conto in Banco for the sum of Two Hundred Mark Banco (Zwey Hundert Mark Banco). Printed with the Hamburg city coat of arms (two white towers on a red field) at top. The Hamburg Life Annuity Lottery was a form of municipal finance that combined lottery participation with life annuity-style returns: participants purchased lottery tickets, and winners received life annuities paid from Hamburg\'s municipal treasury. Such instruments were a common and sophisticated method of city-state borrowing in 18th-century northern Europe.',
  {
    type: 'Lottery Ticket',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'City of Hamburg (Kämmerei)',
    issueDate: '1773-04-23',
    currency: 'Mark Banco',
    language: 'German',
    numberPages: 1,
    period: '18th Century',
    notes: 'Leib-Renten-Lotterey der Stadt Hamburg. No. 7027, 200 Mark Banco. Published April 23, 1773. Credited in Kämmerei Lotterie-Conto in Banco. Life annuity lottery; Hamburg city coat of arms. Municipal finance instrument.',
  }
);

// --- Row 484: Oliver Wolcott, Comptroller of Public Accounts (Connecticut), Interest Receipt No. 2435, 1789 ---
setDoc(484,
  'Oliver Wolcott, Comptroller of Public Accounts (Connecticut): Interest Receipt No. 2435, £1 15s 3d (Hartford, April 4, 1789)',
  'A printed and handwritten interest receipt No. 2435 issued by Oliver Wolcott, Comptroller of the Public Accounts of Connecticut, acknowledging payment of One Pound Fifteen Shillings and Three Pence Lawful Money (£1 15s 3d) to John Douglas. The payment represents interest on four State Notes with a principal of £29.6.4½, computed to the first of February 1789. Issued at the Comptroller\'s Office, Hartford, Connecticut, April 4th, 1789. Signed by Stephen Crosby (or similar). Oliver Wolcott Sr. served as Comptroller of Connecticut\'s public accounts during the post-Revolutionary War period and later as U.S. Secretary of the Treasury. This receipt documents Connecticut\'s efforts to service its Revolutionary War debt through regular interest payments to note holders.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Oliver Wolcott, Comptroller of Public Accounts, Connecticut',
    issueDate: '1789-04-04',
    currency: 'Connecticut Pounds',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'Receipt No. 2435. Oliver Wolcott, Comptroller of Public Accounts, Connecticut. Interest on 4 State Notes (principal £29.6.4½): £1 15s 3d Lawful Money, computed to February 1, 1789. Issued to John Douglas. Hartford, April 4, 1789.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 465–484 (20 documents, batch15).');
