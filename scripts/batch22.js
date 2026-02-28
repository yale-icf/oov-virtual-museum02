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

// NOTE: Rows 620–629 are skipped — pre-existing data already in spreadsheet.

// Row 605: State of Connecticut Treasury-Office Certificate, £12.17.02, No. 11593, June 1, 1782
setDoc(605,
  'State of Connecticut, Treasury-Office Certificate No. 11593, £12.17.02, June 1, 1782',
  'Connecticut state treasury certificate No. 11593 for £12.17.02, issued June 1, 1782 by the Treasury-Office at Hartford. Certifies that a named holder who served in the Continental Army is entitled to receive this sum being one fourth part of the balance found on his orders at the office, in Gold or Silver, with interest thereon, with reference to an Act of the General Assembly. Printed by Hudson and Goodwin, Hartford. Signed by the Treasurer.',
  {
    type: 'bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Connecticut, Treasury-Office',
    issueDate: 'June 1, 1782',
    currency: 'Connecticut pound',
    language: 'English',
    numberPages: '1',
    period: 'American Revolutionary period',
    notes: 'Connecticut state debt certificate; £12.17.02; No. 11593; Continental Army payee; printed by Hudson and Goodwin, Hartford; one-fourth payment in gold or silver'
  }
);

// Row 606: Dutch Life Annuity Tontine Conditions No. 21, 500 Guilders on 20 Lives, 1688
setDoc(606,
  'Dutch Life Annuity Tontine Conditions, No. 21 (Een en Twintig), f.500 on 20 Lives, 1688',
  'Printed copy (COPIA) of life annuity tontine conditions No. 21 (Nummer Een en Twintig), setting out the agreed terms for investing 500 guilders capital on the lives of twenty persons over 30 years old. The document establishes that the full capital sum of 500 guilders per lyer (life-share) must be paid within 45 days; interest payments are to be made annually each August; and the tontine fund is to run for thirty years. References "Pieter de Lange" as a principal figure and sets the first year\'s interest day as August 1, 1688. An important early Dutch life annuity/tontine document.',
  {
    type: 'tontine',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Pieter de Lange (director)',
    issueDate: '1688',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: '17th century',
    notes: 'Life annuity tontine No. 21; f.500 capital on 20 lives over 30 years; August 1688 first interest payment; 30-year term; early Dutch tontine document'
  }
);

// Row 607: Nicolaas Brant Suriname Negotiatie Conditions, f.150,000 at 6%, Amsterdam, May 1, 1765
setDoc(607,
  'Nicolaas Brant Suriname Plantation Negotiatie Conditions, f.150,000 at 6%, Amsterdam, May 1, 1765',
  'Printed conditions and pledge document for the Suriname plantation negotiatie under the direction of Nicolaas Brant, merchant in Amsterdam, dated May 1, 1765. The director\'s pledge statement opens: "Ik Ondergeschreven NICOLAAS BRANT, als dirigerende de Negotiatie breeder in het geannexeerde Plan gemeld, beloove, en verbindende my by dezen, omme my in myne gedagte qualiteit stipulyk te zullen gedragen aan de Conditien..." The negotiatie was to advance 150,000 Dutch guilders to planters in the Colony of Suriname at 6% per annum interest. Articles I–III of the conditions follow, covering mortgage requirements, interest obligations, and planter eligibility.',
  {
    type: 'negotiatie',
    subjectCountry: 'Netherlands|Suriname',
    issuingCountry: 'Netherlands',
    creator: 'Nicolaas Brant',
    issueDate: 'May 1, 1765',
    currency: 'Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Mid-18th century',
    notes: 'Suriname plantation negotiatie; f.150,000 at 6%; Nicolaas Brant director, Amsterdam; Articles I–III of conditions; plantation mortgage requirements; 1765'
  }
);

// Row 608: Notarial Document - Heshuysen/Balmer Dominica Plantation Mortgage, Amsterdam, December 21, 1777
setDoc(608,
  'Notarial Copy - Heshuysen & Compagnie / James Balmer, Dominica Plantation Paaschal Mortgage, Amsterdam, December 21, 1777',
  'Notarial copy (Copie) of a complex financial transaction recorded before notary Hermanus de Wolff Junior, Amsterdam, December 21, 1777. Concerns a mortgage (Mortgage of Hypotheeq) on Plantation Paaschal, situated on the island of Dominica, originally established August 13, 1772, by James Balmer, London merchant, for the benefit of Adolf Jan Heshuysen, Floris Visscher Heshuysen, and Frans Jacob Heshuysen, composing the firm Adolf Jan Heshuysen en Compagnie, Haarlem. The mortgage secures a sum of 7,100 pounds sterling. The plantation had earlier been encumbered with life annuities of 1,269 pounds and 10 shillings per annum on various lives in Great Britain. The document records the formal presentation of this mortgage before the Amsterdam notary.',
  {
    type: 'negotiatie',
    subjectCountry: 'Netherlands|Dominica',
    issuingCountry: 'Netherlands',
    creator: 'Hermanus de Wolff Junior (Notaris); Adolf Jan Heshuysen en Compagnie',
    issueDate: 'December 21, 1777',
    currency: 'British pound sterling|Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 18th century',
    notes: 'Amsterdam notarial copy; Plantation Paaschal, Dominica; Heshuysen & Comp. / James Balmer; mortgage £7,100; life annuities £1,269.10/yr; notary Hermanus de Wolff Junior; December 21, 1777'
  }
);

// Row 609: Notarial Document - Heshuysen/Balmer Dominica Plantation Mortgage, Amsterdam, March 21, 1777
setDoc(609,
  'Notarial Copy - Heshuysen & Compagnie / James Balmer, Dominica Plantation Paaschal Mortgage, Amsterdam, March 21, 1777',
  'Notarial copy (Copie) of an earlier stage of the same Heshuysen/Balmer Dominica plantation mortgage transaction, recorded before notary Hermanus de Wolff Junior, Amsterdam, March 21, 1777. Also concerns the mortgage on Plantation Paaschal (Dominica) by James Balmer, London, for the benefit of Adolf Jan Heshuysen, Floris Visscher Heshuysen, and Frans Jacob Heshuysen (Adolf Jan Heshuysen en Compagnie, Haarlem), with reference to the same mortgage of August 13, 1772. This document predates the December 1777 notarial copy by nine months and represents an earlier phase of the same complex transaction.',
  {
    type: 'negotiatie',
    subjectCountry: 'Netherlands|Dominica',
    issuingCountry: 'Netherlands',
    creator: 'Hermanus de Wolff Junior (Notaris); Adolf Jan Heshuysen en Compagnie',
    issueDate: 'March 21, 1777',
    currency: 'British pound sterling|Dutch guilder',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 18th century',
    notes: 'Amsterdam notarial copy; Plantation Paaschal, Dominica; Heshuysen & Comp. / James Balmer; earlier phase of December 1777 transaction; notary Hermanus de Wolff Junior; March 21, 1777'
  }
);

// Row 610: Credit Foncier Cubain / Banco Territorial de Cuba, 3% Series A Bearer Bond, $96.16/500 Fcs.
setDoc(610,
  'Credit Foncier Cubain / Banco Territorial de Cuba, 3% Series A Bearer Obligation, $96.16 US Gold / 500 Francs',
  'Series A bearer mortgage bond (Obligacion al Portador) of the Credit Foncier Cubain / Banco Territorial de Cuba (Cuban Territorial/Mortgage Bank), bearing 3% annual interest. Face value: $96.16 US Gold Coin, equivalent to 500 French Francs. Interest payable annually on January 1 and July 1 of each year. The left portion of the bond contains extensive conditions text in Spanish and French. The Banco Territorial de Cuba was established to provide agricultural and real estate mortgage credit in colonial Cuba.',
  {
    type: 'bond',
    subjectCountry: 'Cuba',
    issuingCountry: 'Cuba',
    creator: 'Credit Foncier Cubain / Banco Territorial de Cuba',
    issueDate: 'ca. 1880-1898',
    currency: 'United States dollar|French franc',
    language: 'Spanish|French',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Cuban mortgage bank; Series A; 3% bond; $96.16 US Gold = 500 Fcs.; annual interest Jan 1 & Jul 1; colonial Cuba'
  }
);

// Row 611: Connecticut Comptroller's Certificate No. 79, Five Pounds, July 1, 1809
setDoc(611,
  'State of Connecticut, Comptroller\'s Certificate No. 79, Five Pounds, July 1, 1809',
  'Connecticut state Comptroller\'s certificate No. 79, dated July 1, 1809, certifying that a named holder is entitled to receive the sum of Five Pounds Lawful Money, out of any Funds appropriated for the Payment of interest on the liquidated Debt of the State of Connecticut. Printed by Hudson and Goodwin, Hartford. Signed by the Comptroller. A small cancellation hole is punched through the certificate. This is a routine state debt interest payment warrant from the post-Revolutionary era.',
  {
    type: 'bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Connecticut, Comptroller\'s Office',
    issueDate: 'July 1, 1809',
    currency: 'Connecticut pound',
    language: 'English',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Connecticut state debt interest warrant; No. 79; £5; liquidated debt interest payment; printed Hudson and Goodwin, Hartford; punched cancellation hole'
  }
);

// Row 612: Connecticut Comptroller's Certificate No. 3107, Five Shillings, October 2, 1789
setDoc(612,
  'State of Connecticut, Comptroller\'s Certificate No. 3107, Five Shillings, Andrew Kingsbury, October 2, 1789',
  'Connecticut state Comptroller\'s certificate No. 3107, dated October 2, 1789, certifying that Andrew Kingsbury is entitled to receive the sum of Five Shillings Lawful Money, out of any Funds appropriated for the Payment of interest on the liquidated Debt of the State of Connecticut. Printed by Hudson and Goodwin, Hartford. Signed by Ralph Pomeroy, Comptroller. A small cancellation hole is punched through the centre. One of the smallest denomination Connecticut debt certificates, this instrument dates from the first year of the US Constitution.',
  {
    type: 'bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Connecticut, Comptroller\'s Office',
    issueDate: 'October 2, 1789',
    currency: 'Connecticut shilling',
    language: 'English',
    numberPages: '1',
    period: 'American Revolutionary period',
    notes: 'Connecticut state debt interest warrant; No. 3107; 5 shillings; Andrew Kingsbury; signed Ralph Pomeroy, Comptroller; printed Hudson and Goodwin; punched cancellation hole; 1789'
  }
);

// Row 613: Denver & Rio Grande Spoorweg-Maatschappij, Dutch Certificaat Preferred Stock, $1,000, No. 669, 1890
setDoc(613,
  'Denver & Rio Grande Spoorweg-Maatschappij (Preferred Stock), Dutch Certificaat van $1,000, No. 669, Amsterdam, 1890',
  'Dutch bearer certificate (Certificaat) No. 669 for $1,000 face value, representing 10 shares of $100 each in the Preferred Stock of the Denver & Rio Grande Spoorweg-Maatschappij (Denver & Rio Grande Railway Company). The certificate was issued and administered by the Internationale Administratie-Kantoor in Amsterdam, enabling Dutch investors to hold American railroad preferred stock through a Dutch bearer instrument. Interest and dividends payable in Amsterdam. Dated Amsterdam, December 6, 1890. Signed by the administrators.',
  {
    type: 'share',
    subjectCountry: 'United States',
    issuingCountry: 'Netherlands',
    creator: 'Internationale Administratie-Kantoor, Amsterdam',
    issueDate: 'December 6, 1890',
    currency: 'United States dollar',
    language: 'Dutch',
    numberPages: '1',
    period: 'Late 19th century',
    notes: 'Denver & Rio Grande Railway preferred stock; Dutch bearer certificaat; No. 669; $1,000 = 10 shares at $100; Internationale Administratie-Kantoor Amsterdam; December 6, 1890'
  }
);

// Row 614: Deutsche Äußere Anleihe 1924 (Dawes Loan), Swiss Issue, No. 10080, 200 CHF, 1953
setDoc(614,
  'Deutsche Äußere Anleihe 1924 (Dawes-Anleihe), Schweizerische Ausgabe, Fundierungsschuldverschreibung No. 10080, 200 CHF, 1953',
  'Post-war German funding bond (Fundierungsschuldverschreibung) No. 10080 for 200 Swiss Francs, issued by the Federal Republic of Germany (Bundesschuldenverwaltung, Bad Homburg) in 1953. Issued under the terms of the London Debt Agreement on German Foreign Debts (Londoner Abkommen über Deutsche Auslandsschulden 1953), Annex I, in exchange for interest arrears on the Deutsche Äußere Anleihe 1924 (Dawes Loan), Swiss issue. Represents 8 semi-annual interest payments of 25 CHF each that accrued between October 16, 1944 and October 15, 1952. Redemption deferred until German reunification ("nicht vor der Wiedervereinigung Deutschlands"). Signed by Bundesschuldenverwaltung. Bad Homburg v.d.H., 1953.',
  {
    type: 'bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Bundesrepublik Deutschland, Bundesschuldenverwaltung',
    issueDate: '1953',
    currency: 'Swiss franc',
    language: 'German',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'Dawes Loan 1924 Swiss issue; post-WWII arrears bond; No. 10080; 200 CHF; 8 semi-annual arrears 1944-1952; London Debt Agreement 1953; redemption deferred until German reunification'
  }
);

// Row 615: Hope & Co. Dutch Certificate for 3% French Government Annuity, No. 309, Frs.300/year, Amsterdam, 1861
setDoc(615,
  'Certificaat van Drie Percents Fransche Fondsen, No. 309, Frs. 300/year (on Frs. 10,000 capital), Hope & Co., Amsterdam, 1861',
  'Dutch certificate No. 309 (Certificaat van Drie Percents Fransche Fondsen) entitling the bearer to Frs. 300 annual income (Trois Cents Franken Rente) inscribed in the French National Debt Register (Grootboek der Publieke Schuld van Frankryk), representing a nominal capital of Frs. 10,000. Managed by the Administratie-Kantoor in Amsterdam under the direction of Hope & Co., Ketwich & Voombergh, and Wed. W. Borski. Interest to be received and paid out by the Company on behalf of the holder, according to the conditions of the certificate. Amsterdam, October 19, 1861. Signed by all three managing houses. The twentieth issue of this certificate type (Twintigste Uitreikingbrief).',
  {
    type: 'certificate',
    subjectCountry: 'France',
    issuingCountry: 'Netherlands',
    creator: 'Hope & Co.; Ketwich & Voombergh; Wed. W. Borski',
    issueDate: 'October 19, 1861',
    currency: 'French franc',
    language: 'Dutch',
    numberPages: '1',
    period: 'Mid-19th century',
    notes: '3% French public debt (Rente); No. 309; Frs. 300/year income on Frs. 10,000 nominal capital; Hope & Co., Ketwich & Voombergh, Wed. Borski; Amsterdam October 19, 1861; 20th certificate series'
  }
);

// Row 616: Dutch Russian 6% Assignation Funds Certificate No. 1266, 1,000 Roubles, Amsterdam, 1828
setDoc(616,
  'Dutch Certificaat for Russian 6% Assignation Funds, No. 1266, 1,000 Roubles, Amsterdam, October 7, 1828',
  'Dutch certificate No. 1266 (Kapitaal R. 1000) entitling the bearer to 1,000 Roubles in Assignations in the Russian Imperial Commission of Amortization (Keiserlyke Commissie van Amortisatie) at St. Petersburg, bearing 6% interest. The funds are deposited with the Algemeen Directeur van Administratie in Amsterdam, under the direction of Gerardus Blancke en Zoon, Chemet & Weetjen, and Van den Broeke & Comp., as Bewaarders (custodians) of the Fund, with main direction (Hoofddirectie) by Van Vloten & de Gyselaar, Buys & Zoon, and H.W. van Driel Slam. Amsterdam, October 7, 1828. Signed by counter-signatories.',
  {
    type: 'certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Gerardus Blancke en Zoon; Chemet & Weetjen; Van den Broeke & Comp.',
    issueDate: 'October 7, 1828',
    currency: 'Russian ruble',
    language: 'Dutch',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Russian 6% Assignation Funds; No. 1266; 1,000 roubles; custodians: Blancke & Weetjen & Van den Broeke; main directors: Van Vloten & Gyselaar, Buys & Zoon, H.W. van Driel Slam; Amsterdam October 7, 1828'
  }
);

// Row 617: German External Loan 1924 (Dawes Loan), Trilingual Rights Certificate No. A 025220, £40, 1960
setDoc(617,
  'German External Loan 1924 (Dawes-Anleihe), Trilingual Rights Certificate No. A 025220, £40, Bad Homburg, January 4, 1960',
  'Trilingual rights certificate (Bezugschein / Rights Certificate / Bon) No. A 025220 for £40 Sterling, issued by the Federal Republic of Germany under the London Debt Agreement of 1953. Covers both the British Issue (Britische Ausgabe) and French Tranche of the Deutsche Äußere Anleihe 1924 (German External Loan 1924 / Emprunt Extérieur Allemand 1924), also known as the Dawes Loan. The certificate entitles the holder to receive new 3% funding bonds (Fundierungsschuldverschreibungen) from the German Federal Republic in exchange for arrears interest on the original 1924 bonds. Trilingual in German, English, and French. Bad Homburg, January 4, 1960. Signed by Bundesschuldenverwaltung.',
  {
    type: 'bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Bundesrepublik Deutschland, Bundesschuldenverwaltung',
    issueDate: 'January 4, 1960',
    currency: 'British pound sterling',
    language: 'German|English|French',
    numberPages: '1',
    period: 'Mid-20th century',
    notes: 'Dawes Loan 1924; trilingual Bezugschein/Rights Certificate/Bon; No. A 025220; £40; British and French tranches; London Debt Agreement 1953; Bad Homburg January 4, 1960'
  }
);

// Row 618: German Government International 5½% Loan 1930 (Young Plan), Gold Bond No. 16917, $1,000
setDoc(618,
  'German Government International 5½% Loan 1930 (Young Plan), Dollar Gold Bond No. 16917, $1,000, Due 1965',
  'US Dollar-denominated bearer gold bond No. 16917 of the German Government International 5½ Per Cent Loan 1930 (Young Plan), face value $1,000, due June 1, 1965. Interest at 5½% per annum payable semi-annually on June 1 and December 1. Principal and interest payable at the office of J.P. Morgan & Co. in the Borough of Manhattan, City of New York. The left side of this large-format bond contains the full "GENERAL BOND" conditions text. Printed in red/maroon on cream paper. The Young Plan Loan (1930) was issued to fund Germany\'s WWI reparations under the Young Committee schedule.',
  {
    type: 'bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'German Government (Deutsches Reich)',
    issueDate: '1930',
    currency: 'United States dollar',
    language: 'English',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Young Plan Loan 1930; No. 16917; $1,000; 5.5% gold bond; due June 1, 1965; J.P. Morgan & Co., New York; WWI reparations finance'
  }
);

// Row 619: German Government International 5½% Loan 1930 (Young Plan), French Tranche, No. A.2009599, 1,000 Frs.
setDoc(619,
  'German Government International 5½% Loan 1930 (Young Plan), French Tranche, Bearer Bond No. A.2009599, 1,000 French Francs',
  'Trilingual bearer bond No. A.2009599 of the German Government International 5½% Loan 1930 (Young Plan), French tranche, for 1,000 French Francs. Title in three languages: German ("Internationale 5½%ige Anleihe des Deutschen Reichs 1930"), English ("German Government International 5½ Per Cent. Loan, 1930"), and French ("Emprunt International 5½% 1930 du Gouvernement Allemand"). The bond provides a 1,000 Franc bearer obligation, signed by the German Reich representative. The full trilingual conditions are printed across the bond. This is the French (Französische Ausgabe) tranche of the same Young Plan loan issued in the dollar market (see related bond).',
  {
    type: 'bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'German Government (Deutsches Reich)',
    issueDate: '1930',
    currency: 'French franc',
    language: 'German|English|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Young Plan Loan 1930; No. A.2009599; 1,000 French Francs; French (Französische Ausgabe) tranche; trilingual German/English/French; WWI reparations finance'
  }
);

// ROWS 620–629: SKIPPED — pre-existing data already in spreadsheet

// Row 630: Hope & Comp. Russian 6% Bank Assignation Certificate No. 3074, 1,000 Roubles, Amsterdam, 1827
setDoc(630,
  'Hope & Comp. Russian 6% Bank Assignation Funds Certificate No. 3074, 1,000 Roubles, Amsterdam, September 7, 1827',
  'Bilingual Dutch/French certificate (Certificaat / Certificat d\'Inscription Russe) No. 3074 entitling the bearer to an inscription of 1,000 Roubles in 6% Russian Funds (Russische Fondsen), payable in Bank Assignations (Bank-Assignatien), inscribed in the Grand Livre (Groot-Boek) at St. Petersburg. Administered at the Administratie-Kantoor in Amsterdam under the direction of Hope en Comp., Ketwich en Voombergh, and Wed. W. Borski. Amsterdam, September 7, 1827. Signed by the three managing houses. Issued with ten semi-annual coupons from July 1, 1829. The bilingual format reflects the certificate\'s distribution to both Dutch and French-speaking investors.',
  {
    type: 'certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Hope en Comp.; Ketwich en Voombergh; Wed. W. Borski',
    issueDate: 'September 7, 1827',
    currency: 'Russian ruble',
    language: 'Dutch|French',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Russian 6% Bank Assignation Funds; No. 3074; 1,000 roubles; bilingual Dutch/French; Hope & Co., Ketwich & Voombergh, Wed. Borski; Amsterdam September 7, 1827; 10 semi-annual coupons from July 1, 1829'
  }
);

// Row 631: Stadnitski & van Heukelom et al., Russian Assignation Funds Certificate L.A. No. 2081, 1,000 Roubles, 1832
setDoc(631,
  'Russian Public Debt (Assignation Funds) Certificate L.A. No. 2081, 1,000 Roubles, Stadnitski & van Heukelom et al., Amsterdam, December 31, 1832',
  'Bilingual Dutch/French certificate (Certificaat / Certificat) L.A. No. 2081, entitling the bearer to a capital of 1,000 Roubles in Assignations inscribed in the Grand Livre (Grootboek) of the Russian Imperial Commission of Amortization (Keiserlyke Commissie van Amortisatie) at St. Petersburg. Funds held by the Algemeen Directeur van Administratie in Amsterdam under the direction of Stadnitski & van Heukelom, Jacob van Beeck Vollenhoven, Samuel & David Saportas, Lamaison & Bouwer, Johannes Samuel Wurfbain, Hendrik Ovens & Zoon, and de Lanoy & Burlage. Amsterdam, December 31, 1832. Signed by multiple directors. Issued with semi-annual coupons. Based on the prospectus of June 30, 1834.',
  {
    type: 'certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Stadnitski & van Heukelom; Ovens & Zoon; de Lanoy & Burlage et al.',
    issueDate: 'December 31, 1832',
    currency: 'Russian ruble',
    language: 'Dutch|French',
    numberPages: '1',
    period: 'Early 19th century',
    notes: 'Russian 6% Assignation Funds; L.A. No. 2081; 1,000 roubles; bilingual Dutch/French; Stadnitski & van Heukelom group; Amsterdam December 31, 1832; prospectus of June 30, 1834'
  }
);

// Row 632: 3½% England War Loan 1932, Dutch Certificaat SPECIMEN, £100, Amsterdam
setDoc(632,
  '3½% Engeland War Loan 1932, Dutch Certificaat aan Toonder SPECIMEN, £100, Amsterdam (Twintigste Trust-Maatschappij)',
  'Specimen Dutch bearer certificate (Certificaat aan Toonder) No. 15145 for £100, representing holdings in the British Government 3½% England War Loan 1932. Issued by the Administratiekantoor van De Twintigste Trust-Maatschappij (The Twentieth Trust Company) in Amsterdam. Marked "SPECIMEN." The certificate was issued under administration conditions available free of charge at the issuing office. Interest (coupons) payable June 1 and December 1. Administered by Algemeen Kantoor van Administratie te Amsterdam B.V. The England War Loan 1932 was the British government\'s landmark conversion of WWI 5% War Loan to 3½%, the largest ever voluntary debt conversion.',
  {
    type: 'certificate',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'Netherlands',
    creator: 'Administratiekantoor van De Twintigste Trust-Maatschappij, Amsterdam',
    issueDate: 'ca. 1932-1935',
    currency: 'British pound sterling',
    language: 'Dutch',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'SPECIMEN; 3.5% England War Loan 1932; £100; Dutch bearer certificaat No. 15145; De Twintigste Trust-Maatschappij Amsterdam; coupons Jun 1 & Dec 1; WWI debt conversion'
  }
);

// Row 633: Austrian Bau-Los (Housing Lottery Bond) 1921, K 1,200, Series 1,790, No. 060
setDoc(633,
  'Österreichisches Bau-Los Em. 1921, Housing Lottery Bond, K 1,200 (Tausendzweihundert Kronen), Series 1,790, No. 060',
  'Austrian federal government-guaranteed housing lottery bond (Bau-Los / Construction Lottery) 1921 issue, Series 1,790, No. 060, face value 1,200 Kronen (K 1200). Guaranteed by the Federal Government (Vom Bunde garantiert) and secured by mortgage registration (grundbücherlich einverleibt). Issued by the Wohn- und Siedlungsfonds (Housing and Settlement Fund) to finance post-WWI housing construction. Features two vignette illustrations of Austrian suburban cottage homes. The lottery element means bonds were redeemed by prize draws, with winning bonds receiving premium payments. Signed by the Fund\'s authorized representative.',
  {
    type: 'bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Austria',
    creator: 'Österreichischer Wohn- und Siedlungsfonds',
    issueDate: '1921',
    currency: 'Austrian krone',
    language: 'German',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Austrian housing lottery bond; K 1,200; Series 1,790 No. 060; federal government guaranteed; mortgage registered; Wohn- und Siedlungsfonds; post-WWI housing finance; cottage vignettes'
  }
);

// Row 634: City of Baku 5% Loan 1910, Bearer Bond No. 00603, 189 Roubles / £20 / 504 Francs
setDoc(634,
  'City of Baku, 5% Loan 1910, Bearer Bond No. 00603, 189 Roubles = £20 Sterling = 504 Francs',
  'Trilingual (Russian, English, French) bearer bond (Облигация / Bond / Obligation) No. 00603 of the 5% Loan of the City of Baku (5% Заемъ Города Баку 1910 Года / 5% Loan of the City of Baku, 1910 / 5% Emprunt de la Ville de Bakou, 1910). Face value: 189 Russian Roubles = 20 Pounds Sterling = 504 French Francs. Total loan: 26,999,973 Roubles = £2,867,340 = 71,999,928 Francs. Authorized by the Russian Council of Ministers (Vysochaische upolnomochennyi) by decree of October 8/15, 1909. Interest at 5% per annum payable semi-annually on January 15 and July 15 (old style). Signed by the Mayor (Gorodskoy Golova) and members of the Municipal Council of Baku. Baku, capital of the Baku Governorate, Russian Empire (now Azerbaijan).',
  {
    type: 'bond',
    subjectCountry: 'Azerbaijan',
    issuingCountry: 'Russia',
    creator: 'City of Baku Municipal Council',
    issueDate: '1910',
    currency: 'Russian ruble|British pound sterling|French franc',
    language: 'Russian|English|French',
    numberPages: '1',
    period: 'Early 20th century',
    notes: 'Baku municipal bond; No. 00603; 189 roubles / £20 / 504 francs; total loan 26,999,973 roubles; authorized October 1909; trilingual Russian/English/French; signed by Mayor of Baku; Russian Imperial Azerbaijan'
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 605–619 and 630–634 (20 documents, batch22); rows 620–629 skipped (pre-existing data).');
