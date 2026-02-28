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

// --- Row 525: Stadnitski & van Heukelom et al., Certificaat L.A. No. 5320, 1000 Roubles in Assignations, Amsterdam, June 8, 1875 ---
setDoc(525,
  'Stadnitski & van Heukelom et al.: Certificaat / Certificat, 1000 Roubles in Assignations (L.A. No. 5320, Amsterdam, June 8, 1875)',
  'A bilingual Dutch/French certificate L.A. No. 5320 for Duizend Roebels (1,000 Roubles) in Assignations / Mille Roubles en Assignats, representing an inscription at 6% in the Openboek der Keizerlijke Schuld de Russie (State Debt Book of the Imperial Russian Amortization Commission at St. Petersburg). Administered in Amsterdam by the bureau under direction of Stadnitski & van Heukelom, Jacob van Beeck Vollenhoen, Samuel & David Saportas, Lamaison & Bouwer, Johannes Samuel Wurfbain, Hendrik Oyens, and De Lanoy & Burlage. Issued Amsterdam, June 8, 1875. Capital R° 1,000 in Assignation. Includes ten half-yearly coupons. The holder may at any time reclaim the original inscription against return of the certificate and unused coupons.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Stadnitski & van Heukelom; Jacob van Beeck Vollenhoen; Samuel & David Saportas; Lamaison & Bouwer; Johannes Samuel Wurfbain; Hendrik Oyens; De Lanoy & Burlage',
    issueDate: '1875-06-08',
    currency: 'RUB',
    language: 'Dutch, French',
    numberPages: 1,
    period: '19th Century',
    notes: 'Certificaat L.A. No. 5320, 1,000 Roubles in Assignations. 6% inscription, Imperial Russian State Debt Book (Commission at St. Petersburg). Amsterdam, June 8, 1875. Administered by Stadnitski & van Heukelom et al. Ten half-yearly coupons.',
  }
);

// --- Row 526: London Stock Exchange, WWI Regulation 10 Good Delivery Certificate, City of Baker 5% 1910 Bond, April 29, 1916 ---
setDoc(526,
  'London Stock Exchange: WWI Regulation 10 Good Delivery Certificate, City of Baker 5% 1910 Bond No. 58707 (April 29, 1916)',
  'A printed certificate issued under the Temporary Regulations for the Re-Opening of the Stock Exchange, Regulation 10, certifying that the security "City of Baker 5% 1910, Bond for £20, Numbered 58707" has been expressly passed by the Committee as a good delivery, special cause having been shown. Signed by the Secretary, Share and Loan Department, Stock Exchange, dated April 29, 1916. Regulation 10 required that no securities to bearer or endorsed in blank could be delivered as good delivery unless stamped prior to October 1, 1914 and accompanied by a broker/dealer declaration, or specially passed by the Committee. The regulations were introduced at the outbreak of World War I to prevent trading in securities possibly held by enemy nationals.',
  {
    type: 'Certificate',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'London Stock Exchange (Share and Loan Department)',
    issueDate: '1916-04-29',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'London Stock Exchange WWI Regulation 10 good delivery certificate. City of Baker 5% 1910 Bond No. 58707, £20. Signed Secretary, Share and Loan Department. April 29, 1916. WWI wartime enemy-ownership regulations.',
  }
);

// --- Row 527: US Treasury, 6% Funded Debt Certificate, $10,000, Daniel Ternwiles & Wife, Register Office, June 16, 1792 ---
setDoc(527,
  'Treasury of the United States: 6% Funded Debt Certificate, $10,000, Daniel Ternwiles & Wife (Register Office, June 16, 1792)',
  'A printed and handwritten certificate from the Treasury of the United States, Register Office, dated June 16, 1792, No. 42.94. "Be it known that there is due from the United States of America unto Daniel Ternwiles & Wife, or their Assigns, the Sum of Ten Thousand Dollars, bearing Interest at Six per Cent per Annum from the first Day of January A.D. One Thousand Eight Hundred and One, inclusively; payable quarter-yearly, and subject to Redemption by Payments not exceeding, in one Year, the Proportion of Eight Dollars upon a Hundred of the Stock bearing Interest at Six per Cent; created by Virtue of an Act making Provision for the Debt of the United States, passed on the fourth Day of August, 1790." Signed by Joseph Nourse, Register of the Treasury. Notarized by Clement Biddle, Notary Public for the Commonwealth of Pennsylvania. A key instrument of Alexander Hamilton\'s federal debt consolidation under the Funding Act of 1790.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Treasury of the United States; Joseph Nourse, Register',
    issueDate: '1792-06-16',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Treasury 6% Funded Debt Certificate No. 42.94, $10,000. Daniel Ternwiles & Wife. Register Office, June 16, 1792. Signed Joseph Nourse, Register. Notarized by Clement Biddle. Created under Funding Act of August 4, 1790. Hamilton\'s federal debt consolidation.',
  }
);

// --- Row 528: Dutch Uitgestelde Schuld (Deferred Debt) Certificate No. 6231, f.500, ca. 1814-1815 ---
setDoc(528,
  'Bewijs van Voldoening aan de Wet: Uitgestelde Schuld Certificate No. 6231, f. 500 (Netherlands, ca. 1814–1815)',
  'A Dutch certificate of compliance with the law (Bewijs van Voldoening aan de Wet) No. 6231, for a Uitgestelde Schuld (Deferred/Delayed Debt) of f. 500 capital, issued pursuant to the Commission\'s decree (Besluit van Z.K.H.) of June 7, 1814, No. 23, concerning the Wisselbank exchange advances and conditions of the Wet (Law) of May 14, 1814. Holder W.F. Holling and the Burgen[?] are recorded. The certificate confirms compliance with Articles 5, 12, and 18 of the aforementioned Law, with the Kapitalen recorded in the Oude Grootboek der (Hollandsche Schuld), re-registered in the name of C. Niemanstedt [?]. References Afschrijvings-Billet No. [?] van het Grootboek, and a new Grootboek amount of Veertien Honderd Guldens. The Uitgestelde Schuld was the class of Dutch national debt whose interest payments were deferred during the Napoleonic period, restructured after the restoration of the Dutch monarchy in 1813-1814.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Dutch Government Commission (Commissie voor de Uitgestelde Schuld)',
    issueDate: '1815-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Bewijs van Voldoening aan de Wet, Uitgestelde Schuld No. 6231, f. 500. Pursuant to Besluit Z.K.H. June 7, 1814, No. 23, and Wet of May 14, 1814. Holder: W.F. Holling. Post-Napoleonic Dutch deferred debt restructuring. Recorded in Hollandsche Schuld Grootboek.',
  }
);

// --- Row 529: United States of America, Bill of Exchange No. 131, $300 / 1300 Livres Tournois, October 9, 1776 ---
setDoc(529,
  'United States of America: Bill of Exchange No. 131, $300 (1300 Livres Tournois) for Interest on Loan Office Certificates (October 9, 1776)',
  'A printed and handwritten bill of exchange No. 131, issued by the United States of America on October 9, 1776. At Thirty Days Sight, payable to Annie Brown (or Order) in Thirteen Hundred Livres Tournois, equivalent to Three Hundred Dollars, for Interest due on Money borrowed by the United States. Countersigned by [?] Smith, Commissioner of the Continental Loan Office, State of Pennsylvania. Signed by H. Atkinson, Treasurer of Loans. One of the earliest-dated examples in the collection, issued just months after the Declaration of Independence, as the Continental Congress began issuing Loan Office Certificates and related bills of exchange to service its domestic war debt.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Pennsylvania',
    issueDate: '1776-10-09',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 131, $300 / 1,300 Livres Tournois. October 9, 1776. Payable to Annie Brown. Interest on Loan Office Certificates. Pennsylvania Loan Office; signed by H. Atkinson, Treasurer of Loans. Among the earliest-dated Continental bills of exchange.',
  }
);

// --- Row 530: Volksstaat Hessen, Roggen-Anleihe, 5% Schuldverschreibung, 10 Zentner Roggen, No. 002152, Darmstadt, December 12, 1923 ---
setDoc(530,
  'Volksstaat Hessen: Roggen-Anleihe, 5% Schuldverschreibung über den Geldwert von 10 Zentner Roggen (Abteilung 1 No. 002152, Darmstadt, December 12, 1923)',
  'A 5% Schuldverschreibung (bond certificate) of the Volksstaat Hessen (People\'s State of Hesse), Roggen-Anleihe (Rye Loan), Abteilung 1, Buchstabe D, No. 002152. The bond\'s face value equals the market price of 10 Zentner (hundredweights) of rye (Roggen), denominated to preserve real purchasing power during the German hyperinflation. Issued in Darmstadt, December 12, 1923, by the Hessische Staatsschuldenverwaltung (Hessian State Debt Administration). The State of Hesse pledges its entire property and revenues as security. The bond carries 5% annual interest and is part of the broader Weimar-era practice of issuing commodity-indexed Sachwertanleihen (real-value bonds) tied to agricultural and industrial commodities to circumvent the destruction of nominal monetary values by hyperinflation.',
  {
    type: 'Bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Germany',
    creator: 'Hessische Staatsschuldenverwaltung (Volksstaat Hessen)',
    issueDate: '1923-12-12',
    currency: 'German Marks',
    language: 'German',
    numberPages: 1,
    period: '20th Century',
    notes: 'Roggen-Anleihe, Volksstaat Hessen. 5% Schuldverschreibung, 10 Zentner Roggen. Abteilung 1, Buchst. D, No. 002152. Darmstadt, December 12, 1923. Hessische Staatsschuldenverwaltung. Commodity-indexed rye bond (Sachwertanleihe). Weimar hyperinflation.',
  }
);

// --- Row 531: United States of America, Bill of Exchange No. 64, $36 / 180 Livres Tournois, November 12, 1778 ---
setDoc(531,
  'United States of America: Bill of Exchange No. 64, $36 (180 Livres Tournois) for Interest on Loan Office Certificates (November 12, 1778)',
  'A printed and handwritten bill of exchange No. 64, issued by the United States of America on November 12, 1778. At Thirty Days Sight, payable to Jesse White (or Order) in One Hundred and Eighty Livres Tournois, equivalent to Thirty-six Dollars, for Interest due on Money borrowed by the United States. Countersigned by Nath. Appleton, Commissioner of the Continental Loan Office in the State of Massachusetts Bay. Signed by H. Atkinson, Treasurer of Loans. Part of the same late-1778 series as other Continental bills of exchange in the collection.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Massachusetts Bay',
    issueDate: '1778-11-12',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 64, $36 / 180 Livres Tournois. November 12, 1778. Payable to Jesse White. Interest on Loan Office Certificates. Massachusetts Bay Loan Office; Nath. Appleton, Commissioner; H. Atkinson, Treasurer of Loans.',
  }
);

// --- Row 532: United States of America, Bill of Exchange No. 71, $120 / 600 Livres Tournois, November 4, 1778 ---
setDoc(532,
  'United States of America: Bill of Exchange No. 71, $120 (600 Livres Tournois) for Interest on Loan Office Certificates (November 4, 1778)',
  'A printed and handwritten bill of exchange No. 71, issued by the United States of America on November 4, 1778. At Thirty Days Sight, payable to M. John Tompkins (or Order) in Six Hundred Livres Tournois, equivalent to One Hundred and Twenty Dollars, for Interest due on Money borrowed by the United States. Countersigned by Nath. Appleton, Commissioner of the Continental Loan Office in the State of Massachusetts Bay. Signed by H. Atkinson, Treasurer of Loans. Part of the same November 1778 series, predating No. 64 (goetzmann0531) by eight days.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Massachusetts Bay',
    issueDate: '1778-11-04',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 71, $120 / 600 Livres Tournois. November 4, 1778. Payable to M. John Tompkins. Interest on Loan Office Certificates. Massachusetts Bay Loan Office; Nath. Appleton, Commissioner; H. Atkinson, Treasurer of Loans.',
  }
);

// --- Row 533: Dutch Suriname Plantation Negotiatie Conditions, Lever en de Bruine, Directors, Samuel & Jacobus Walland, Commissioners ---
setDoc(533,
  'Conditien Eener Negotiatie: Suriname Plantation Loan, f. 1,000,000, Lever en de Bruine, Directors; Samuel & Jacobus Walland, Commissioners (Amsterdam)',
  'A printed prospectus of conditions (Conditien eener Negotiatie) for a plantation mortgage negotiatie of one million (Tienmaal honderd Duyzend) Dutch guilders, under the direction of Lever en de Bruine, Kooplieden t\'Amsterdam, and with Samuel en Jacobus Walland as Commissioners. The negotiatie is open to Planters in the COLONIE ZURINAMEN (Suriname) who wish to participate. The conditions set out the mortgage obligations (Hypotheek) on the colonial plantations, the terms for loans to planters, the administration of the fund by the Amsterdam directors, the role of the Local Director of the Colony of Suriname, and the conditions under which the planters\' production and property is pledged as security. This is an example of the Dutch colonial plantation mortgage fund system (plantageleningen), through which Amsterdam investors financed West Indian sugar, coffee, and cotton plantations.',
  {
    type: 'Prospectus',
    subjectCountry: 'Suriname',
    issuingCountry: 'Netherlands',
    creator: 'Lever en de Bruine; Samuel en Jacobus Walland (Commissioners)',
    issueDate: '1760-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Conditien eener Negotiatie, f. 1,000,000. Suriname plantation mortgage negotiatie. Directors: Lever en de Bruine, Amsterdam. Commissioners: Samuel & Jacobus Walland. Dutch plantageleningen system for West Indian colonial plantation finance.',
  }
);

// --- Row 534: Dutch Negotiatie for Vlaardingen Orphanage (Weeshuys), f.75,000, Batavian Republic, May 9, 1800 ---
setDoc(534,
  'Negotiatie, f. 75,000: Loterij ten behoeve van het Weeshuys der Stad Vlaardingen, Bataafschen Volks (May 9, 1800)',
  'A printed lottery/negotiatie prospectus of f. 75,000 Dutch guilders (Vyf en Zeventig Duyzend Gulden), constituted by order of the Vertegenwoordigend Lichaam ver Bataafschen Volks (Representative Body of the Batavian People), dated May 9, 1800. Established for the maintenance of the Orphanage (Weeshuys) of the city of Vlaardingen. The negotiatie consists of 1,500 shares at f. 50 each, with a lottery prize structure: the first prize is f. 10,000, then f. 5,000, f. 2,500, f. 1,000, f. 500, f. 250, f. 125, and smaller prizes, totaling f. 48,000. Prize drawings are held on July 31 and January 15 annually. Holders of the remaining (non-prize) classes receive annual interest (Instrerende Lijfrentebrieven) divided into five classes by age. An innovative Batavian Republic instrument combining charitable endowment, lottery, and life annuity.',
  {
    type: 'Lottery Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Vertegenwoordigend Lichaam van het Bataafschen Volks; Weeshuys Vlaardingen',
    issueDate: '1800-05-09',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Negotiatie f. 75,000 / 1,500 shares at f. 50. For Weeshuys (Orphanage) of Vlaardingen. Batavian Republic, May 9, 1800. Lottery prizes: f.10,000 down to f.125. Annual life annuity interest by age class for non-prize holders.',
  }
);

// --- Row 535: Dutch Charitable Bond for the Poor of Haarlem, f.500 at 2.5%, No. 110, January 1, 1805 ---
setDoc(535,
  'Voor den Armen (For the Poor), Haarlem: Charitable Bond No. 110, f. 500 at 2½% (January 1, 1805)',
  'A Dutch charitable bond No. 110 (Voor den Armen — For the Poor), Haarlem, issued by the Ontvanger der Stadslijke Belasting (Receiver of the City Tax), Haarlem, by authority of the special qualification of the Raad (Council) of the city, by resolution of June 5, 1804. The bond acknowledges receipt from M. Bodisco of the sum of Vyf-Honderd Guldens (f. 500) placed at interest at the rate of Penning 40 (2½%) per year from January 1, 1805. The capital (f. 500) and interest (f. 12.50 per year) are secured upon all city revenues and goods, especially the revenues of the former city Guilds. Multiple later annotations record coupon payments from 1822 through the 1860s. An example of municipal charitable bond finance in the Napoleonic-era Netherlands.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'City of Haarlem (Ontvanger der Stadslijke Belasting)',
    issueDate: '1805-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Voor den Armen, Haarlem. Charitable bond No. 110, f. 500 at Penning 40 (2½%). Issued by Ontvanger der Stadslijke Belasting, Haarlem. Council resolution June 5, 1804. Signed M. Bodisco. Coupon annotations from 1822–1860s. Municipal charitable finance.',
  }
);

// --- Row 536: William of Orange-Nassau: Archaic Dutch Princely Bond / Obligation, 17th Century ---
setDoc(536,
  'Willem, Prince of Orange-Nassau: Archaic Dutch Princely Bond / Obligation (Netherlands, 17th Century)',
  'A printed archaic Dutch-language bond or financial obligation beginning "Y VVILHELM: BY DER GRATIE GODES, Prince van [Vernorde-der-dijck]..." referencing the Prince of Orange-Nassau and financial obligations associated with his treasury (Tresorier General and Rentmeester). The document references the name "FREDERICUS HENRICX, Prince van Oraenien" (Frederick Henry, Prince of Orange, 1584–1647) and appears to concern a princely debt obligation, potentially related to the fiscal administration of the Orange-Nassau domains. The archaic spelling and typography suggest a mid-17th century document. This is a rare and historically important early modern sovereign financial instrument from the Dutch Golden Age.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Prince of Orange-Nassau (Willem / Frederick Henry)',
    issueDate: '1640-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '17th Century',
    notes: 'Archaic Dutch princely bond/obligation. "Y VVILHELM BY DER GRATIE GODES, Prince van [?]..." References FREDERICUS HENRICX (Frederick Henry), Prince of Orange. Tresorier General and Rentmeester noted. Mid-17th century Dutch Golden Age fiscal document.',
  }
);

// --- Row 537: King Gustav III of Sweden, Royal Loan Bond, 400,000 Rijksdaalders, Amsterdam, February 9, 1782 ---
setDoc(537,
  'Wy Gustavus, Koning van Sweeden: Royal Loan Bond, 400,000 Rijksdaalders Hollandsch Courant (Amsterdam, February 9, 1782)',
  'A printed royal Swedish government bond beginning "Wy GUSTAVUS, door Gods Genade Koning van Sweeden, der Gothen en Wenden &c. &c. &c., Erfheer van Noorwegen, Hertog van Schleswyck, Hollstein, Stornarn en Ditmarschen, Grave van Oldenburg, Delmenhorst &c. &c." Issued in Amsterdam on February 9, 1782. The bond is for Vierhonderd Duizend Rijksdaalders Hollandsch Courant Geld (400,000 Rijksdaalers Dutch currency), raised through the Amsterdam money market via Johannes Verge [?] as agent and translator. The document sets out the terms of the Swedish royal loan, interest payments, and the security provided by the Swedish Crown. King Gustav III of Sweden (reigned 1771–1792) frequently borrowed in the Amsterdam capital market to finance his ambitious domestic and foreign policies, including the Russo-Swedish War of 1788–1790.',
  {
    type: 'Bond',
    subjectCountry: 'Sweden',
    issuingCountry: 'Netherlands',
    creator: 'King Gustav III of Sweden',
    issueDate: '1782-02-09',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Swedish royal loan bond, King Gustav III. 400,000 Rijksdaalders Hollandsch Courant. Amsterdam, February 9, 1782. Issued through Amsterdam money market. Johannes Verge, agent/translator. Gustav III borrowed extensively in Amsterdam to finance Swedish policy.',
  }
);

// --- Row 538: Karel de Tweede (Charles II), Habsburg-related Dutch Bond Document, early 18th century ---
setDoc(538,
  'Wy Karel, de Tweede: Habsburg Royal Bond / Obligation Document (Amsterdam, early 18th Century)',
  'A printed Dutch-language bond or obligation document beginning "Wy KAREL, de Tweede, door Gods/Godes Roome Keyser..." (We Charles, the Second, by the Grace of God, Holy Roman Emperor...). The document references Willem Gideon Donz. tot Amsterdam as a director or agent, and discusses shares (Aandelen) and obligations (Obligaties) in the context of a loan or financial instrument. The text mentions various parties, interest payments in Hollandsch Courant Geld, and conditions related to a sovereign bond. "Karel de Tweede" here likely refers to Emperor Charles VI (1685–1740), also known as Charles III as a claimant to the Spanish throne, who was the Habsburg Holy Roman Emperor from 1711 to 1740 and sponsored various financial ventures including the Ostend Company. The document appears to be a Dutch-intermediated Habsburg imperial bond from the early 18th century.',
  {
    type: 'Bond',
    subjectCountry: 'Austria',
    issuingCountry: 'Netherlands',
    creator: 'Habsburg Imperial Government (Charles VI / Karel de Tweede)',
    issueDate: '1720-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Dutch bond/obligation document. "Wy KAREL, de Tweede, door Gods Roome Keyser..." Habsburg imperial bond. Agent: Willem Gideon Donz., Amsterdam. Aandelen and Obligaties in Hollandsch Courant Geld. Likely refers to Emperor Charles VI (1711–1740). Dutch-intermediated Habsburg imperial finance.',
  }
);

// --- Row 539: German Government International 5½% Loan 1930, Belgian Tranche, 100 Belgas / 500 Belgian Francs, No. A.22399 ---
setDoc(539,
  'German Government International 5½% Loan, 1930: Belgian Issue Bond to Bearer, 100 Belgas / 500 Belgian Francs (No. A.22399)',
  'A trilingual (German/English/French) bearer bond No. A.22399 of the Internationale 5½%ige Anleihe des Deutschen Reichs 1930 / German Government International 5½ Per Cent. Loan, 1930 / Emprunt International 5½ P.C. 1930 du Gouvernement Allemand. Belgian issue (Belgische Ausgabe / Belgian Issue / Tranche Belge). Denomination: 100 (Hundred) Belgas = 500 (Five Hundred) Belgian Francs. Printed with an ornate design including the Weimar German eagle seal. Signed by the Reichsschuldenverwaltung (Reich Debt Administration). The Young Plan International Loan of 1930 was a major international bond issue by the Weimar Republic, placed simultaneously in multiple European currencies (British pounds, French francs, Belgian francs, Swedish kronor, Swiss francs, and US dollars) as part of the reparations settlement under the Young Plan.',
  {
    type: 'Bond',
    subjectCountry: 'Germany',
    issuingCountry: 'Belgium',
    creator: 'German Government (Reichsschuldenverwaltung)',
    issueDate: '1930-01-01',
    currency: 'Belgian Francs',
    language: 'German, English, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'German Government International 5½% Loan 1930. Belgian Issue. No. A.22399, 100 Belgas = 500 Belgian Francs. Trilingual German/English/French. Reichsschuldenverwaltung. Young Plan international bond issue, simultaneous multi-currency placement.',
  }
);

// --- Row 540: Potosian Land Grant, Back/Reverse Side with Vignette, 400 Acres ---
setDoc(540,
  'Potosian Land Grant: Reverse Side / Back of 400 Acre Certificate (London, ca. 1825)',
  'The reverse side (back) of a Potosian Land Grant certificate for 400 Acres. The reverse shows the ornate printed header of the certificate featuring the POTOSIAN LAND GRANT title, a central coat of arms or heraldic vignette (400 Acres / 400 Acres), and handwritten endorsements or signatures at the bottom (including initials "RW," "AS," and a further signature "W.R. Ramos" or similar). The Potosian Land Grant certificates were issued by a British company in London ca. 1825 as part of speculative investment schemes in the Bolivian Potosí region following South American independence. This reverse side complements the front face of the certificate (see goetzmann0506), together comprising the complete land grant document.',
  {
    type: 'Land Grant',
    subjectCountry: 'Bolivia',
    issuingCountry: 'United Kingdom',
    creator: 'Potosian Land Grant Company (London)',
    issueDate: '1825-04-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Potosian Land Grant, reverse/back side of 400-acre certificate. Shows printed header with coat of arms and handwritten endorsements. Complements front face (goetzmann0506). London, ca. 1825. British speculative investment in Bolivian Potosí region.',
  }
);

// --- Row 541: Dutch Essequibo Plantation Negotiatie, De Vyver Plantation, Remy & Comp., Amsterdam, 1789 ---
setDoc(541,
  'Dutch Essequibo Plantation Negotiatie: De Vyver Plantation, Remy & Comp., Directors (Amsterdam, July 1789)',
  'A Dutch legal document drawn up before notaries Albertus Bakker and Paules Cordes, Raaden in den Edelen Aghtbaren Heeren van Public, Criminele en Civiele Justicie der Rivier en onderhorige districten van Essequebo (Essequibo Colony). The parties include Johan Gottlieb Detericks and Vrouwe Johanna Maria von de Vyver, owners of plantations in Essequibo. The document concerns the establishment of a negotiatie (Nieuwe Creditors) for the plantation "De Vyver" and other Midland Plantations, with capital of f. 13,000 (Dertienduizend Guldens), under the management of Remy & Comp., Kooplieden t\'Amsterdam as Directors, and with the Amsterdam firm directing operations under the terms of a July 1789 plan. This is an example of the Essequibo plantation mortgage lending system administered from Amsterdam in the late 18th century.',
  {
    type: 'Contract',
    subjectCountry: 'Guyana',
    issuingCountry: 'Netherlands',
    creator: 'Remy & Comp. (Amsterdam); Albertus Bakker; Paules Cordes (Notaries)',
    issueDate: '1789-07-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Dutch Essequibo plantation negotiatie. "De Vyver" plantation and Midland Plantations. Capital f. 13,000. Directors: Remy & Comp., Amsterdam. Notaries: Albertus Bakker, Paules Cordes, Essequibo Colony. July 1789. Dutch colonial plantation mortgage system.',
  }
);

// --- Row 542: French Princes-in-Exile Bond, Louis Stanislas Xavier & Charles Philippe, No. 34, 1000 Dutch Guilders, Amsterdam, September 28, 1793 ---
setDoc(542,
  'French Princes in Exile (Louis Stanislas Xavier & Charles Philippe): Obligation No. 34, f. 1,000 Dutch Guilders (Amsterdam, September 28, 1793)',
  'A remarkable historical bond document issued under authority of the French Princes in exile, bearing the authorization of LOUIS STANISLAS XAVIER (the future King Louis XVIII) and CHARLES PHILIPPE (the Comte d\'Artois, future King Charles X), countersigned by Le Maréchal, Duc de Broglie (from Hamm en Westphalie, signed on the twenty-fifth day of September 1793). No. 34. The Amsterdam house of J. Bourcourd and Wedowe F. Croese & Comp. certifies receipt, on behalf of Their Royal Highnesses the Princes of France, from an unnamed party, of the sum of EEN DUIZEND GULDENS Hollandsch Courant (f. 1,000 Dutch guilders) in exchange for a share in the stated Obligation, with corresponding coupons attached. Amsterdam, September 28, 1793. Notarized and authenticated. This bond was issued by the French royal family members during their exile following the execution of Louis XVI, as a means of raising funds in the Dutch capital market for the Royalist cause.',
  {
    type: 'Bond',
    subjectCountry: 'France',
    issuingCountry: 'Netherlands',
    creator: 'Louis Stanislas Xavier (Louis XVIII); Charles Philippe (Comte d\'Artois); Maréchal Duc de Broglie',
    issueDate: '1793-09-28',
    currency: 'NLG',
    language: 'Dutch, French',
    numberPages: 1,
    period: '18th Century',
    notes: 'French émigré princes\' bond No. 34, f. 1,000 Dutch Guilders. Authorized by Louis Stanislas Xavier (future Louis XVIII) and Charles Philippe (Comte d\'Artois). Duc de Broglie, Hamm en Westphalie, September 25, 1793. Amsterdam, J. Bourcourd & Wedowe F. Croese & Comp., September 28, 1793. French Royalist exile finance.',
  }
);

// --- Row 543: Amsterdam Forward Contract for British Consolidated Stock, Ricardo & De Lara, Amsterdam, May 13, 1805 ---
setDoc(543,
  'Ricardo & De Lara (Amsterdam): Forward Contract for British Consolidated Stock, £[?] (Amsterdam, May 13, 1805)',
  'A handwritten forward sale contract (Dutch-language) arranged through Ricardo en De Lara, Warpenfort over het Wardhuis, No. 70, Amsterdam. The undersigned commits to sell a quantity of Pounds Sterling of the Consolidated Stock of Great Britain (British Consols / Geconsolideerde Schuld van [?] Ann [?]) at a set price, to be delivered at the Recontre (settlement date) in London at the agreed price. All deliveries and conditions are governed by the December 1803 Reglement referenced. Amsterdam, May 13, 1805. Signed by the counterparty. The Ricardo firm in Amsterdam (related to the famous Ricardo family including economist David Ricardo) was a leading broker in Anglo-Dutch financial transactions, facilitating the forward trading of British government bonds in Amsterdam — a major center for international consol dealing.',
  {
    type: 'Contract',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'Netherlands',
    creator: 'Ricardo & De Lara (Amsterdam)',
    issueDate: '1805-05-13',
    currency: 'GBP',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Amsterdam forward contract for British Consolidated Stock (Consols). Ricardo en De Lara, No. 70 Amsterdam. Governed by December 1803 Reglement. Amsterdam, May 13, 1805. Ricardo family Amsterdam firm; Anglo-Dutch financial intermediation.',
  }
);

// --- Row 544: Dutch Plantation Negotiatie Conditions, Daniel Changuion, Essequibo & Demerara, f.400,000 at 6%, ca. 1816 ---
setDoc(544,
  'Conditien van Negotiatie: Daniel Changuion, Director; Plantation Loans for Essequibo & Demerara Planters, f. 400,000 at 6% for 10 Years (Amsterdam, ca. 1816)',
  'A printed set of conditions (Conditien van Negotiatie) for a Dutch plantation mortgage fund, under the direction of Daniel Changuion, to provide a sum of f. 400,000 Dutch guilders to planters in the Rio Essequebo and Rio Demerary (Essequibo and Demerara, present-day Guyana) for the improvement and continuation of their plantations, at 6% interest per annum for 10 years. [1816, 10 Claviers, 10 Coupons]. The six articles set out: the obligation of planters to mortgage their plantations and inventories (enslaved persons, livestock, equipment), the conditions for valuation of the mortgaged property at one-fifth of assessed worth, the requirement for planters to deliver produce to the Directors, the annual valuation of the plantation, and provisions for default and foreclosure. A significant example of Dutch colonial plantation mortgage finance in the post-Napoleonic period.',
  {
    type: 'Prospectus',
    subjectCountry: 'Guyana',
    issuingCountry: 'Netherlands',
    creator: 'Daniel Changuion (Director)',
    issueDate: '1816-01-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Conditien van Negotiatie. Director: Daniel Changuion. Plantation loans for Essequibo & Demerara planters. f. 400,000 at 6% p.a. for 10 years. 10 Claviers, 10 Coupons. Ca. 1816. Mortgages on plantations and inventories. Dutch colonial plantation finance, post-Napoleonic period.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 525–544 (20 documents, batch18).');
