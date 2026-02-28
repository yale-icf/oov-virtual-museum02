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

// --- Row 485: Dutch Negotiatie on French Life Annuities, Subscription List, Amsterdam, May 1, 1787 ---
setDoc(485,
  'Negotiatie van de Compagnie Française Lijfrente: Subscription Record, Nominees Nos. 30–59 (Amsterdam, May 1, 1787)',
  'A printed document dated May 1, 1787, compiled before Notaries Abraham Fock and Adriaan Schaap as Commissioners, listing nominees Nos. 30 through 59 for the Compagnie Française Lijfrente Negotiatie van Maart 1781. The nominees are predominantly women, registered under numbered entries. The Dutch negotiatie on French life annuities was a popular investment vehicle of the 1780s whereby Dutch investors purchased shares in a pooled fund that in turn bought French royal life annuities (rentes viagères), selecting young female nominees to maximize the expected payment duration. Annuity income accumulated for surviving nominees\' benefit, concentrating payments among the longest-lived. The document records the official register of nominees as at May 1, 1787, serving as proof of each nominee\'s enrollment in the annuity pool.',
  {
    type: 'Tontine',
    subjectCountry: 'France',
    issuingCountry: 'Netherlands',
    creator: 'Abraham Fock; Adriaan Schaap (Commissioners)',
    issueDate: '1787-05-01',
    currency: 'FRF',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Subscription list Nos. 30–59 for the Compagnie Française Lijfrente Negotiatie van Maart 1781. Amsterdam, May 1, 1787. Commissioners: Abraham Fock and Adriaan Schaap. Nominees predominantly women. Dutch negotiatie on French royal life annuities (rentes viagères).',
  }
);

// --- Row 486: Hope & Comp. et al., Certificaat der Vijfde Reeks, 5% Russische Fondsen in Zilver, 500 Roubles, Amsterdam, 1834 ---
setDoc(486,
  'Hope & Comp., Ketwich & Voombergh, en Weduwe Willem Borski: Certificaat der Vijfde Reeks, 5% Russische Fondsen in Zilver, 500 Roubles (No. 13, Amsterdam, August 29, 1834)',
  'A certificaat (certificate) of the Fifth Series (Vijfde Reeks) of 5% Russian Funds in Silver (Russische Fondsen in Zilver), No. 13, for a capital of 500 Roubles, issued in Amsterdam on August 29, 1834, by the leading Dutch banking houses Hope & Comp., Ketwich & Voombergh, and Weduwe Willem Borski (Widow Willem Borski). Interest at 5% per annum, payable on April 1 and October 1, administered through the Notariskantoor Commelin & Weyland. The last coupon shown is April 1, 1839. These Dutch certificates facilitated retail Dutch investment in Russian government silver bonds without direct ownership of the underlying Russian instruments, a common Amsterdam intermediation practice. Hope & Company was one of Europe\'s most prominent banking houses and played a central role in placing Russian government debt in Western markets.',
  {
    type: 'Certificate',
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    creator: 'Hope & Comp.; Ketwich & Voombergh; Weduwe Willem Borski',
    issueDate: '1834-08-29',
    currency: 'RUB',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Certificaat der Vijfde Reeks, 5% Russische Fondsen in Zilver. No. 13, 500 Roubles. Amsterdam, August 29, 1834. Issued by Hope & Comp., Ketwich & Voombergh, Weduwe Willem Borski. Interest 5% p.a., April 1 and October 1. Last coupon April 1, 1839.',
  }
);

// --- Row 487: Dutch Certificate No. 1475 for Swedish Crown Debt in Holland, Amsterdam, October 10, 1816 ---
setDoc(487,
  'Certificaat No. 1475: Dutch Certificate for Swedish Crown Debt Negotiated in Holland (Amsterdam, October 10, 1816)',
  'A Dutch certificaat No. 1475 evidencing the holder\'s position as a creditor in the Swedish Crown (Zweedsche Kroon) debt negotiated in Holland. The holder, W.G. van de Poll, has paid the first three installments toward repayment of debts owed by His Majesty the King of Sweden to Dutch creditors, with an agreed schedule to pay a fixed percentage of the nominal capital per year. Nominal capital of approximately 3,000 Dutch guilders remains partly outstanding. The certificate also records that remaining portions of the Swedish government obligations (Lijfrenten en Obligatiën ten laste de Zweedsche Kroon) are to be distributed among creditors. Issued in Amsterdam on October 10, 1816, and signed by W.G. van de Poll. Reflects the post-Napoleonic restructuring of Swedish sovereign debt held by Dutch investors.',
  {
    type: 'Certificate',
    subjectCountry: 'Sweden',
    issuingCountry: 'Netherlands',
    creator: 'W.G. van de Poll',
    issueDate: '1816-10-10',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Certificaat No. 1475. Dutch certificate for Swedish Crown debt in Holland. Holder: W.G. van de Poll. First three installments paid; remaining nominal capital ~3,000 Dutch guilders. Amsterdam, October 10, 1816. Post-Napoleonic Swedish debt restructuring.',
  }
);

// --- Row 488: Imperial Chinese Government Gold Loan of 1908, £20 Bond, No. B106822, Paris, March 15, 1908 ---
setDoc(488,
  'Imperial Chinese Government Gold Loan of 1908: Bond for £20 Sterling / Obligation de £20 Sterling (No. B106822, Paris, March 15, 1908)',
  'A bilingual English/French bearer bond No. B106822 for £20 Sterling (Obligation de £20 Sterling) issued by the Imperial Chinese Government under the Gold Loan of 1908 (Emprunt Or de 1908), with a total issue of £5,000,000 Sterling. Issued in Paris on March 15, 1908, through the Banque de l\'Indo-Chine. Printed in green with an ornate decorative border and vignette showing Chinese harbor and railway scenes. The loan was backed by the revenues of the Chinese Imperial Maritime Customs and other government revenues. Interest and principal payable in London, Paris, Berlin, and other European financial centers. Part of the series of late Qing dynasty foreign loans; the 1908 Gold Loan financed railway construction and government expenditures during a period of rapid Chinese modernization effort.',
  {
    type: 'Bond',
    subjectCountry: 'China',
    issuingCountry: 'France',
    creator: 'Imperial Chinese Government; Banque de l\'Indo-Chine',
    issueDate: '1908-03-15',
    currency: 'GBP',
    language: 'English, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Imperial Chinese Government Gold Loan 1908. £5,000,000 total. Bond No. B106822, £20 Sterling. Paris, March 15, 1908. Banque de l\'Indo-Chine. Bilingual English/French. Green ornate border. Backed by Chinese Imperial Customs revenues. Late Qing railway finance.',
  }
);

// --- Row 489: Hollandsche Garantie- & Trust Compagnie, 6% Certificaat, Reichsschuldbuch, 2000 Reichsmark, Amsterdam, November 1928 ---
setDoc(489,
  'Hollandsche Garantie- & Trust Compagnie: 6% Certificaat van Inschrijving in het Grootboek van het Duitsche Rijk (Reichsschuldbuch), 2000 Reichsmark (No. 0128, Amsterdam, November 1928)',
  'A 6% Certificate of Registration in the German National Debt Register (Groot Boek van het Duitsche Rijk / Reichsschuldbuch), No. 0128, for 2,000 Reichsmark, Series per 1943 (maturing March 31, 1943). Issued by the Hollandsche Garantie- & Trust Compagnie in Amsterdam, November 1928. The certificate represents a 6% inscription in the Reichsschuldbuch, registered in the name of the Hollandsche Garantie- & Trust Compagnie as trustee for Dutch investors, with conditions authorized and witnessed by Notary B.H. Grepels. Orange Dutch tax stamp visible. Issued during the Weimar Republic\'s period of currency stabilization, when the German government attracted foreign capital through long-term bond issues following the hyperinflation crisis. This Dutch-issued certificate intermediated German sovereign debt for retail Dutch investors.',
  {
    type: 'Certificate',
    subjectCountry: 'Germany',
    issuingCountry: 'Netherlands',
    creator: 'Hollandsche Garantie- & Trust Compagnie',
    issueDate: '1928-11-01',
    currency: 'German Marks',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: '6% Certificaat van Inschrijving, Grootboek Duitsche Rijk (Reichsschuldbuch). No. 0128, 2,000 Reichsmark, Series per 1943. Amsterdam, November 1928. Hollandsche Garantie- & Trust Compagnie. Notary B.H. Grepels. Dutch investment vehicle for Weimar-era German government bonds.',
  }
);

// --- Row 490: Royal Dutch Petroleum Company, Warrant No. 015284, Series per 1943, The Hague, February 1937 ---
setDoc(490,
  'Koninklijke Nederlandsche Maatschappij tot Exploitatie van Petroleumbronnen in Nederlandsch-Indië: Warrant No. 015284 (Series per 1943, \'s-Gravenhage, February 1937)',
  'A stock warrant (Warrant) No. 015284, Series per 1943, of the Koninklijke Nederlandsche Maatschappij tot Exploitatie van Petroleumbronnen in Nederlandsch-Indië (Royal Dutch Company for the Exploitation of Petroleum Sources in Dutch East Indies), established in The Hague. The holder has the right until March 31, 1940 to purchase one share at f.1,000 nominal value from the above company, or from April 1, 1940 until March 31, 1943 against payment of f.500. Valid until March 31, 1943 (van onwaarde na 31 Maart 1943). Issued \'s-Gravenhage, February 1937. Signed by two directors. Royal Dutch was the Dutch predecessor entity of the Royal Dutch/Shell oil company. This warrant was part of the company\'s equity capital-raising activities in the late 1930s, reflecting the complex financial structure of the early petroleum industry.',
  {
    type: 'Warrant',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Koninklijke Nederlandsche Maatschappij tot Exploitatie van Petroleumbronnen in Nederlandsch-Indië (Royal Dutch)',
    issueDate: '1937-02-01',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '20th Century',
    notes: 'Royal Dutch Petroleum Company warrant No. 015284, Series per 1943. Right to buy 1 share: until March 31, 1940 at f.1,000; from April 1, 1940 at f.500. Valid to March 31, 1943. \'s-Gravenhage, February 1937. Royal Dutch/Shell predecessor. Green and orange.',
  }
);

// --- Row 491: Compagnie des Indes, Legal Document re Share Transfer and Lottery Proceeds, Paris, January 13, 1745 ---
setDoc(491,
  'Compagnie des Indes: Legal Authorization Document for Share Transfer and Lottery Proceeds, Paris (January 13, 1745)',
  'A printed and handwritten legal document dated January 13, 1745, Paris, relating to the French East India Company (Compagnie des Indes). The document references Gabriel-Jérôme de Bullion (Chevalier, Maître de Camp de Régiment de Prusse Infanterie, Conseiller du Roy en son Conseil, Prévôt de la Ville, Prévôté et Vicomté de Paris), acting as Directeur de ladite Compagnie. The text references the King\'s Council edict of February 3, 1724 (re-establishing the reorganized Compagnie des Indes), and concerns the handling of Loterie de Saint-Quentin proceeds and share-related financial arrangements. Multiple named parties are involved, including Dame Juliane Elizabeth de Guise Holler and her husband or agent, regarding the sale, transfer, and remuneration of Company shares and related financial instruments under the Compagnie des Indes charter. A significant document illustrating the legal and financial complexity of French East India Company share transactions in the mid-18th century.',
  {
    type: 'Legal Document',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'Compagnie des Indes (French East India Company)',
    issueDate: '1745-01-13',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century',
    notes: 'Compagnie des Indes legal document, January 13, 1745, Paris. Director: Gabriel-Jérôme de Bullion. Concerns share transfers and Loterie de Saint-Quentin proceeds. References King\'s Council edict February 3, 1724. Parties include Dame Juliane Elizabeth de Guise Holler.',
  }
);

// --- Row 492: Liverpool Corn Trade Association, C. Perpetual Mortgage Debenture No. 291, £100 – Conditions and Transfer Record ---
setDoc(492,
  'Liverpool Corn Trade Association, Limited: C. Perpetual Mortgage Debenture No. 291, £100 – Conditions and Transfer Endorsements',
  'The reverse side of C. Perpetual Mortgage Debenture No. 291 of the Liverpool Corn Trade Association, Limited, showing the printed conditions (The Conditions within referred to) governing the debenture, and multiple handwritten transfer endorsements recording the ownership history. Transfer records note: transfer dated May 25, 1910 to Ernest Clifford Grasman at 58 Crofton Road, Essex County, registered in the books of the Association in June 1919; a further transfer dated October 9, 1933. The conditions detail the debenture\'s terms, including interest payment provisions, security over the Association\'s property, and redemption terms. This reverse side complements the front-face certificate (No. 291), together constituting the complete debenture instrument.',
  {
    type: 'Bond',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'Liverpool Corn Trade Association, Limited',
    issueDate: '1897-06-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Reverse of C. Perpetual Mortgage Debenture No. 291, £100, Liverpool Corn Trade Association, Ltd. Shows printed conditions and transfer record: transferred May 25, 1910 to Ernest Clifford Grasman; registered June 1919; further transfer October 9, 1933.',
  }
);

// --- Row 493: Liverpool Corn Trade Association, C. Perpetual Mortgage Debenture No. 291, £100, Liverpool, June 1897 ---
setDoc(493,
  'Liverpool Corn Trade Association, Limited: C. Perpetual Mortgage Debenture No. 291, £100 at 3½% (Liverpool, June 1897)',
  'C. Perpetual Mortgage Debenture No. 291 of the Liverpool Corn Trade Association, Limited, incorporated under the Companies Acts 1862 to 1893. Capital of the Association: £60,000 in 400 Shares of £150 each. Registered Office: 8, Brunswick Street, Liverpool. Issued under authority of Clause 54 of the Articles of Association and a Board resolution dated May 26, 1897, as part of an issue of £100,000 C. Perpetual Mortgage Debentures carrying interest at £3½ per centum per annum, payable on February 1 and August 1. The Association charges all its property, estate, and effects, present and future, as security. The debenture is transferable by endorsement. Given under the Common Seal of the Association, June 1897. Signed by the Chairman, a Director, and the Secretary (Edward Graham). The Liverpool Corn Trade Association was a leading institution in Liverpool\'s commodities market.',
  {
    type: 'Bond',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'Liverpool Corn Trade Association, Limited',
    issueDate: '1897-06-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'C. Perpetual Mortgage Debenture No. 291, £100 at 3½% p.a. Liverpool Corn Trade Association, Ltd. Issue of £100,000. Capital £60,000 in 400 shares of £150. 8 Brunswick Street, Liverpool. Common Seal, June 1897. Secretary: Edward Graham.',
  }
);

// --- Row 494: Dutch Notarial London Certificate for Bank of England 3% Consolidated Annuities, ca. late 18th century ---
setDoc(494,
  'London Certificatie: Dutch Notarial Certificate for Bank of England 3% Consolidated Annuities (London, ca. late 18th Century)',
  'A handwritten Dutch notarial certificate (London Certificatie) by Pieter Hendrik Hoogenbergh, a Dutch public notary authorized and admitted in London. The notary certifies, at the request of the named holder, that the individual is duly registered in the Grand Ledger of Transfers and Accounts of Shares in the Consolidated Capital of Three Per Cent Bank Annuities (Drie Per Cent Bank Annuïteiten) at the Bank of England, in the holder\'s own name. Authenticated with a red wax seal bearing the notary\'s heraldic device. Concluded "In Testimonium Veritatis" with the notary\'s signature. Not. Pub. No. 7. Issued in London, anno Seventien Honderd [?]. This type of certificate was routinely used by Dutch investors in British government consols (3% Consolidated Annuities) to certify and authenticate their holdings for purposes of Dutch probate, tax, or legal proceedings.',
  {
    type: 'Certificate',
    subjectCountry: 'United Kingdom',
    issuingCountry: 'United Kingdom',
    creator: 'Pieter Hendrik Hoogenbergh (Notary Public, London)',
    issueDate: '1790-01-01',
    currency: 'GBP',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century',
    notes: 'Dutch notarial London Certificate. Pieter Hendrik Hoogenbergh, Not. Pub. No. 7. Certifies holder is registered at Bank of England in 3% Consolidated Bank Annuities. Red wax seal. "In Testimonium Veritatis." London, ca. late 18th century.',
  }
);

// --- Row 495: United States of America, Bill of Exchange No. 57, $60 / 300 Livres Tournois, November 21, 1778 ---
setDoc(495,
  'United States of America: Bill of Exchange No. 57, $60 (300 Livres Tournois) for Interest on Loan Office Certificates (November 21, 1778)',
  'A printed and handwritten bill of exchange No. 57, issued by the United States of America on November 21, 1778. At Thirty Days Sight, payable to Jeremiah Green Esq. (or Order) in Three Hundred Livres Tournois (equivalent to Sixty Dollars), for Interest due on Money borrowed by the United States. Countersigned by Nath. Appleton, Commissary of the Loan Office in the State of Massachusetts Bay. Signed by H. Atkinson, Treasurer of Loans. These bills of exchange were issued by the Continental Congress as a mechanism for paying interest to domestic holders of Loan Office Certificates, allowing creditors to receive their interest in the form of bills that could be negotiated in France, reflecting the U.S. government\'s reliance on French credit during the Revolutionary War.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Massachusetts Bay',
    issueDate: '1778-11-21',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 57, $60 / 300 Livres Tournois. November 21, 1778. Payable to Jeremiah Green Esq. Interest on Loan Office Certificates. Countersigned by Nath. Appleton, Commissary; signed by H. Atkinson, Treasurer of Loans. Massachusetts Bay. Continental Revolutionary War finance.',
  }
);

// --- Row 496: United States of America, Bill of Exchange No. 814, $24 / 120 Livres Tournois, November 24, 1778 ---
setDoc(496,
  'United States of America: Bill of Exchange No. 814, $24 (120 Livres Tournois) for Interest on Loan Office Certificates (November 24, 1778)',
  'A printed and handwritten bill of exchange No. 814, issued by the United States of America on November 24, 1778. At Thirty Days Sight, payable to Timothy Newell Esq. (or Order) in One Hundred and Twenty Livres Tournois (equivalent to Twenty-four Dollars), for Interest due on Money borrowed by the United States. Countersigned by Nath. Appleton, Comptroller of the Commissioner of the Great Loan-Office in the State of Massachusetts Bay. Signed by H. Atkinson, Treasurer of Loans. Part of the same series as Bill No. 57 (goetzmann0495), issued three days later to a different payee. Illustrates the systematic use of Livres Tournois-denominated bills of exchange to service domestic Revolutionary War debt through French currency instruments.',
  {
    type: 'Bill of Exchange',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States of America, Loan Office, Massachusetts Bay',
    issueDate: '1778-11-24',
    currency: 'Livres Tournois',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'US Bill of Exchange No. 814, $24 / 120 Livres Tournois. November 24, 1778. Payable to Timothy Newell Esq. Interest on Loan Office Certificates. Countersigned by Nath. Appleton, Comptroller; signed by H. Atkinson, Treasurer of Loans. Massachusetts Bay. Continental Revolutionary War finance.',
  }
);

// --- Row 497: State of Massachusetts Bay, State Note, £230, Payable to Elisha Curtis, January 1, 1780 ---
setDoc(497,
  'State of Massachusetts Bay: Interest-Bearing State Note, £230, Payable to Elisha Curtis (January 1, 1780)',
  'An ornately printed state note from the State of Massachusetts Bay, dated the First Day of January, A.D. 1780. Issued in behalf of the State, the subscriber (as Committee member) promises to pay Elisha Curtis, or his or her Order, the Sum of Two Hundred and Thirty Pounds, bearing interest at six per cent per annum, with interest payable on the first Mondays of January, April, July, and October each year. Issued pursuant to a Law of the State titled "An Act to create a suitable Fund" authorizing additional taxation to service state obligations. The principal is recorded as redeemable only by application in person or by attorney. Signed by William [Cranch] and a second Committee member, with a Treasurer\'s signature. A key instrument of Massachusetts war finance during the Revolutionary period, when the State issued a variety of interest-bearing certificates to fund its military and governmental expenditures.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Massachusetts Bay',
    issueDate: '1780-01-01',
    currency: 'Massachusetts Pounds',
    language: 'English',
    numberPages: 1,
    period: '18th Century',
    notes: 'State of Massachusetts Bay interest-bearing state note, £230. Payable to Elisha Curtis. January 1, 1780. Interest 6% p.a., payable quarterly. Pursuant to Act authorizing additional taxation. Signed by Committee and Treasurer. Revolutionary War Massachusetts state finance.',
  }
);

// --- Row 498: Mexican Five Per Cent Deferred Stock Bond with Coupon Sheet (London, mid-19th century) ---
setDoc(498,
  'Mexican Five Per Cent Deferred Stock: Bond with Full Coupon Sheet (London, mid-19th Century)',
  'A large printed bearer bond certificate for the Mexican Five Per Cent Deferred Stock, featuring an elaborate decorative border and the Mexican national eagle vignette at top center. The bond includes a complete attached coupon sheet on the right side. The Mexican 5% Deferred Stock was one of the principal classes of Mexican external debt created through the conversion and restructuring of Mexico\'s London-market obligations, administered through the Council of the United Mexican States in London. The "deferred" designation indicates that interest payments were suspended and would recommence upon specified conditions being met. The bond includes an extensive text detailing the terms of the loan, the conditions of payment, and the security offered. The Mexican Deferred Stock was traded on the London Stock Exchange and was the subject of prolonged default, diplomatic disputes, and multiple renegotiations throughout the mid-to-late 19th century.',
  {
    type: 'Bond',
    subjectCountry: 'Mexico',
    issuingCountry: 'United Kingdom',
    creator: 'Mexican Government; Council of the United Mexican States (London)',
    issueDate: '1851-01-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'Mexican 5% Deferred Stock. Large format bearer bond with Mexican eagle vignette, ornate border, and full attached coupon sheet. Part of Mexican external debt restructuring, London market, mid-19th century. Extensive bond indenture text included.',
  }
);

// --- Row 499: Share Subscription Contract, Generale Keijserlijke Indische Compagnie, Antwerp, April 30, 1608 ---
setDoc(499,
  'Share Subscription Contract for the Generale Keijserlijke Indische Compagnie (Antwerp, April 30, 1608)',
  'A handwritten share subscription and forward contract for the Generale Keijserlijke Indische Compagnie (General Imperial Indian Company), signed in Antwerp on April 30, 1608. Registered as No. 36. The subscriber promises, in exchange for a received cash premium (Contante premie), to deliver and transfer a specified number of shares (Actien) in the Generale Keijserlijke Indische Compagnie to the counterparty or their order by the first available transfer day at the Company, with conditions relating to transport costs, reimbursement of proceeds, and the value of the shares. The document bears an elaborate embossed seal at top center. Contracted under Antwerp law. A rare and historically significant early modern financial instrument documenting forward share trading in what appears to be a Habsburg-sponsored East India trading company based in the Spanish Netherlands, operating contemporaneously with the Dutch VOC (founded 1602). This contract captures some of the earliest documented equity derivatives activity in European financial history.',
  {
    type: 'Share Subscription Contract',
    subjectCountry: 'Belgium',
    issuingCountry: 'Belgium',
    creator: 'Generale Keijserlijke Indische Compagnie',
    issueDate: '1608-04-30',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '17th Century',
    notes: 'Share subscription/forward contract, Generale Keijserlijke Indische Compagnie. Antwerp, April 30, 1608. Registered No. 36. Subscriber promises to deliver shares against received cash premium. Habsburg Imperial East India Company, Spanish Netherlands. Among the earliest known equity derivatives documents. Contemporary with Dutch VOC (founded 1602).',
  }
);

// --- Row 500: State of New York Per Cent Stock, Walter Terry, Albany, October 10, 1815 ---
setDoc(500,
  'State of New York: Per Cent Stock Certificate, $300, Walter Terry of Connecticut (Albany, October 10, 1815)',
  'A printed and handwritten New York State stock certificate issued by the Comptroller\'s Office, Albany, New York, dated October 10, 1815. The certificate acknowledges that the People of the State of New York owe Walter Terry of the State of Connecticut the sum of Three Hundred Dollars, bearing interest per cent per annum, payable on the first of January and July each year, created pursuant to an Act of the Legislature. The stock is transferable in person or by attorney according to rules established for that purpose, recorded in the Comptroller\'s accounts. Stamped "Funded" in red ink. Signed by John Ely, Deputy Comptroller, and the New York State Comptroller. An example of early 19th-century American state public finance, issued in the period of New York\'s post-War of 1812 economic expansion.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of New York, Comptroller\'s Office',
    issueDate: '1815-10-10',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '19th Century',
    notes: 'New York State Per Cent Stock, $300. Walter Terry, State of Connecticut. Albany, October 10, 1815. Interest payable January 1 and July 1. Stamped "Funded" in red. Signed by John Ely, Dep. Comptroller.',
  }
);

// --- Row 501: Lollenpolder en Hassendael Negotiatie, Legal Agreement, Amsterdam, March 4, 1817 ---
setDoc(501,
  'Lollenpolder en Hassendael Negotiatie: Legal Agreement, 292 Shares (Amsterdam, March 4, 1817)',
  'A printed and handwritten legal notarial document passed before Notary Everard Cornelis Bordt, Amsterdam, dated March 4, 1817, concerning the financial and administrative affairs of the Lollenpolder en Hassendael polder investment negotiatie. Principal parties include Cornelis Freymerius (merchant, of C. Freymerius en Zijn), Thomas Cuming (acting as vice-agent for the polder company), and Werbard van Vloter Aanraede (broker/Makelaar) and Siek Jan Voombergh (merchant), both acting as Commissioners of the Negotiatie. The document addresses the qualifications and obligations of the holders of 292 shares (twee honderd vice-en-negentig Aandelen) in the Negotiatie, the management responsibilities of Commissioners, and conditions for administration of the polder\'s capital. References the Associatie Knip of Vader of Zijn. Registered in Amsterdam on February 21, 1817 at D.S. f.159.',
  {
    type: 'Contract',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'C. Freymerius en Zijn; Everard Cornelis Bordt (Notary)',
    issueDate: '1817-03-04',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '19th Century',
    notes: 'Lollenpolder en Hassendael Negotiatie legal agreement. Amsterdam, March 4, 1817. Notary Everard Cornelis Bordt. Parties: Cornelis Freymerius, Thomas Cuming (vice-agent), Werbard van Vloter Aanraede and Siek Jan Voombergh (Commissioners). 292 shares. Registered Amsterdam February 21, 1817.',
  }
);

// --- Row 502: Unilever N.V. Optiebewijs No. 86786, Rotterdam, July 1937 ---
setDoc(502,
  'Unilever N.V.: Optiebewijs (Option Certificate) No. 86786 (Rotterdam, July 1937)',
  'Option certificate (Optiebewijs) No. 86786 of Unilever N.V., established in Rotterdam. Printed in blue and white. The holder of eight such option certificates has the right, against surrender of these certificates, to acquire one ordinary share certificate (certificaat van gewoon aandeel) of Unilever N.V. with a nominal value of f. 100, at the closing price of f. 130½, valid until June 30, 1942 (van onwaarde na 30 Juni 1942). The terms include conditions linked to the Van den Bergh\'s Fabrieken N.V. 3½% Obligatielening. Dated Rotterdam, July 1937. Signed by the Secretary and a member of the Board of Directors (Lid van de Raad van Bestuur). This is a second example of the same series as No. 97790 (goetzmann0469), differing only in serial number, demonstrating the wide distribution of Unilever option certificates in this period.',
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
    notes: 'Optiebewijs No. 86786. Unilever N.V., Rotterdam. Eight certificates entitle holder to acquire 1 ordinary share (f.100 nominal) at f.130½; valid until June 30, 1942. Related to Van den Bergh\'s Fabrieken N.V. 3½% Obligatielening. Rotterdam, July 1937. Blue-printed. Same series as No. 97790 (goetzmann0469).',
  }
);

// --- Row 503: Negotiatie Concordia Res Parvae Crescunt, Reçu No. 54, Amsterdam, January 1902 ---
setDoc(503,
  'Negotiatie "Concordia Res Parvae Crescunt": Reçu No. 54, f. 500 Aandelen (Amsterdam, January 1902)',
  'A receipt (Reçu) No. 54 from the Dutch investment fund/negotiatie "Concordia Res Parvae Crescunt" (Latin motto: "Through harmony small things grow"), acknowledging the transfer (overgenomen) from the Heer [Martonokle?] of shares valued at f. 500 (Aandelen) in the above-named Negotiatie. The transfer is made in accordance with the resolution adopted at the Shareholders\' Meeting (Vergadering van Aandeelhouders) held on December 4, 1893. Amsterdam, January 1902. Signed for the Associatiekas (Association Treasury). Amount: f. 430.553. "Concordia Res Parvae Crescunt" was a widely used name for Dutch investment pools (negotiaties) in the 19th and early 20th centuries, typically investing in mortgages, bonds, and other fixed-income instruments on behalf of their shareholders.',
  {
    type: 'Receipt',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Negotiatie Concordia Res Parvae Crescunt',
    issueDate: '1902-01-01',
    currency: 'NLG',
    language: 'Dutch, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Negotiatie "Concordia Res Parvae Crescunt." Reçu No. 54. Transfer of f. 500 aandelen from [Martonokle?]. Per resolution Shareholders\' Meeting, December 4, 1893. Amsterdam, January 1902. Amount: f. 430.553. Signed for Associatiekas.',
  }
);

// --- Row 504: Negotiatie Concordia Res Parvae Crescunt, Reçu No. 54 (duplicate / variant copy) ---
setDoc(504,
  'Negotiatie "Concordia Res Parvae Crescunt": Reçu No. 54, f. 500 Aandelen – Duplicate Copy (Amsterdam, January 1902)',
  'A duplicate or variant copy of Reçu No. 54 from the Dutch investment negotiatie "Concordia Res Parvae Crescunt," acknowledging the transfer from the Heer [Martonokle?] of shares valued at f. 500 in the Negotiatie, per the resolution of the Shareholders\' Meeting of December 4, 1893. Amsterdam, January 1902. Amount: f. 430.553. Signed for the Associatiekas. Appears identical in content and format to goetzmann0503, suggesting this is a second (retained) copy of the same transaction receipt, as was common practice in Dutch negotiatie administration where multiple copies were produced for different parties or for archival purposes.',
  {
    type: 'Receipt',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'Negotiatie Concordia Res Parvae Crescunt',
    issueDate: '1902-01-01',
    currency: 'NLG',
    language: 'Dutch, French',
    numberPages: 1,
    period: '20th Century',
    notes: 'Negotiatie "Concordia Res Parvae Crescunt." Reçu No. 54 (duplicate copy). f. 500 aandelen, [Martonokle?]. Per resolution December 4, 1893. Amsterdam, January 1902. Amount: f. 430.553. Likely retained/archival copy of same transaction as goetzmann0503.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 485–504 (20 documents, batch16).');
