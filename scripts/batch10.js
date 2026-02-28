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

// --- Row 620: U.S. Treasury 3% Debt Certificate, $14,000, 1792 ---
setDoc(620,
  'United States Treasury: 3% Debt Certificate, No. 4693, $14,000 (Donato A. Burton of London, 1792)',
  "This handwritten and engrossed document is a United States Treasury 3% Debt Certificate, No. 4693, issued from the Register's Office on the 14th of June, 1792. It attests that the United States of America owes to Donato A. Burton of London, or their Assigns, the sum of Fourteen Thousand Dollars, bearing interest at three per Centum per Annum from the first day of April 1792, payable quarterly and subject to redemption by payment of the principal whenever provision shall be made therefor by Law. The debt is recorded in the Register's Office and is transferable only by appearance in person or by attorney at the proper office. The document is signed by Joseph Nourse as Register of the Treasury, with an attestation below from Clement Biddle, Notary Public for the Commonwealth of Pennsylvania in Philadelphia, certifying the authenticity of the instrument on the fourteenth day of June, 1792. This certificate represents one of the early U.S. federal debt instruments issued under Alexander Hamilton's funding system.",
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States. Treasury Department',
    issueDate: '1792-06-14',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "U.S. Treasury 3% Debt Certificate No. 4693. $14,000 issued to Donato A. Burton of London. Signed by Joseph Nourse, Register of the Treasury. Notarized by Clement Biddle, Notary Public, Philadelphia, June 14, 1792.",
  }
);

// --- Row 621: Compagnie de Colonisation Américaine, 100-acre land share ---
setDoc(621,
  "Compagnie de Colonisation Américaine: Action de 100 Acres de Terres en Virginie et Kentucki, Série B, No. A/234 (1300 Francs)",
  "This printed and engraved share certificate is an action (share) of the Compagnie de Colonisation Américaine (American Colonization Company), representing the ownership of 100 acres of land in the states of Virginia and Kentucky, United States. The certificate is from Série B, No. A/234, at a face value of 1300 Francs, issued to Monsieur Charles Guillaume Juste Jerome as proprietor of the 100 acres. The text explains that the land forms part of 1,500,000 acres belonging to the Company, under the guarantee of De Beers and Compagnie, a Paris banking house. The certificate carries extensive dividend coupon strips on both the left and right margins, each labeled 'Coupon d'Action, Série B, de la Compagnie de Colonisation Américaine' with sequential coupon numbers and dates. The central text sets out the terms of the share, including the Company's cultivation and colonization obligations, the 80-franc annual dividend target, and the shareholder's rights under French law. This document is an example of French speculative investment in American land companies during the 1790s.",
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'France',
    creator: 'Compagnie de Colonisation Américaine',
    issueDate: '1790-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Compagnie de Colonisation Américaine, Action Série B, No. A/234, 100 acres in Virginia and Kentucky, 1300 Francs. Issued to Charles Guillaume Juste Jerome. Guaranteed by De Beers et Compagnie, Paris. With dividend coupon strips.",
  }
);

// --- Row 622: French Emprunt Forcé de l'An 4 (Forced Loan, Year 4 = 1795-96) ---
setDoc(622,
  "France, Emprunt Forcé de l'An 4: Récépissé No. 3869, Commune No. 9 (1795-96)",
  "This printed document is a receipt (récépissé) for the French Republican Forced Loan of Year 4 (Emprunt Forcé de l'An 4), corresponding to 1795-96 on the Gregorian calendar. The upper portion is a formal receipt from the designated Receiver of the Commune, acknowledging payment of a sum 'en numéraire ou valeur représentative aux termes de la Loi' (in coin or representative value as specified by the Law of 19 Frimaire, Year 4). The document carries the heading 'DÉPARTEMENT... COMMUNE No. 9' and is numbered 3869. Below the receipt is a detachable sheet of nine coupon bonds labeled 'Emprunt Forcé de l'An 4,' each bearing signatures, indicating the various installment payments of the forced loan that were scheduled across the year. The Emprunt Forcé de l'An 4 was a compulsory capital levy imposed by the French Directory on wealthy citizens to fund the Revolutionary government, and represented one of the most significant fiscal measures of the mid-1790s.",
  {
    type: 'Government Bond',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'France. République française',
    issueDate: '1795-01-01',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Emprunt Forcé de l'An 4 (Forced Loan of Year 4 = 1795-96). Récépissé No. 3869, Commune No. 9. Includes detachable coupon strip sheet with 9+ coupon bonds. Issued pursuant to the Law of 19 Frimaire, Year 4.",
  }
);

// --- Row 623: Insurance Company of Pennsylvania stock transfer, 1795 ---
setDoc(623,
  "Insurance Company of the State of Pennsylvania: Stock Transfer Endorsement, John Oldden to Henry Schlaesman (October 30, 1795)",
  "This handwritten document is a stock transfer endorsement on the back of a share certificate of the Insurance Company of the State of Pennsylvania. The text reads: 'I John Oldden — Do hereby assign and transfer to Henry Schlaesman Two and a half — Shares in the Capital or Joint Stock of the Insurance Company of the State of Pennsylvania. In witness whereof I have hereunto set my Hand the Thirtieth day of October 1795.' The document is witnessed by Jo. Palmer, and signed by Alex. Elmslie as Attorney. The Insurance Company of the State of Pennsylvania, founded in 1794, was among the earliest joint-stock insurance companies in the United States and represents an important early chapter in American corporate finance. This transfer endorsement was typically written on the blank verso of the original share certificate and executed in order to formally record a change of ownership.",
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Insurance Company of the State of Pennsylvania',
    issueDate: '1795-10-30',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Insurance Company of the State of Pennsylvania. Transfer endorsement: John Oldden to Henry Schlaesman, 2.5 shares, October 30, 1795. Signed by Alex. Elmslie, Attorney. Witnessed by Jo. Palmer.",
  }
);

// --- Row 624: Insurance Company of Pennsylvania stock transfer, 1797 ---
setDoc(624,
  "Insurance Company of the State of Pennsylvania: Stock Transfer Certificate, Henry Philips to Alexander Wiltocks (March 20, 1797)",
  "This handwritten document is a stock transfer certificate for the Insurance Company of the State of Pennsylvania, combining a ledger record on the left with the formal assignment text on the right. The left portion contains a tabular record with columns for Date, Folio, Ledger, Transferer and Transferee, Number of Shares, and Number of Certificate. The entry records the date March 20, 1797, a transfer from Henry Philips (Folio 228) to Alexander Wiltocks (Folio 135), for 3 shares. The right portion of the document reads: 'I Henry Philips Do hereby assign and transfer to Alexander Wiltocks Three — Shares in the Capital or Joint Stock of the Insurance Company of the State of Pennsylvania. In witness whereof I have hereunto set my Hand the Twentieth day of March 1797.' The document is witnessed by John Lewis and signed by Henry Philips. This certificate illustrates the early mechanisms of equity transfer used in American corporate practice at the close of the eighteenth century.",
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Insurance Company of the State of Pennsylvania',
    issueDate: '1797-03-20',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Insurance Company of the State of Pennsylvania. Transfer: Henry Philips to Alexander Wiltocks, 3 shares, March 20, 1797. Witnessed by John Lewis. Includes ledger record (Folio 228→135).",
  }
);

// --- Row 625: Massachusetts Bay Consolidated Note, £285, January 1, 1780 ---
setDoc(625,
  "Massachusetts Bay: Consolidated Note, No. 287, £285 (Richard Fry, January 1, 1780)",
  "This printed and handwritten document is a Massachusetts Bay Consolidated Note, No. 287, dated the First Day of January, A.D. 1780, issued during the American Revolutionary War. The text reads: 'IN Behalf of the State of Massachusetts Bay, I the Subscriber do hereby promise and oblige Myself and Successors in the Office of TREASURER of said State, to pay unto Richard Fry or to his Order, the Sum of Two hundred and eighty five Pounds on or before the First Day of March... 1783, with Interest at Six per Cent. per Annum.' The note specifies that both principal and interest were to be paid in the then current money of the State, adjusted to compensate for wartime inflation—calibrated against the prices of specified commodities (corn, beef, sheeps wool, sole leather). The sum of £285 represents thirty-two and a half times the value of these commodities at 1777 prices. The note references an Act of the General Assembly of February 6, 1779, providing for Massachusetts' contribution to the Continental Army. The document is signed by Wm. Dawes and R. Cranch as Committee members, and by Hy. Jones as Treasurer.",
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Massachusetts Bay. Treasury',
    issueDate: '1780-01-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Massachusetts Bay Consolidated Note No. 287, £285, January 1, 1780. Issued to Richard Fry. 6% interest, due March 1, 1783. Commodity-indexed principal. Signed by Wm. Dawes, R. Cranch (Committee) and Hy. Jones, Treasurer.",
  }
);

// --- Row 626: U.S. Loan Office Certificate, $533.33, New York, 1790 ---
setDoc(626,
  "United States Loan Office: 3% Debt Certificate, No. 4513, $533.33 (Jean Francois Paul Grand of Switzerland, 1790)",
  "This handwritten and engrossed document is a United States Loan Office Certificate, No. 4513, issued at the State of New York Loan Office on the 17th of October, 1790. The document reads: 'BE it known that there is due from the United States of America unto Jean Francois Paul Grand of Switzerland or his Assigns, the Sum of Five hundred & thirty three Dollars & thirty three Cents As a Debt bearing Interest at Three per Centum per Annum, from the first day of Oct. 1790 inclusively, payable quarter-yearly, and subject to Redemption by Payment of said Sum whenever Provision shall be made therefor by Law.' The certificate is signed by M. Clarkson as Commissioner of Loans. An attestation follows from John Wilkes, Notary Public of the State of New York, certifying the document's authenticity. The document bears the United States arms stamp and was issued to Jean Francois Paul Grand, a member of the prominent Swiss-French banking family Grand & Co. who were major holders of early U.S. federal debt. It represents an early instance of European investment in the new American republic's funded debt.",
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'United States. Treasury Department',
    issueDate: '1790-10-17',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "U.S. Loan Office Certificate No. 4513, New York, $533.33 at 3%. Issued to Jean Francois Paul Grand of Switzerland, October 17, 1790. Signed by M. Clarkson, Commissioner of Loans. Notarized by John Wilkes, Notary Public, New York.",
  }
);

// --- Row 627: French Royal Rente Receipt, Garde du Trésor Royal, 1698 ---
setDoc(627,
  "France, Garde du Trésor Royal: Quittance for 8,000 Livres (400 Livres de Rente sur l'Hôtel de Ville de Paris, April 10, 1698)",
  "This handwritten document is a receipt (quittance) issued on April 10, 1698, by Jean de Turmenyes, Conseiller du Roi and Garde du Trésor Royal (Keeper of the Royal Treasury), acknowledging receipt from Jean Eloy Udtin, argentier (financial agent) of the Archbishop of Rouen, of eight thousand livres in gold louis, silver, and other coin. The payment constitutes the principal for four hundred livres of annual perpetual rente, to be constituted by the Prévôts des Marchands et Échevins (Provost of Merchants and Aldermen) of the City of Paris on the one million livres of annual and perpetual rente newly alienated by His Majesty pursuant to the Royal Edit of March 1698, registered as required. The rente is to be drawn from the revenues of the Aides et Gabelles (excise and salt taxes) of France, for the benefit of Udtin at the rate of the Denier Vingt (5% per annum). The document is signed by 'de Turmenyes' and captioned 'Quittance du Garde du Trésor Royal, Année mil six cent quatre-vingt-dix-huit.' It represents one of the early modern instruments of French royal debt finance, linking crown fiscal necessity to municipal rente issuance.",
  {
    type: 'Bond',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: 'France. Trésor Royal',
    issueDate: '1698-04-10',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Quittance du Garde du Trésor Royal, No. 391, April 10, 1698. 8,000 livres principal for 400 livres of annual rente on the Hôtel de Ville de Paris (Aides et Gabelles). Issued to Jean Eloy Udtin. Signed by Jean de Turmenyes, Garde du Trésor Royal. Pursuant to Edit of March 1698.",
  }
);

// --- Row 628: Piscataqua Bridge share certificate, Town of Portsmouth, 1793 ---
setDoc(628,
  "Piscataqua Bridge Corporation: Share Certificate No. 262, Town of Portsmouth, New Hampshire (December 7, 1793)",
  "This printed and handwritten document is a share certificate of the Piscataqua Bridge Corporation, certifying that the Town of Portsmouth in the County of Rockingham, in the State of New Hampshire, is the proprietor of Share Number Two hundred sixty two in the Piscataqua Bridge. The certificate states that the share is transferable by making an assignment on the back and causing the same to be recorded on the Proprietors' Records. The document is dated the seventh day of December, One thousand seven hundred and ninety three, and is signed by James Sheafe as President, by a second officer as Treasurer, and attested by Nathaniel Adams as Proprietors' Clerk. The Piscataqua Bridge was a privately chartered toll bridge spanning the Piscataqua River between Portsmouth, New Hampshire and Kittery, Maine, one of the earliest bridge corporations in the United States. The certificate illustrates the early American practice of financing public infrastructure through joint-stock corporations with transferable shares.",
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Piscataqua Bridge Corporation',
    issueDate: '1793-12-07',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Piscataqua Bridge Corporation, Share No. 262, Town of Portsmouth, Rockingham County, New Hampshire, December 7, 1793. President: James Sheafe. Proprietors' Clerk: Nathaniel Adams. Early American toll bridge infrastructure share.",
  }
);

// --- Row 629: Action de la Caisse d'Épargnes et de Bienfaisance Lafarge, Paris, 1798 ---
setDoc(629,
  "Caisse d'Épargnes et de Bienfaisance Lafarge, Seconde Société: Action No. 124824 (90 Livres, Paris, September 30, 1798)",
  "This printed and handwritten document is an action (share) of the Caisse d'Épargnes et de Bienfaisance Lafarge, Seconde Société (Lafarge Savings and Charitable Society, Second Society), No. 124824, Paris. The document records that Paul Beauttier and Catherine Ferraudin, his wife, have paid to the Caisse the sum of ninety livres as the price of the present action, conferring on the life of Catherine Ferraudin the right to participate in the rentes and accroissements (growth) of the fund up to a maximum of three thousand livres of rente per annum, to be established by the certificat de vie (certificate of life). The document is dated at Paris, the thirtieth of September 1798 (or the nearby period), and bears the signatures of the Directeur Général and the Inspecteur Général, along with registration stamps confirming it was registered under No. 10317. A rubric of rules for the institution appears at the foot of the document. The Caisse d'Épargnes Lafarge was a late-eighteenth-century Parisian mutual savings institution offering tontine-style life annuities, an early precursor to modern insurance and retirement savings schemes.",
  {
    type: 'Tontine',
    subjectCountry: 'France',
    issuingCountry: 'France',
    creator: "Caisse d'Épargnes et de Bienfaisance Lafarge",
    issueDate: '1798-09-30',
    currency: 'FRF',
    language: 'French',
    numberPages: 1,
    period: '18th Century or before',
    notes: "Caisse d'Épargnes et de Bienfaisance Lafarge, Seconde Société. Action No. 124824, 90 livres, Paris, September 30, 1798. Issued to Paul Beauttier and Catherine Ferraudin (on the life of Catherine Ferraudin). Tontine-style life annuity savings instrument. Registration No. 10317.",
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 620-629 (10 individual documents).');
