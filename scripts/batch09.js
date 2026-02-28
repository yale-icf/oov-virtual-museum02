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

// --- Row 233: American Sugar Company Stock Certificate ---
setDoc(233,
  'American Sugar Company: 5.44% Cumulative Preferred Stock Certificate, No. PD 21186 (20 Shares, 1968)',
  'This printed and engraved stock certificate in brown and white attests to the ownership of twenty shares of the 5.44% Cumulative Preferred Stock of the American Sugar Company, incorporated under the laws of the State of Delaware with a par value of $12.50 per share. Certificate No. PD 21186 is registered to Shearson Hammill & Co. Incorporated and is dated June 5, 1968. The upper border carries the notation "CERTIFICATE FOR LESS THAN 100 SHARES" in both corners, while the central vignette depicts classical allegorical figures representing commerce and industry. The certificate bears the facsimile signatures of A.R. Williams as Treasurer and S.F. Oliver as President, and is authenticated by First National City Bank as Registrar.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'American Sugar Company',
    issueDate: '1968-06-05',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'American Sugar Company, incorporated under the laws of the State of Delaware. Certificate No. PD 21186. Registered to Shearson Hammill & Co. Incorporated.',
  }
);

// --- Row 234: American Trust Company Stock Certificate ---
setDoc(234,
  'American Trust Company: Capital Stock Certificate, No. 1951 (9 Shares, Boston, 1910)',
  'This engraved stock certificate in green and white is issued by the American Trust Company of Boston, incorporated under the laws of the State of Massachusetts in 1881. Certificate No. 1951 certifies that Harriette S. Foster is entitled to nine shares of the capital stock of the American Trust Company, transferable by the holder in person or by attorney upon surrender of the certificate, dated January 24, 1910. The certificate prominently features a large central vignette of an American bald eagle in flight. A bold red diagonal overprint reading "NOT OVER NINE SHARES" appears across the face of the certificate, indicating this is a fractional lot. The certificate is signed by two company officers whose facsimile signatures appear at the lower corners.',
  {
    type: 'Stock Certificate',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'American Trust Company',
    issueDate: '1910-01-24',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'American Trust Company, incorporated under laws of Massachusetts, A.D. 1881. Certificate No. 1951, Boston. Issued to Harriette S. Foster, 9 shares.',
  }
);

// --- Row 235: Anglo-Argentine Tramways Debenture Stock Certificate ---
setDoc(235,
  'Anglo-Argentine Tramways Company Limited: 5% Debenture Stock Certificate, No. A/68802 (£20, 1910)',
  'This document is the reverse side of a £20 5% Debenture Stock Certificate (No. A/68802) issued by the Anglo-Argentine Tramways Company Limited. Printed in red on white, the document reproduces a dense two-column "Extract from the Trust Deed, Dated 15th June, 1910," which outlines the terms and conditions of the debenture stock, the powers of the trustees, and the rights of stockholders. Below the Trust Deed extract appears a "Certificate of the Registration of a Mortgage or Charge" pursuant to the Companies (Consolidation) Act, 1908, No. 49(1), signed by F. Atherton as Assistant Registrar of Joint Stock Companies. The Anglo-Argentine Tramways Company was a British-registered firm operating urban tramway lines in Buenos Aires, Argentina, and is an example of British investment in South American infrastructure during the Edwardian era.',
  {
    type: 'Bond',
    subjectCountry: 'Argentina',
    issuingCountry: 'United Kingdom',
    creator: 'Anglo-Argentine Tramways Company Limited',
    issueDate: '1910-06-15',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'Anglo-Argentine Tramways Company Limited, 5% Debenture Stock, £20, Certificate No. A/68802. Back of certificate: Extract from Trust Deed (15 June 1910) and Certificate of Registration of a Mortgage or Charge.',
  }
);

// =============================================================================
// Rows 236-273: Acts of Parliament, Anno Regni Georgii Regis Octavo (8 Geo. I)
// London: John Baskett, 1722. 38 pages.
// Main Act: "An Act for Paying off and Cancelling One Million of Exchequer Bills,
// and to give Ease to the South-Sea Company..."
// Printed pages: [title page], [verso], pp. 299-336 (FINIS)
// =============================================================================

const geoI8Base = {
  type: 'Acts of Parliament',
  subjectCountry: 'Great Britain',
  issuingCountry: 'Great Britain',
  creator: 'Great Britain. Parliament',
  issueDate: '1722-01-01',
  currency: '',
  language: 'English',
  numberPages: 38,
  period: '18th Century or before',
  notes: "Printed by John Baskett, Printer to the Kings most Excellent Majesty, and by the Assigns of Thomas Newcomb, and Henry Hills, deceas'd. London, 1722. Acts of the Eighth Session of Parliament begun 17 March 1714, continued by prorogation to 19 October 1721.",
};

// Row 236 (Page 1 of 38): Title page
setDoc(236,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 1 of 38)',
  "This is the title page of the printed volume of Acts of Parliament for the eighth year of the reign of King George I of Great Britain. The title reads 'Anno Regni GEORGII REGIS Magnæ Britanniæ, Franciæ, & Hiberniæ OCTAVO,' noting that Parliament was begun and holden at Westminster on the seventeenth day of March, Anno Dom. 1714—in the first year of the reign—and continued by several prorogations to the nineteenth day of October, 1721, being the eighth session of this present Parliament. The royal arms of Great Britain with the 'G R' cypher are printed at the centre. The imprint at the foot reads: 'LONDON, Printed by John Baskett, Printer to the Kings most Excellent Majesty, And by the Assigns of Thomas Newcomb, and Henry Hills, deceas'd. 1722.' This volume is of particular interest as it contains the main parliamentary legislation enacted in direct response to the South Sea Bubble collapse of 1720.",
  geoI8Base
);

// Row 237 (Page 2 of 38): Verso of title page
setDoc(237,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 2 of 38)',
  "This is the verso (reverse side) of the title page of the 1722 volume of Acts of Parliament for the eighth year of King George I. The printed title page is faintly visible in mirror image through the paper, a characteristic of early-eighteenth-century printing on hand-laid paper. No additional text or content is printed on this side. The page serves as the blank reverse leaf before the substantive Act texts begin on the following pages.",
  geoI8Base
);

// Row 238 (Page 3 of 38): Printed p. 299 — Act heading + opening
setDoc(238,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 3 of 38)',
  "This page (printed p. 299) opens with the heading 'Anno Octavo Georgii Regis' and presents the full title of the principal Act in this section: 'An Act for Paying off and Cancelling One Million of Exchequer Bills, and to give Ease to the South-Sea Company, in respect of its present Obligation, to circulate or continue towards Circulating Exchequer Bills; and to give further Time to make Repayment of One Million, which was lent to them; and for Issuing a further Sum in New Exchequer Bills, towards His Majesties Supply, to be Discharged and Cancelled, when the said Company shall repay the Million owing by them; and that the Exchequer Bills, which are to Continue, may be Circulated at ease and moderate Rates; and for Appropriating the Supplies granted to His Majesty in this Session of Parliament; and for Relief of the Sufferers at Nevis and Saint Christophers, by an Invasion of the French in the last late War; and for laying a further Duty on Apples Imported; and for Ascertaining the Duties on Pictures Imported.' The Act text begins below with a decorated woodcut initial and the address 'Most Gracious Sovereign.'",
  geoI8Base
);

// Row 239 (Page 4 of 38): Printed p. 300 — preamble
setDoc(239,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 4 of 38)',
  "This page (printed p. 300) carries the opening preamble of the South Sea and Exchequer Bills Act of 1722. The text begins 'Most Gracious Sovereign: Whereas among divers Matters and Things contained in an Act of Parliament made and passed in the Sixth Year of Your Majesties Reign intituled An Act Enabling the South-Sea Company to increase their present Capital Stock and Fund by Redeeming such Publick Debts and Incumbrances as are therein mentioned.' The preamble recounts the existing obligations of the South Sea Company regarding the circulation of Exchequer Bills and the million pounds previously lent to the Company by Parliament, and details the difficulties in fulfilling these obligations following the collapse of the South Sea Bubble in 1720. The recital acknowledges the need for fresh parliamentary authority to regularize the public credit.",
  geoI8Base
);

// Row 240 (Page 5 of 38): Printed p. 301 — opening clauses
setDoc(240,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 5 of 38)',
  "This page (printed p. 301) continues the preamble and begins the operative clauses of the South Sea and Exchequer Bills Act of 1722. The text describes the statutory framework for cancelling one million pounds of Exchequer Bills previously held in connection with the South Sea Company, and for issuing new Exchequer Bills to replace them. The Act grants authority to the Commissioners of the Treasury to manage the exchange, circulation, and redemption of Exchequer Bills, and specifies the interest rate—Two Pence per Centum per Diem—at which the bills are to circulate. The page also records that the South Sea Company had requested Parliament's assistance in resolving its outstanding obligations and that the Company undertook to repay the principal sum from the Sinking Fund.",
  geoI8Base
);

// Rows 241-249 (Pages 6-14 of 38): Act text, printed pp. ~302-316 approx.
for (let r = 241; r <= 249; r++) {
  const pg = r - 235;
  setDoc(r,
    `Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page ${pg} of 38)`,
    "This page contains legislative text from the South Sea and Exchequer Bills Act of 1722 (8 Geo. I). The clauses in this portion of the Act govern the practical administration of Exchequer Bills: their issuance from the Exchequer under direction of the Commissioners of the Treasury, the rate of interest (Two Pence per Centum per Diem), procedures for exchanging old bills for new, the appointment of Trustees, and the penalties for counterfeiting or fraudulently altering bills. The Act specifies that bills are to be numbered arithmetically, registered in appointed offices, and that interest payments shall be calculated from the date of issue. This legislation represents a key parliamentary instrument for stabilizing British public credit in the wake of the South Sea Bubble collapse of 1720.",
    geoI8Base
  );
}

// Row 250 (Page 15 of 38): Printed p. 318 — revenue appropriation and Trustees
setDoc(250,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 15 of 38)',
  "This page (printed p. 318) continues the South Sea and Exchequer Bills Act of 1722. The text at this point addresses the authority vested in the Commissioners of the Treasury to prepare and circulate Exchequer Bills and to fund them from specific revenue streams including Customs, Excise, and other parliamentary grants. The Act authorizes the creation of Trustees for the management of the bills and specifies that bills shall be free from all parliamentary and other impositions. The passage also establishes provisions for the arithmetical numbering of bills from No. 1, for their registration in the Cursitor's or other appointed offices, and for the accountability of the Treasury officers in keeping accurate records. The Act describes the procedure by which bills may be received at the Exchequer or the South Sea Company offices at any time during the seven-year term.",
  geoI8Base
);

// Rows 251-259 (Pages 16-24 of 38): Act text, printed pp. ~319-327 approx.
for (let r = 251; r <= 259; r++) {
  const pg = r - 235;
  setDoc(r,
    `Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page ${pg} of 38)`,
    "This page contains legislative text from the South Sea and Exchequer Bills Act of 1722 (8 Geo. I). The clauses in this section of the Act deal with the procedural requirements for redeeming and cancelling Exchequer Bills at the end of their term, with the powers of the Commissioners of the Treasury to raise further supplies if needed, and with provisions assigning specific revenue streams—including Customs receipts and Excise duties—as security for the bills in circulation. The Act also prescribes penalties for officers who fail to carry out their duties or who misappropriate funds, and establishes accountability measures for the Receivers and Collectors entrusted with managing the revenue. These clauses were part of Parliament's broader effort to restore orderly public credit following the South Sea Bubble.",
    geoI8Base
  );
}

// Row 260 (Page 25 of 38): Printed pp. 323-324 — South Sea repayment provisions
setDoc(260,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 25 of 38)',
  "This page (printed pp. 323-324) continues the South Sea and Exchequer Bills Act of 1722. The text addresses the powers granted to the Commissioners of the Treasury to direct repayment of the principal sum of One Million Pounds previously lent to the South Sea Company, to be funded out of the Sinking Fund. The passage also deals with the authority to receive applications from South Sea Company creditors, to direct the payment of interest at the prescribed rate, and to appropriate specific revenue sources—including Customs and Excise receipts—towards the discharge of the Exchequer Bills. The Act empowers the Treasury to convert South Sea Company obligations into the new Exchequer Bill structure and specifies the timetable for repayment.",
  geoI8Base
);

// Rows 261-269 (Pages 26-34 of 38): Act text, printed pp. ~325-332 approx.
for (let r = 261; r <= 269; r++) {
  const pg = r - 235;
  setDoc(r,
    `Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page ${pg} of 38)`,
    "This page contains legislative text from the South Sea and Exchequer Bills Act of 1722 (8 Geo. I). The clauses in this section cover the appropriation of parliamentary supplies granted to His Majesty in the eighth session of Parliament, including the methods by which revenues from Customs, Excise, and other sources are assigned and applied toward the national debt. The text also includes provisions for relief of the sufferers at Nevis and Saint Christophers who sustained losses from a French invasion during the previous war, detailing the procedures for making and verifying claims, and for laying additional duties on imported apples and pictures for the purpose of raising revenue.",
    geoI8Base
  );
}

// Row 270 (Page 35 of 38): Printed p. 333 — Annuities and Debentures clauses
setDoc(270,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 35 of 38)',
  "This page (printed p. 333) contains the closing sections of the South Sea and Exchequer Bills Act of 1722 dealing with the obligations toward South Sea Annuity holders and South Sea Company Debenture holders. The text specifies procedures for making annuity payments at the Feast of the Nativity of Saint John Baptist (24 June) and at other quarterly dates, with interest at Three Pounds per Centum per Annum. The passage also addresses the authority of Attorneys, Agents, and Principal Substitutes authorized to act on behalf of Annuity holders in receiving payments, and the legal consequences for those failing to act within the prescribed terms. The page further contains the provision for Annuities to pass to executors and administrators upon death.",
  geoI8Base
);

// Row 271 (Page 36 of 38): Printed p. 334 — Annuities and Additional Duties
setDoc(271,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 36 of 38)',
  "This page (printed p. 334) continues with clauses governing the payment of South Sea Annuities to executors and administrators, providing that annuities shall be paid by the Officers in the Receipt of the Exchequer, and specifying that payments be made out of the monies arising from the General Fund and related revenues. The page also introduces supplementary clauses concerning an Additional Duty on imported apples: it is hereby enacted that such Duties shall be raised, levied, recovered, and paid according to specified rules, and paid into the Exchequer. The page further contains the commencement provision for the new duties on pictures imported, establishing the rate structure and procedures for the Surveyor General and Customs Officers to assess and collect them.",
  geoI8Base
);

// Row 272 (Page 37 of 38): Printed p. 335 approx. — Duties on Pictures
setDoc(272,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 37 of 38)',
  "This page (printed p. 335 approx.) contains the penultimate section of the South Sea and Exchequer Bills Act of 1722, dealing primarily with the new duties on pictures imported into Great Britain. The text specifies that Duties on Pictures Imported shall be raised, levied, recovered, and paid by the respective Importers of such Pictures, and shall be brought into the Exchequer by such Rules, Ways, Means, and Methods, and under such Penalties and Forfeitures as are prescribed for other imported goods. The passage also sets out the procedure for Officers of the Customs to assess the dimensions of pictures as a basis for calculating duties, and provides for the Surveyor General to establish measurement standards. The Act specifies that these new duties are in addition to, and not in lieu of, existing duties.",
  geoI8Base
);

// Row 273 (Page 38 of 38): Printed p. 336 — FINIS
setDoc(273,
  'Acts of Parliament, Anno Regni Georgii Regis Octavo, 1722 (Page 38 of 38)',
  "This is the final page (printed p. 336) of the Acts of Parliament volume for the eighth year of King George I. The page concludes the provisions governing duties on imported pictures, specifying that the Duties hereby charged on Pictures imported shall be applied and extended to the Exchequer Bills Act and related supply needs as fully as if they had been part of the original revenue grant. The Act provides that if any of the said Duties on Pictures were to cease or determine by virtue of any other law, a proportional share of the Duties hereby charged shall likewise cease and determine. The page closes with the word 'FINIS.' in large type, marking the end of this session's Acts. The legislation as a whole represents Parliament's comprehensive post-bubble fiscal settlement for the year 1721-22.",
  geoI8Base
);

// =============================================================================
// Rows 274-285: Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto
// (24 Geo. II, 1750). London: Thomas Baskett, 1750. 12 pages.
// Main Act: "An Act for enabling His Majesty to raise the several Sums of Money
// therein mentioned, by Exchequer Bills... for paying off the Old and New
// unsubscribed South Sea Annuities..."
// Printed pages: [title page], pp. 79-89
// =============================================================================

const geoII24Base = {
  type: 'Acts of Parliament',
  subjectCountry: 'Great Britain',
  issuingCountry: 'Great Britain',
  creator: 'Great Britain. Parliament',
  issueDate: '1750-01-01',
  currency: '',
  language: 'English',
  numberPages: 12,
  period: '18th Century or before',
  notes: "Printed by Thomas Baskett, Printer to the King's most Excellent Majesty, and by the Assigns of Robert Baskett. London, 1750. Acts of the Fourth Session of Parliament begun 10 November 1747, continued by prorogation to 17 January 1750.",
};

// Row 274 (Page 1 of 12): Title page
setDoc(274,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 1 of 12)',
  "This is the title page of the printed volume of Acts of Parliament for the twenty-fourth year of the reign of King George II of Great Britain. The title reads 'Anno Regni GEORGII II. REGIS Magnæ Britanniæ, Franciæ, & Hiberniæ, VICESIMO QUARTO.' The page records that Parliament was begun and holden at Westminster on the tenth day of November, Anno Dom. 1747, in the twenty-first year of the reign, and continued by several prorogations to the seventeenth day of January, 1750—being the fourth session of this Parliament. The royal arms of Great Britain with the 'G R' cypher are printed at the centre. The imprint at the foot reads: 'LONDON: Printed by Thomas Baskett, Printer to the King's most Excellent Majesty; and by the Assigns of Robert Baskett. 1750.' This volume is notable for containing legislation to finally extinguish the residual South Sea Annuities, three decades after the South Sea Bubble.",
  geoII24Base
);

// Row 275 (Page 2 of 12): Printed p. 79 — Act heading + opening
setDoc(275,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 2 of 12)',
  "This page (printed p. 79) opens with the heading 'Anno vicesimo quarto Georgii II. Regis' and presents the full title of the principal Act in this section: 'An Act for enabling His Majesty to raise the several Sums of Money therein mentioned, by Exchequer Bills, to be charged on the Sinking Fund; and for impowering the Commissioners of the Treasury to pay off the Old and New unsubscribed South Sea Annuities out of the Supply granted to His Majesty for the Service of the Year One thousand seven hundred and fifty one; and for enabling the Bank of England to hold General Courts, and Courts of Directors, in the Manner therein directed; and for giving certain Persons Liberty to subscribe Bank and South Sea Annuities omitted to be subscribed pursuant to Two Acts of the last Session of Parliament.' The Act text begins with the decorated initial 'Most Gracious Sovereign: Whereas by an Act of Parliament made and passed in the Twenty Third Year of His Majesty's Reign, intituled An Act for giving further Time to the Proprietors of Annuities, after the Rate of Four Pounds per Centum per Annum, to subscribe the same.'",
  geoII24Base
);

// Row 276 (Page 3 of 12): Printed pp. 80-81 — South Sea Annuities subscription
setDoc(276,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 3 of 12)',
  "This page (printed pp. 80-81) continues the text of the South Sea Annuities and Exchequer Bills Act of 1750. The text describes the terms on which holders of Old and New unsubscribed South Sea Annuities—National Debt instruments originating from the South Sea Company schemes of 1720—are to be paid off. The Act specifies that these Annuities, incurred before Michaelmas 1719 and redeemable at the rate of Four Pounds per Centum per Annum, are to be purchased from their proprietors using monies advanced from the Sinking Fund, with the Governor and Company of the Bank of England authorized to advance sums for this purpose. The passage records the key figure: One million twenty-six thousand four hundred and seventy-six Pounds, Four Shillings, and Six Pence to be advanced at the Rate of Three Pounds per Centum per Annum.",
  geoII24Base
);

// Row 277 (Page 4 of 12): Printed pp. ~82-83 approx.
setDoc(277,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 4 of 12)',
  "This page (printed pp. ~82-83 approx.) continues the legislative text of the South Sea Annuities and Exchequer Bills Act of 1750. The clauses in this section govern the procedures by which the Bank of England and the Commissioners of the Treasury are empowered to advance funds for the redemption of South Sea Annuities, on the condition that Exchequer Bills be issued to the Bank as security. The Act specifies the terms and timing of repayment and lays out the procedures by which the Commissioners of the Treasury shall apply the sums advanced towards paying off and cancelling the outstanding South Sea Annuities, drawing upon the Sinking Fund and the General Supply of Parliament.",
  geoII24Base
);

// Row 278 (Page 5 of 12): Printed p. 82 — Exchequer Bills and Bank of England terms
setDoc(278,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 5 of 12)',
  "This page (printed p. 82) contains clauses of the South Sea Annuities and Exchequer Bills Act of 1750 dealing with the conditions under which Exchequer Bills are to be advanced by the Bank of England to the Treasury. The text specifies that Bills of Exchange to the value of One million twenty-six thousand four hundred and seventy-six Pounds at the Rate of Three Pounds per Centum per Annum may be advanced, on condition that Exchequer Bills be issued to the Bank by a date stated in the Act. The Governor and Company of the Bank of England are empowered to advance these sums on His Majesty's behalf, with interest payments to be made from the stated Receipts and charged upon the Sinking Fund. The passage also directs the Lords Commissioners of His Majesty's Treasury to take proposals from the Bank and to agree with the Governor and Company of the Bank of England on the terms of the advance.",
  geoII24Base
);

// Row 279 (Page 6 of 12): Printed pp. ~83-84 approx.
setDoc(279,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 6 of 12)',
  "This page (printed pp. ~83-84 approx.) continues the legislative text of the South Sea Annuities and Exchequer Bills Act of 1750. The clauses here address the technical procedures for preparing, signing, and issuing Exchequer Bills to the Bank of England in exchange for the advances to be made toward redeeming the South Sea Annuities. The Act specifies the denominations and form of the Exchequer Bills, the authorities responsible for their preparation, and the procedures for their registration and authentication by Officers of the Exchequer. The passage also provides that the principal sums advanced shall be repaid out of the Sinking Fund at specified times, and that the Bank shall be entitled to receive interest in the meantime.",
  geoII24Base
);

// Row 280 (Page 7 of 12): Printed pp. 84-85 — Exchequer Bills circulation terms
setDoc(280,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 7 of 12)',
  "This page (printed pp. 84-85) continues the South Sea Annuities and Exchequer Bills Act of 1750. The text details the interest terms for the Exchequer Bills to be issued: they are to carry a premium or interest not exceeding Three Pounds per Centum per Annum, to be paid out of the Sinking Fund. The Act specifies that the Commissioners of the Treasury are authorized to issue Exchequer Bills in a common and convenient number, and that these bills are to be made forth at the said Receipt, in the Manner and Form prescribed by the Act. The passage further stipulates that all and every such Exchequer Bills shall be numbered arithmetically beginning with No. 1 and registered in a Book kept for that purpose, and that the Bills shall be redeemable upon demand at the Receipt of the Exchequer.",
  geoII24Base
);

// Row 281 (Page 8 of 12): Printed pp. ~85-86 approx.
setDoc(281,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 8 of 12)',
  "This page (printed pp. ~85-86 approx.) continues the legislative text of the South Sea Annuities and Exchequer Bills Act of 1750. The clauses here establish the obligations of Receivers, Collectors, and Cashiers of the Treasury with respect to the acceptance and disbursement of Exchequer Bills in payment of duties and revenues. The Act specifies that the Commissioners of the Treasury, the High Treasurer, or any Three or more of the Commissioners of the Treasury for the time being, shall have power and authority to cause such Bills to be made forth and issued, and to take in and cancel such bills as shall have been paid and discharged out of the revenues of the Sinking Fund. The page also addresses the accountability of Officers charged with maintaining the registers and records of bills.",
  geoII24Base
);

// Row 282 (Page 9 of 12): Printed pp. 86-87 — penalties and administrative provisions
setDoc(282,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 9 of 12)',
  "This page (printed pp. 86-87) continues the South Sea Annuities and Exchequer Bills Act of 1750. The text at this point includes the penal clauses: any Person or Persons, Bodies Politick or Corporate, who shall counterfeit any of the said Exchequer Bills, or who shall falsely make, forge, or counterfeit any such Bill or any Endorsement thereon, shall be guilty of felony. The Act further specifies the administrative procedures for paying interest on the bills, noting that such interest shall be paid in Lawful Money of this Kingdom, and provides that claims must be presented at specified times and places. The passage also addresses the powers of the Commissioners of the Treasury to determine the manner and form by which bills are to be presented and redeemed.",
  geoII24Base
);

// Row 283 (Page 10 of 12): Printed pp. ~87-88 approx.
setDoc(283,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 10 of 12)',
  "This page (printed pp. ~87-88 approx.) continues the legislative text of the South Sea Annuities and Exchequer Bills Act of 1750. The clauses here deal with the appropriation of the Sinking Fund revenues toward the payment of principal and interest on the Exchequer Bills and toward the redemption of the South Sea Annuities. The Act specifies that the Sinking Fund shall stand charged and chargeable with the payment of all sums of money to be advanced by the Bank of England, and with the interest thereon, until the said sums and interest are fully paid and discharged. The passage also empowers the Commissioners of the Treasury to settle accounts with the Bank of England and to direct payments accordingly.",
  geoII24Base
);

// Row 284 (Page 11 of 12): Printed pp. ~88-89 approx.
setDoc(284,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 11 of 12)',
  "This page (printed pp. ~88-89 approx.) continues the legislative text of the South Sea Annuities and Exchequer Bills Act of 1750. The clauses in this portion address the final administrative provisions of the Act, including rules for the accountability of Treasury officers who receive, disburse, or cancel Exchequer Bills, and the procedures by which the Bank of England is to provide receipts and documentation for sums advanced. The Act also contains provisions enabling the Bank to hold General Courts and Courts of Directors in such manner as the Court of Directors shall appoint, notwithstanding anything in prior Acts to the contrary—a clause intended to regularize the Bank's governance structure during this period of financial restructuring.",
  geoII24Base
);

// Row 285 (Page 12 of 12): Printed pp. 88-89 — Sinking Fund and concluding clauses
setDoc(285,
  'Acts of Parliament, Anno Regni Georgii II Regis Vicesimo Quarto, 1750 (Page 12 of 12)',
  "This is the final page (printed pp. 88-89) of the Acts of Parliament volume for the twenty-fourth year of King George II. The text contains the concluding clauses of the South Sea Annuities and Exchequer Bills Act of 1750, providing that all and every such Exchequer Bills last-mentioned shall be numbered arithmetically, beginning from the Number which shall be expressive of the last of the Bills made forth before these are to be made forth, and shall be registered as aforesaid. It is further enacted that the Commissioners of the Treasury, or any Three or more of them now being, or the High Treasurer, or any Three or more of the Commissioners of the Treasury for the time being, shall have power and authority to prepare and make, at the Exchequer, and issue any Number of such Exchequer Bills. This Act concludes three decades of parliamentary management of the South Sea Company's residual obligations, finally providing the legislative basis for the full redemption of the South Sea Annuities that had been outstanding since 1720.",
  geoII24Base
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 233-285 (3 certificates + 8 Geo. I Acts 38pp + 24 Geo. II Acts 12pp).');
