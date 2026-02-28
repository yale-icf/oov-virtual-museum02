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

// --- PAINE PAMPHLET CONTINUATION (same volume, rows 203-213) ---

const paineBase = {
  type: 'Pamphlet',
  subjectCountry: 'Great Britain',
  issuingCountry: 'United States',
  creator: 'Thomas Paine',
  issueDate: '1796-01-01',
  currency: '',
  language: 'English',
  numberPages: 23,
  period: '18th Century or before',
  notes: 'Thomas Paine, "The Decline and Fall of the English System of Finance" bound with "Speech of Thomas Paine." Philadelphia: John Ormrod for Benjamin Franklin Bache, 1796. Beinecke Library, American Tracts 1796, P161.',
};

// 0203 - pages 22-23
setDoc(203,
  'Paine, The Decline and Fall of the English System of Finance – Pages 22–23 (Page 13 of 23)',
  'This page spread (pp. 22–23) continues Paine\'s argument about the mathematical inevitability of the English funding system\'s collapse. He discusses the role of government paper lent to merchants (citing a sum of fifty millions), and remarks on the practice of issuing accommodation paper that competes with genuine commercial credit. Paine compares the English funding system unfavorably to Holland, noting that there is no country in Europe that could be made the dupe of such a delusion on so large a scale, and identifies this episode as a monument of wonder about the extent to which it has been carried. He addresses those who formerly believed the funding system would amount to between 150 and 200 million shillings, and shows their predictions were wildly underestimated.',
  paineBase
);

// 0204 - pages 24-25
setDoc(204,
  'Paine, The Decline and Fall of the English System of Finance – Pages 24–25 (Page 14 of 23)',
  'This page spread (pp. 24–25) discusses the three distinct ways in which Bank of England notes enter circulation: as a bank of discount (for merchants\' bills of exchange), as a bank of deposit, and as a banker for the government (advancing money to the Exchequer and discounting Exchequer bills). Paine argues that the bank\'s role as lender to the government is the most dangerous function: it places the bank in a position of advancing money it does not have, against the security of future tax revenues, creating a circular mechanism of paper money inflation. He explains how each of these roles independently generates excess paper in circulation, and why all three together make the system structurally unstable.',
  paineBase
);

// 0205 - pages 26-27
setDoc(205,
  'Paine, The Decline and Fall of the English System of Finance – Pages 26–27 (Page 15 of 23)',
  'This page spread (pp. 26–27) addresses the Bank of England\'s role as the government\'s financial agent in greater detail, focusing on how the bank advances Exchequer bills and effectively creates money for the government\'s use. Paine argues that this function—advancing notes to the government against future tax revenues—is the principal source of the excess paper now circulating in England. He notes that the bank is neither merely an ordinary commercial bank nor simply a bank of deposit, but a great engine of the national finance system that continually creates new paper money and pays the government\'s obligations before tax revenues arrive. He identifies this as the mechanism that will eventually force the bank to suspend cash payments when the public demands specie.',
  paineBase
);

// 0206 - pages 28-29
setDoc(206,
  'Paine, The Decline and Fall of the English System of Finance – Pages 28–29 (Page 16 of 23)',
  'This page spread (pp. 28–29) discusses the historical precedent set by France and the limit imposed by the nature of things on the quantity of paper money a nation can sustain. Paine draws on Necker\'s data on French finances before the Revolution, noting that France had approximately 90 million livres sterling in paper money while having about the same sum in gold and silver—and that the limit of paper issuance before depreciation equals the total stock of gold and silver in the country. He applies this principle to England, arguing that the Bank of England has already exceeded this natural limit and therefore the depreciation is already underway, even if not yet clearly visible in price levels.',
  paineBase
);

// 0207 - pages 30-31
setDoc(207,
  'Paine, The Decline and Fall of the English System of Finance – Pages 30–31 (Page 17 of 23)',
  'This page spread (pp. 30–31) addresses the second generation of debt and creditors created by the funding system. Paine notes that a new class of creditors—much larger and more numerous than the first—has been produced by the system, who hold bank notes and government paper as their form of wealth. When the system collapses, he argues, this class will be ruined along with the original investors. He discusses the conduct of Pitt and Grenville in managing the funding system, noting that by lending money to political allies (the "borough-holders") and then having them lend it back to the government as bank notes, the very men who insist the funded debt should be repaid are the same men advancing more paper to postpone reckoning.',
  paineBase
);

// 0208 - pages 32-33 (end of "Decline and Fall")
setDoc(208,
  'Paine, The Decline and Fall of the English System of Finance – Pages 32–33 (Page 18 of 23)',
  'This page spread (pp. 32–33) contains the concluding pages of Paine\'s "Decline and Fall of the English System of Finance." He discusses the five-pound bank notes circulated among small traders, the consequences of the bank\'s issuing notes of so small a denomination, and the resulting inability of the bank to hold sufficient reserves. He then discusses the French experience with the funding system under Necker and Calonne, and concludes that the English system is more structurally similar to the French system than English commentators acknowledge. The text ends with Paine\'s signature: "THOMAS PAINE, Paris, 19th Germinal, 4th year of the Republic, April 8, 1796"—followed by a brief separate heading "ON THE GULF OF BANKRUPTCY," quoting the parliamentary debate epigraph used on the title page.',
  paineBase
);

// 0209 - "Speech of Thomas Paine" title page + opening text
setDoc(209,
  'Paine, Speech in the French Convention, July 7, 1795 – Title Page (Page 19 of 23)',
  'This page opens the second text bound in this volume: "SPEECH OF THOMAS PAINE, As delivered in the Convention, July 7, 1795. Wherein he alludes to the preceding Work." An editorial note explains that on the motion of Louis-Legendre, the Convention invited Paine to deliver his opinions on the Constitution and the Declaration of Rights, and that despite opposition during the proceedings Thomas Paine attended and read his speech, of which this is a literal translation. The opening of the speech addresses the difficulties Paine experienced during his imprisonment ("During a residence of more than six years in France, I have been a close confessor to the Convention...") and his commitment to republican principles over self-interest. This speech, delivered eighteen months before the pamphlet was published, directly references the fiscal arguments Paine had been developing about the English system.',
  paineBase
);

// 0210 - Speech pages continuing
setDoc(210,
  'Paine, Speech in the French Convention, July 7, 1795 – Pages 14–15 (Page 20 of 23)',
  'This page spread continues Paine\'s speech to the French Convention of July 7, 1795. Paine describes his conduct during the Reign of Terror and his imprisonment, noting that he never abandoned the principles of liberty though he suffered for them under Robespierre. He argues that a constitution must be grounded in the Declaration of Rights, and examines specific articles of the proposed French Constitution against the standard of the Declaration. He notes that the article on liberty of the press and freedom of conscience in the French Constitution has a "tendency to infringe them," and insists the constitution should protect liberty not merely in theory but in practice. The speech is translated from the French version Paine delivered, and represents his continuing engagement with republican constitutional theory.',
  paineBase
);

// 0211 - Speech pages 36-37
setDoc(211,
  'Paine, Speech in the French Convention, July 7, 1795 – Pages 36–37 (Page 21 of 23)',
  'This page spread (pp. 36–37) continues Paine\'s analysis of specific articles of the proposed French Constitution against the standard of the Declaration of Rights. He focuses on the article governing property and taxation, arguing that the Constitution is incompatible with the third article of the Declaration of Rights (that the principle of all sovereignty resides in the nation). He introduces a distinction between direct and indirect taxation, arguing that land taxes fall on the landed proprietor while indirect taxes on articles of consumption fall on the consumer, yet neither satisfactorily corresponds to the principle of equal contribution. He compares the tax obligations of a merchant versus a landowner versus a simple mechanic, arguing none contributes proportionally to the exigencies of the state under the proposed framework.',
  paineBase
);

// 0212 - Speech pages 38-39
setDoc(212,
  'Paine, Speech in the French Convention, July 7, 1795 – Pages 38–39 (Page 22 of 23)',
  'This page spread (pp. 38–39) addresses the social and political implications of the proposed French Constitution. Paine argues that government must act for the public good—not merely the good of property-holders—and that a constitution which leaves the poor without subsistence has no legitimacy. He discusses the obligation of the French Citizen as a soldier, noting that a proprietor who wishes to live in peace has more to lose than a soldier-citizen and should therefore contribute more to the common defense. Paine concludes that while he has accepted the proposition to submit his opinions, he must say frankly that many of the dispositions of the Constitution conflict with the principles of the Declaration of Rights, and that a committee should compare the two texts and make the Constitution perfectly consistent with the Declaration.',
  paineBase
);

// 0213 - page 40 + back cover (Beinecke)
setDoc(213,
  'Paine, Speech in the French Convention – Final Page and Back Cover (Page 23 of 23)',
  'This image shows the final page (p. 40) and back cover of the pamphlet volume containing both "The Decline and Fall of the English System of Finance" and "Speech of Thomas Paine." The last paragraph of the speech concludes: "But to discard all considerations of a personal and subordinate nature, it is essential to the well-being of the republic, that the practical or organic part of the constitution should correspond with its principles; and as this does not appear to be the case in the plan that has been presented to you, it is absolutely necessary that it should be submitted to the revision of a committee... in order to ascertain the difference between the two, and to make such alterations as shall render them perfectly consistent and compatible with each other." The back cover bears a handwritten library notation: "Beinecke Library, American Tracts 1796, P161," confirming this copy\'s provenance from Yale University\'s Beinecke Rare Book & Manuscript Library.',
  paineBase
);

// --- JACOB HENRIQUES LOTTERY SCHEME, LONDON 1753: row 214 ---

setDoc(214,
  'Jacob Henriques: Proposed Lottery Scheme, London, April 1753',
  'This printed broadside by Jacob Henriques of London, dated April 1753, proposes a new lottery scheme of 100,000 tickets at four pounds each, to be drawn by quadruple numbers so that 25,000 numbers will determine the whole lottery. The prize scheme offers: 4 prizes of £10,000, 8 of £5,000, 40 of £1,000, 130 of £500, 340 of £100, 800 of £50, 2,000 of £20, and 12,000 of £10, totaling 15,272 prizes of £394,000 with 84,728 blanks (a ratio of only five and a half blanks per prize). An additional £6,000 is distributed to the first and last drawn tickets. Twelve per cent is to be deducted from prizes for the intended purpose and paid either in money or in annuities at 3 per cent. Henriques argues that "My Proposal for lessening the national Debt is real Good, and my Intention for it is still better; and if the Parliament and Government are pleased for the Sake of their Community to encourage me with their prudent public Approbation, I will find out other great Matters for the Glory, Honour, and Welfare of his excellent Majesty." Beinecke Library call number annotations "65.802" and "cw" appear at top right.',
  {
    type: 'Prospectus',
    subjectCountry: 'Great Britain',
    issuingCountry: 'Great Britain',
    creator: 'Jacob Henriques',
    issueDate: '1753-04-01',
    currency: 'GBP',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Jacob Henriques, proposed lottery scheme broadside, London, April 1753. 100,000 tickets at £4 each.',
  }
);

// --- US CONTINENTAL LOTTERY TICKET, CLASS THE FIRST, NO. 16m769: row 215 ---

setDoc(215,
  'United States Continental Lottery Ticket, Class the First, No. 16m769, 1776',
  'This printed lottery ticket reads: "United States Lottery No. 16m769 – CLASS the FIRST. THIS TICKET entitles the Bearer to receive such PRIZE as may be drawn against its Number, according to a Resolution of CONGRESS, passed at Philadelphia, November 18, 1776." Signed in manuscript by "D. Jackson" (a lottery manager) and marked "I." at the lower left to indicate the first class. The Continental Congress authorized this lottery by resolution on November 18, 1776 to help finance the Revolutionary War, organizing it in successive classes with tickets priced at $20 each. The printed ornamental border, the engraved script heading, and the official congressional authorization text are all identical to the other surviving Continental Lottery ticket in this collection (goetzmann0186), confirming both were printed from the same plate and issued under the same authority. This ticket is among the few known surviving examples of the first class of the United States Continental Lottery.',
  {
    type: 'Ticket',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'D. Jackson (lottery manager)',
    issueDate: '1776-11-18',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'United States (Continental) Lottery ticket, Class the First, No. 16m769. Authorized by Resolution of Congress, Philadelphia, November 18, 1776. Signed by D. Jackson.',
  }
);

// --- SOUTH SEA ANNUITIES DIVIDEND WARRANT, 1730: rows 216-217 ---

setDoc(216,
  'South Sea Annuities Dividend Warrant, No. 624, to Conrade de Gols, 1730 (Page 1 of 2)',
  'This printed and manuscript dividend warrant is headed "14th" and numbered "No. 624," addressed to "Mr. Conrade De Gols, Sir." It directs: "You may pay on the 13th Day of May, 1730 [to] Mr. John Bardin the Sum of Sixteen Pounds Eight Shillings for Half a Year\'s Annuity, at 4 per Cent. per Annum, on the Sum of £820 Interest or Share in the Joint-Stock of South-Sea-Annuities, Erected by Act of Parliament in the Ninth Year of the Reign of His late Majesty King George, made for dividing the whole Capital Stock of the South-Sea-Company into Two equal Parts or Moieties, and for converting one of the said Moieties into certain Annuities... which became due on the 25th Day of March, 1730, and take a Discharge on the Back hereof. South-Sea-House the 30th Day of April, 1730." Signed by Williams and E. Anderson. This warrant is evidence of the post-Bubble reorganization of the South Sea Company, in which the company\'s capital stock was divided into equity shares (South Sea Stock) and annuities (South Sea Annuities) to satisfy creditors.',
  {
    type: 'Bond',
    subjectCountry: 'Great Britain',
    issuingCountry: 'Great Britain',
    creator: 'South Sea House (Williams; E. Anderson)',
    issueDate: '1730-04-30',
    currency: 'GBP',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'South Sea Annuities dividend warrant, No. 624, 14th payment, to Conrade de Gols, directing payment to John Bardin. £16 8s on £820 annuity at 4% p.a. South Sea House, 30 April 1730.',
  }
);

setDoc(217,
  'South Sea Annuities Dividend Warrant – Endorsement to John Bardin Junr., 1730 (Page 2 of 2)',
  'This image shows the reverse of South Sea Annuities dividend warrant No. 624 (goetzmann0216). The back carries manuscript endorsements acknowledging receipt and transferring the payment: "John Bardin Junr." written at the top in large script, with the instruction "Pay Mr John Bardin Just[?]" signed by "Williams." This endorsement format was standard for dividend warrants and bills of exchange in 18th-century England: the warrant was directed to the company cashier (Mr. Conrade de Gols), then endorsed by the payee (John Bardin) to authorize transfer of the dividend payment to another party. The South Sea Annuities were created in 1720–22 as part of the financial reconstruction following the South Sea Bubble, converting the inflated share capital of the South Sea Company into permanent government-backed annuities paying 4 per cent per annum.',
  {
    type: 'Bond',
    subjectCountry: 'Great Britain',
    issuingCountry: 'Great Britain',
    creator: 'South Sea House',
    issueDate: '1730-04-30',
    currency: 'GBP',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Reverse and endorsement of South Sea Annuities dividend warrant No. 624. Endorsed by John Bardin Junr. Page 2 of 2.',
  }
);

// --- ANTWERP EXCHANGE OBLIGATION, 1725: rows 218-219 ---

setDoc(218,
  'Antwerp Exchange Obligation, Generale Keyserlyke Company, 1725 (Page 1 of 2)',
  'This printed and handwritten Dutch-language obligation, issued at Antwerp at the Exchange ("t\'Antwerpen bg A. Du Casté og de Borse 1725"), bears an engraved coat of arms at the top. The text ("Ck Onderteekene... beloove en obligeer my... een Contante premie") records a financial obligation by J.B. Franciscus, acting as voordeel-Man (profit-man or agent), to deliver and pay for shares in the Generale Keyserlyke Jut-die Compagnie (Imperial General... Company) at Antwerp. The contract specifies conditions for reimbursement of all legitimate costs, payment of dividends, and liability if the shares are not delivered on the appointed transport-day. The document was signed on a seventeenth of April in the year seventeen hundred and twenty-five, in Antwerp, before witnesses Jan van Lancker and Joan Parcolo Wittebog as guarantor. The Antwerp Exchange was one of the most important trading floors in the Spanish/Austrian Netherlands.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'J.B. Franciscus (voordeel-Man); witnessed by Jan van Lancker',
    issueDate: '1725-04-17',
    currency: '',
    language: 'Dutch',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Dutch-language exchange obligation, Antwerp Exchange (Borse), 1725. Issued by A. Du Casté at Antwerp. Guarantor: Joan Parcolo Wittebog. Page 1 of 2.',
  }
);

setDoc(219,
  'Antwerp Exchange Obligation – Reverse with Guarantee Inscription, 1725 (Page 2 of 2)',
  'This image shows the reverse of the 1725 Antwerp exchange obligation (goetzmann0218). The back of the document shows the obligation\'s text showing through the paper, and bears manuscript endorsements at the top: "adree dit contract aen d\'ordre mons. Joan Kramp zonder guarcano" (address this contract to the order of Mr. Joan Kramp without guarantee) and the signature of "Joan Parcolo Wittebog" as guarantor. The reverse endorsements were written to transfer the obligation to Joan Kramp, reflecting the negotiable character of such exchange contracts on the Antwerp Bourse. The coat of arms watermark from the front also shows through on the reverse. This document is evidence of the sophisticated secondary market in financial obligations that operated in the Austrian Netherlands during the early 18th century.',
  {
    type: 'Bond',
    subjectCountry: 'Netherlands',
    issuingCountry: 'Netherlands',
    creator: 'J.B. Franciscus (voordeel-Man)',
    issueDate: '1725-04-17',
    currency: '',
    language: 'Dutch',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Reverse of Antwerp Exchange obligation, 1725. Endorsement by Joan Parcolo Wittebog transferring to Joan Kramp. Page 2 of 2.',
  }
);

// --- ALEXANDER HAMILTON CIRCULAR LETTER, TREASURY DEPARTMENT, 1795: row 220 ---

setDoc(220,
  'Alexander Hamilton Circular Letter on State Debt Certificates, Treasury Department, 1795',
  'This manuscript letter, headed "Treasury Department Oct 13, 1795 (Circular)," is signed "A. Hamilton" and addressed to loan commissioners across the United States. Hamilton writes: "I find from a letter from one of the Commissioners of Loans, that it is conceived, certificates of the State Debts cannot be subscribed to the new Loans unless that they express that they are issued for services or supplies towards the prosecution of the late war." He advises that this overly narrow construction would improperly exclude legitimate state debt certificates, and that any certificate of a state that was issued for compensations and expenditures toward prosecution of the war and the defense of the United States should be accepted as eligible for subscription to the new federal loan, even if the certificate does not explicitly state this purpose. This circular represents Hamilton\'s active management of the first U.S. federal debt consolidation under the Funding Act of 1790, clarifying eligibility rules for converting state war debts into federal securities.',
  {
    type: 'Letter',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Alexander Hamilton (Secretary of the Treasury)',
    issueDate: '1795-10-13',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Alexander Hamilton circular letter, Treasury Department, October 13, 1795. On eligibility of state debt certificates for subscription to new federal loans under the Funding Act of 1790.',
  }
);

// --- MASSACHUSETTS BILLS OF EXCHANGE SCHEDULE, THOMAS HOPKINSON, 1778: rows 221-222 ---

setDoc(221,
  'Massachusetts Bills of Exchange Schedule, Thomas Hopkinson, Philadelphia, 1778 (Page 1 of 2)',
  'This handwritten schedule, titled "Massachusetts," records a distribution of bills of exchange in multiple denominations for the State of Massachusetts, signed by Thomas Hopkinson, LL [Loan-office Loan?], Philadelphia, September 22, 1778. The columns list denomination (from $12 to $1,200), sets, serial number ranges, and amounts in dollars. Totals: 1,803 sets of bills amounting to $93,000. Specific entries: $12 denomination—305 sets (nos. 36-340)=$3,660; $18—305 sets=$5,490; $24—305 sets=$7,320; $30—305 sets=$9,150; $36—305 sets=$10,980; $60—100 sets=$6,000; $120—100 sets=$12,000; $300—48 sets=$14,400; $600—20 sets=$12,000; $1,200—10 sets=$12,000. This document records the allocation of Continental loan-office certificates or bills of exchange drawn on behalf of Massachusetts during the Revolutionary War, consistent with the Continental Congress\'s financing operations in 1778.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Thomas Hopkinson (loan office official)',
    issueDate: '1778-09-22',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Massachusetts bills of exchange schedule, 1,803 sets totaling $93,000. Signed Thomas Hopkinson, LL, Philadelphia, September 22, 1778. Page 1 of 2.',
  }
);

setDoc(222,
  'Massachusetts Bills of Exchange Schedule – Reverse and Receipt Endorsement, 1778 (Page 2 of 2)',
  'This image shows the reverse of the Massachusetts bills of exchange schedule (goetzmann0221), bearing the docketing endorsement at the top: "List of Bills Exchange, Rec\'d Oct. 12th 1778—" along with a partial manuscript note about "for Massachusetts Money" and a signature. A yellow wax seal remnant is visible at the lower right. The docketing inscription records when this list of bills was received by the relevant office, October 12, 1778, approximately three weeks after the original schedule was prepared by Thomas Hopkinson on September 22, 1778. This reverse endorsement confirms the administrative processing of wartime financial instruments through the Continental Congress\'s loan office network.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Thomas Hopkinson (loan office official)',
    issueDate: '1778-09-22',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Reverse and receipt endorsement of Massachusetts bills of exchange schedule. Received October 12, 1778. Page 2 of 2.',
  }
);

// --- ROWLEY, MASSACHUSETTS FINE WARRANT, 1782: rows 223-225 ---

setDoc(223,
  'Rowley, Massachusetts: Selectmen\'s Fine Warrant to Constable Samuel Searl, 1782 (Page 1 of 3)',
  'This handwritten document from the Selectmen of Rowley, Massachusetts, dated August 15, 1782, is addressed "To Constable Samuel Searl" and orders him to collect fines from the persons named on the accompanying list. The text reads: "the Within List Contains Each Person there in Named is afresed as their Proportion of the fine Levid on the Sixth Class in the Town of Rowley for not procuring a man to Serve in the Continental Army pr. year agreeable to a Resolve of the Generel Court. Each Persons Tax is in the line with their name and the Same is Committed to you to Colect." Signed by the Selectmen of Rowley: Thomas Mighill, Daniel Spofford, Joseph Poor, [?] Pickard, and Isaac Smith. This document illustrates the Massachusetts system of class-based conscription, by which towns were divided into classes of taxpayers, each class responsible for providing a soldier for a year or paying a fine in lieu of service.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Selectmen of Rowley, MA (Thomas Mighill, Daniel Spofford, Joseph Poor, et al.)',
    issueDate: '1782-08-15',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'Rowley, MA Selectmen\'s fine warrant to Constable Samuel Searl, August 15, 1782. Fines for Sixth Class members failing to provide a Continental Army recruit. Page 1 of 3.',
  }
);

setDoc(224,
  'Rowley, Massachusetts: Fine Assessment List, Sixth Class, 1782 (Page 2 of 3)',
  'This handwritten two-column document lists the names and fine assessments (in pounds, shillings, pence) for members of the Sixth Class in Rowley, Massachusetts who failed to procure a man for the Continental Army in 1782. Left column names include Mr. Jacob Pearson & Son (£2.10), Mr. Mark Thorla, Mr. John Thorla, Mr. Henry Poor, Joseph Poor (£4.4.9), George Poor, Benjamin Poor, Mr. Thomas Lull & Son (£3.14.9), Mr. Oliver Jenny & Son (£5.2.1), Mr. Moses Wheeler jr., Moses Wheeler 3rd, Mr. Israel Adams & Son (£3.18.9), Mr. Hayes Pearson (£2.13.4), Mr. Timothy Jackman, Mr. Benjamin Jackman, Capt. Timothy Jackman, Mr. Benjamin Jackman jr., Mr. Moses Dale (£4.19.9), Deacon Joseph Searl & Son (£5.13.9), Mr. Nathaniel Jenny & Son (£8.5.5), Mr. Stephen Dale & Son (£14.4.9). Right column includes Capt. Daniel Chute & Sons, Mr. Nathaniel Hoges, Mr. Moses Jenny, Mr. Samuel [?], Mr. Thomas Plummer, Ensign Lemuel Roger [Newbury], Mr. Jeremiah Pearson, Mr. Daniel Hale, Mr. Samuel Thorla & Adams, Mr. Thomas Smith, Mr. Nathaniel Adams, and others. This list documents the local fiscal burden imposed by the Continental Army\'s manpower demands on a small Essex County farming town.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Selectmen of Rowley, MA',
    issueDate: '1782-08-15',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'Rowley, MA fine assessment list for Sixth Class, 1782. Names and amounts in £.s.d. Page 2 of 3.',
  }
);

setDoc(225,
  'Rowley, Massachusetts: Fine Warrant Reverse, 1782 (Page 3 of 3)',
  'This image shows the reverse of the Rowley, Massachusetts Sixth Class fine assessment document (goetzmann0224). The back of the sheet is almost entirely blank, apart from a faint docketing inscription at the top reading "of the Sixth Class," and the paper has suffered substantial water damage visible as a large brown stain in the upper center. The blank reverse, with only a minimal docketing note, is typical of 18th-century American administrative documents—the recto carried all substantive information while the verso served only as an identifying wrapper or filing surface. This document, together with the selectmen\'s warrant (goetzmann0223) and the assessment list (goetzmann0224), forms a complete three-page record of local Revolutionary War-era conscription finance in Essex County, Massachusetts.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Selectmen of Rowley, MA',
    issueDate: '1782-08-15',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'Reverse of Rowley, MA Sixth Class fine assessment document, 1782. Water damaged. Page 3 of 3.',
  }
);

// --- STATE OF MASSACHUSETTS BAY BOND, SETH DAVENPORT, 1777: rows 226-227 ---

setDoc(226,
  'State of Massachusetts Bay War Bond, £20, Seth Davenport, 1777 (Page 1 of 2)',
  'This printed and handwritten bond certificate is issued by the State of Massachusetts Bay, dated February 26, 1777. It reads: "Borrowed and received of Seth Davenport, the Sum of Twenty Pounds, Lawful Money, for the Use and Service of the State of Massachusetts-Bay; and in behalf of said State I do hereby promise and oblige myself and Successors in the Office of Treasurer or Receiver-General, to repay the Possessor by the First Day of June 1780, the aforesaid Sum of [Twenty Pounds] Lawful Money, in Spanish mill\'d Dollars at Six Shillings each, or in the several Species of coined Silver and Gold enumerated in an Act made and passed in the Twenty third Year of his late Majesty King George the Second... with Interest to be paid annually at Six per Cent." Signed by D. Jeffries and the Committee (J. Summer). This is a direct emission of state war debt during the American Revolutionary War, promising repayment by June 1780 at 6% annual interest, before the federal consolidation of state debts under Hamilton\'s Funding Act of 1790.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Massachusetts Bay (Treasurer D. Jeffries; Committee)',
    issueDate: '1777-02-26',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'State of Massachusetts Bay war bond, £20 at 6% p.a., issued to Seth Davenport, February 26, 1777. Due June 1, 1780. Page 1 of 2.',
  }
);

setDoc(227,
  'State of Massachusetts Bay Bond – Reverse with Consolidation Endorsement, 1777 (Page 2 of 2)',
  'This image shows the reverse of the State of Massachusetts Bay war bond issued to Seth Davenport on February 26, 1777 (goetzmann0226). The back bears manuscript endorsements recording the bond\'s subsequent history: "Seth Davenport Jr., B.323/4, Consolid. 173.1.8, May 1783, N° 7959." These notations record that the bond was held by Seth Davenport Jr. (son of the original bearer), was listed in a consolidation ledger as entry B.323/4, with a consolidated value of £173.1.8, in May 1783—indicating that multiple Massachusetts bonds were combined into a single consolidated state certificate at that time, which was then assigned number 7959. This endorsement is direct evidence of Massachusetts\'s post-war debt consolidation effort, which preceded and informed Alexander Hamilton\'s federal funding scheme of 1790.',
  {
    type: 'Bond',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'State of Massachusetts Bay',
    issueDate: '1777-02-26',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Reverse of Massachusetts Bay war bond. Consolidation endorsement: B.323/4, £173.1.8, May 1783, N° 7959. Page 2 of 2.',
  }
);

// --- DUTCH BOND FOR FRENCH BOURBON PRINCES IN EXILE, AMSTERDAM 1793: row 228 ---

setDoc(228,
  'Dutch Bond for the French Bourbon Princes in Exile, Amsterdam, 1793',
  'This printed Dutch-language bond is headed with a reference to the Peace of Westphalia and bears the names "Louis Stanislas Xavier, Charles Philippe" (i.e., the Comte de Provence and the Comte d\'Artois—the exiled Bourbon princes, brothers of Louis XVI) as issuers, with the title "Le Maréchal, Duc de Brolie" (the Maréchal de Broglie) as signing authority. Numbered N° 373, it records that "J. Bourcourd en Wedowe F. Croese & Comp." of Amsterdam acknowledge receipt of "EEN DUIZEND GULDENS Hollands Courant" (One Thousand Dutch Guilders) in exchange for a bond with attached coupons, dated Amsterdam, September 28, 1793. Signed by the Garde du Trésor Royal and countersigned "Par ordre de leurs Altesses Royales" (By order of their Royal Highnesses). This bond was issued during the period when the French princes were raising funds in the Netherlands to finance their counter-revolutionary campaigns against the French Republic, using Amsterdam bankers as their financial agents.',
  {
    type: 'Bond',
    subjectCountry: 'France',
    issuingCountry: 'Netherlands',
    creator: 'Louis Stanislas Xavier (Comte de Provence); Charles Philippe (Comte d\'Artois); J. Bourcourd & Wedowe F. Croese & Comp. (Amsterdam)',
    issueDate: '1793-09-28',
    currency: 'NLG',
    language: 'Dutch',
    numberPages: 1,
    period: '18th Century or before',
    notes: 'Dutch bond (1,000 guilders) issued by French Bourbon princes in exile (Louis Stanislas Xavier and Charles Philippe) through Amsterdam bankers J. Bourcourd & Wedowe F. Croese & Comp., September 28, 1793. Signed on behalf of the Garde du Trésor Royal.',
  }
);

// --- RUSSIAN IMPERIAL TREASURY NOTE, 50 RUBLES, 1915: row 229 ---

setDoc(229,
  'Russian Imperial Treasury Note, 50 Rubles, 4%, 1915',
  'This ornate printed Russian Imperial Treasury Note ("Билета Государственного Казначейства в Пятьдесят Рублей" – Treasury Note for Fifty Rubles), No. 352814, dated 1915, bears the double-headed imperial eagle and is printed in green and black. The note bears a 4% annual interest rate and is declared valid until August 1, 1929 ("Билет действителен по 1 Августа 1929 г."). Four attached coupons at the left side are numbered 5 through 8 and each represent one ruble of interest per coupon. The note is signed by the Director of the Imperial State Treasury and a Bookkeeper. Russian Short-term Treasury Notes (казначейские билеты) were issued throughout the late Imperial period as interest-bearing short-term government obligations; this 1915 example was issued during World War I as Russia\'s financial situation became increasingly strained. The note was never redeemed, as the Imperial government collapsed in the revolutions of 1917.',
  {
    type: 'Bond',
    subjectCountry: 'Russia',
    issuingCountry: 'Russia',
    creator: 'Imperial Russian State Treasury',
    issueDate: '1915-01-01',
    currency: 'RUB',
    language: 'Russian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Russian Imperial Treasury Note, 50 rubles, 4%, No. 352814, 1915. Valid until August 1, 1929. With four interest coupons. Never redeemed due to collapse of Imperial government in 1917.',
  }
);

// --- GREEK GUARANTEED GOLD LOAN BOND, 2½%, 1898: row 230 ---

setDoc(230,
  'Kingdom of Greece Gold-Guaranteed Loan Bond, 2½%, 2,500 Drachmas, 1898',
  'This elaborately printed bond certificate is headed "Βασίλειον της Ελλάδος" (Kingdom of Greece) and titled "Δάνειον εις Χρυσόν Ηγγυημένον 2½ θό 1898 / Ομολογία Ανωνύμου 2,500 Δραχμών" (Gold-Guaranteed Loan 2½% 1898 / Bearer Bond of 2,500 Drachmas). No. 30,731. The bond is also labeled "Emprunt Hellénique Garanti 2½% Or de 1898" (Greek Guaranteed Gold Loan 2½% of 1898) in French along the left border. Printed in blue and red with classical allegorical figures (a male warrior and a draped female figure flanking the central text) and the Greek royal coat of arms. The bond is signed by the Greek Minister of Finance and authenticated by the Governor and Company of the Bank of England. The 1898 Greek Guaranteed Loan was arranged under international financial supervision following Greece\'s default of 1893 and defeat in the Greco-Turkish War of 1897, with France, Great Britain, and other powers guaranteeing repayment as part of an International Financial Control Commission imposed on Greece.',
  {
    type: 'Bond',
    subjectCountry: 'Greece',
    issuingCountry: 'Greece',
    creator: 'Kingdom of Greece (Ministry of Finance); guaranteed by Bank of England',
    issueDate: '1898-01-01',
    currency: 'GRD',
    language: 'Greek',
    numberPages: 1,
    period: '19th Century',
    notes: 'Kingdom of Greece Gold-Guaranteed Loan bond, 2½% 1898, 2,500 drachmas, No. 30,731. International guarantee; authenticated by Bank of England. Issued under International Financial Commission supervision following Greek default of 1893.',
  }
);

// --- ROMANIAN ARMY EQUIPMENT BOND, 500 LEI, 1940: row 231 ---

setDoc(231,
  'Romania: Army Equipment Bond (Bon pentru Înzestrarea Armatei), 500 Lei, 1940',
  'This ornate multi-colored bond certificate is issued by "România, Casa Autonomă a Monopolului Apărării Naţionale" (Romania, Autonomous Administration of National Defense Monopolies) and titled "Bon pentru Înzestrarea Armatei" (Bond for Army Equipment), Series A, No. 361,551, for 500 lei. The bond is dated 1940 and authorized by a law of April 6, 1940. The text indicates payment on March 1, 1945 by the Autonomous Administration of National Defense Monopolies. The certificate is signed by the Governor and Director General of the Autonomous Administration. The design is richly printed in red, green, and gold with Romanian national ornamental patterns, the Romanian coat of arms, and three small vignette images of military scenes at the bottom (artillery, infantry, cavalry). The bond was issued to finance Romanian military rearmament as World War II began, at a time when Romania was attempting to strengthen its defenses before the Nazi-Soviet Pact pressure forced it to cede Bessarabia and northern Transylvania in 1940.',
  {
    type: 'Bond',
    subjectCountry: 'Romania',
    issuingCountry: 'Romania',
    creator: 'Casa Autonomă a Monopolului Apărării Naţionale (Romania)',
    issueDate: '1940-04-06',
    currency: 'RON',
    language: 'Romanian',
    numberPages: 1,
    period: '20th Century',
    notes: 'Romanian Army Equipment Bond, Series A, No. 361,551, 500 lei, 1940. Issued to finance military rearmament. Due March 1, 1945.',
  }
);

// --- AMERICAN & BRITISH SECURITIES COMPANY STOCK CERTIFICATE, 1919: row 232 ---

setDoc(232,
  'American & British Securities Company Common Stock Certificate, 100 Shares, 1919',
  'This printed stock certificate, No. 18, certifies that "Hattie J. Barnhill" is the owner of one hundred shares of Common Stock of the American & British Securities Company, incorporated under the laws of the State of Delaware. The certificate is stamped "COMMON" in large red letters diagonally across the center. Terms specify that each share of preferred stock and each share of common stock outstanding shall be entitled to one vote at stockholders\' meetings. The certificate is dated September 1919 (a "29" day is partially legible), signed by the Secretary and another officer, with a transferred signature "Edward P[?] Holmes" on the back. A handwritten annotation "98-0186.06b" appears at top right, indicating a museum or archive accession number. The American & British Securities Company was a Delaware investment holding company of the type common in the early 20th century, capitalizing on Anglo-American financial connections in the immediate post-World War I period.',
  {
    type: 'Stock',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'American & British Securities Company',
    issueDate: '1919-09-29',
    currency: 'USD',
    language: 'English',
    numberPages: 1,
    period: '20th Century',
    notes: 'American & British Securities Company common stock certificate, 100 shares, No. 18, issued to Hattie J. Barnhill, September 1919. Incorporated in Delaware.',
  }
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 203–232 (Paine pamphlet continuation, Jacob Henriques lottery, Continental Lottery ticket, South Sea Annuities warrant, Antwerp obligation, Hamilton circular, Massachusetts bills, Rowley fine warrant, Massachusetts Bay bond, Dutch Bourbon bond, Russian Treasury note, Greek bond, Romanian bond, American & British Securities stock).');
