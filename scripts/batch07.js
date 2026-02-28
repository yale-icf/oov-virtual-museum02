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

// --- ENGLISH STATE LOTTERY BROADSIDES (continued): rows 170–178 ---

const lotteryBase = {
  type: 'Advertisement',
  subjectCountry: 'Great Britain',
  issuingCountry: 'Great Britain',
  currency: 'GBP',
  language: 'English',
  numberPages: 1,
  period: '19th Century',
};

// 0170 - State Lottery Now Drawing (Third Day/Fifth Day broadside)
setDoc(170,
  'English Lottery Broadside: State Lottery Third and Fifth Day Prizes',
  'This printed English State Lottery broadside announces drawing prizes in bold typography: "THIRD DAY, First-drawn Ticket, £10,000" and "FIFTH DAY, First-drawn Ticket, £30,000." Below, two "Interesting Questions" frame the lottery as a rational financial opportunity: "Is there any other possible mode by which I may gain, from a stake of a few shillings only, nearly Two Thousand Pounds?" and "Have I no ambition to make my fortune by a mode of adventure, where the loss is trifling, but the gain may be great?" Tickets and shares are offered at all the Licensed Lottery Offices and their agents in the country. Printed by Evans & Ruffy, 29 Budge Row, Wallbrook—the same printers associated with several other lottery trade cards in this collection—suggesting a date in the range ca. 1790s–1820s.',
  { ...lotteryBase, issueDate: '1800-01-01', creator: 'Evans & Ruffy (printers)', notes: 'English State Lottery broadside, ca. 1790s–1820s. Evans & Ruffy, Printers, 29 Budge Row, Wallbrook.' }
);

// 0171 - Grand Lottery Now Drawing (prize scheme with lottery agent figure)
setDoc(171,
  'English Lottery Broadside: Grand Lottery Drawing with Prize Scheme',
  'This printed English State Lottery broadside is headed "Grand Lottery Now Drawing" and features a central wood-engraved figure of a lottery agent in academic robes holding a banner reading "PRIZES," with the prize scheme flanking him: two Prizes of £40,000, two of £20,000, two of £10,000, three of £5,000, and further capitals. Below, the same two "Interesting Questions" appear: one asking if there is any other way to turn a few shillings into nearly two thousand pounds. The next day of drawing is announced as Thursday, January 22. Printed by Evans & Ruffy, 29 Budge Row, Wallbrook. This broadside belongs to a series of lottery advertising pieces produced by Evans & Ruffy distributed through licensed lottery offices across England.',
  { ...lotteryBase, issueDate: '1800-01-22', creator: 'Evans & Ruffy (printers)', notes: 'English State Lottery broadside, ca. 1790s–1820s. Evans & Ruffy, Printers, 29 Budge Row, Wallbrook.' }
);

// 0172 - "Positive!!! By Order of the Lords of His Majesty's Treasury" – Bish
setDoc(172,
  'English Lottery Broadside: "Positive!!! All Lotteries End for Ever" – Bish, 1826',
  'This printed lottery broadside is issued "By Order of the Lords of His Majesty\'s Treasury" and announces in dramatic typography: "On the 18th This Month (OCTOBER) All Lotteries End for Ever." It declares that on that day "SIX £30,000 WILL BE DISTRIBUTED, as the Parting Gifts of Fortune." Tickets and shares are sold by Bish, Stock-Broker, at No. 4 Cornhill and 9 Charing Cross, who advertises his success selling the two most recent winning prizes of £21,000 each (drawn 3rd May). Thomas Bish was one of the principal licensed lottery contractors during the final decade of the English State Lottery, and this broadside documents the last official batch of drawings authorized before the abolition act took effect in October 1826.',
  { ...lotteryBase, issueDate: '1826-10-18', creator: 'Bish (Stock-Broker, lottery contractor)', notes: 'English State Lottery final drawing broadside, October 1826. Bish, Stock-Broker, No. 4 Cornhill and 9 Charing Cross.' }
);

// 0173 - "In One Day, Six £20,000, Hazard & Co." – October 1825
setDoc(173,
  'English Lottery Broadside: Hazard & Co. "In One Day, Six £20,000" – October 1825',
  'This two-color printed lottery broadside (green ornamental border with red central text) announces a drawing by Hazard & Co., the Contractors: "IN ONE DAY, SIX £20,000, ALL TO BE DRAWN 18th October, All Money No Blanks." A text block notes that Parliament has determined not to pass any more Lottery Acts and that this drawing is "Nearly the End of Lotteries." Hazard & Co. cite two recent prize successes they sold (2,179... Prize of £25,000 and 6,032... Prize of £25,000) and list offices at 93 Royal Exchange, 26 Cornhill, and 324 Oxford Street. Printed by Whiting & Branston, Beaufort House. A handwritten annotation "Oct 1825" establishes the date, placing this among the penultimate series of English State Lottery drawings.',
  { ...lotteryBase, issueDate: '1825-10-18', creator: 'Hazard & Co. (lottery contractors)', notes: 'English State Lottery broadside, October 1825. Hazard & Co., 93 Royal Exchange, 26 Cornhill, 324 Oxford Street. Printer: Whiting & Branston, Beaufort House.' }
);

// 0174 - "1st March, Six £20,000, Hazard & Co." – February 1826
setDoc(174,
  'English Lottery Broadside: Hazard & Co. "Six £20,000 – Only Two More Chances" – 1826',
  'This two-color printed lottery broadside (blue and salmon/red) announces "1st MARCH, SIX £20,000 IN ONE DAY" by Hazard & Co. Its central design is a large wheel or medallion inscribed "1st MARCH" around the rim, with the six £20,000 prizes radiating from the hub, and the legend "Remember you have only two more chances when Lotteries close for ever." Tickets and shares are sold by Hazard & Co., 93 Royal Exchange, 26 Cornhill, and 324 Oxford Street, and printed by Whiting & Branston, Beaufort House. A handwritten annotation "Feb 1826" dates this to the penultimate English State Lottery drawing, one of the final two ever held before the institution\'s abolition.',
  { ...lotteryBase, issueDate: '1826-03-01', creator: 'Hazard & Co. (lottery contractors)', notes: 'English State Lottery broadside, March 1826 (penultimate drawing). Hazard & Co., 93 Royal Exchange, 26 Cornhill, 324 Oxford Street. Printer: Whiting & Branston, Beaufort House.' }
);

// 0175 - "An Exact Representation of the Drawing... for the LAST TIME" – Carroll, July 18, 1826
setDoc(175,
  'English Lottery Broadside: "The Very Last Lottery of All" – Carroll, 18 July 1826',
  'This printed broadside is headed "An Exact Representation of the DRAWING of the STATE LOTTERY, as it will take place on TUESDAY, the 18th Day of JULY, 1826, for the LAST TIME in this Kingdom." A large wood-engraved image depicts the Lottery drawing room at Coopers\' Hall with fourteen labeled figures: the two Wheels (Numbers and Prizes), the Proclaimers, Bluecoat Boys who draw the numbers and prizes, Commissioners who watch and verify, the President who knocks with a hammer when a prize is proclaimed, and the Commissioners\' Clerks. Below: "THE VERY LAST LOTTERY OF ALL CONTAINS SIX Prizes of £30,000! All in One Day, 18th JULY." A note warns that ticket demand may exceed supply. Sold by Carroll, Joint Contractor, 19 Cornhill; 7 Charing Cross; 26 Oxford-St. This broadside is a primary document recording the final English State Lottery drawing, ending an institution that had operated continuously since 1694.',
  { ...lotteryBase, issueDate: '1826-07-18', creator: 'Carroll (Joint Lottery Contractor)', notes: 'Broadside for the final English State Lottery drawing, 18 July 1826. Carroll, Joint Contractor, 19 Cornhill, 7 Charing Cross, 26 Oxford-St.' }
);

// 0176 - Soldier/"Lott'ry Volunteer" broadside
setDoc(176,
  'English Lottery Broadside: "Spoils like these are worth dividing" – Lottery Volunteer',
  'This printed lottery broadside features a wood-engraved figure of a uniformed military volunteer holding a lottery ticket aloft with the motto "Spoils like these are worth dividing." The verse plays on martial and lottery imagery: "WITH much to hope, and nothing left to fear, / Who would not be a Lott\'ry Volunteer? / Who but would wish,—for Prizes now abound,— / To catch a brilliant THIRTY THOUSAND POUND? / Two still remain; one TWENTY\'s left behind, / One of TEN THOUSAND, Three of FIVE you\'ll find..." The card exploits the patriotic military culture of the Napoleonic era—as Lieutenant Cheerly (goetzmann0145) and similar cards in this series do—to equate lottery prize-winning with battlefield glory. The next drawing is announced for Thursday, January 22. Evans & Ruffy, Printers, 29 Budge Row, Wallbrook.',
  { ...lotteryBase, issueDate: '1800-01-22', creator: 'Evans & Ruffy (printers)', notes: 'English State Lottery broadside, ca. 1790s–1820s. Evans & Ruffy, Printers, 29 Budge Row, Wallbrook.' }
);

// 0177 - "The richest Pine Apple ever seen" – Gye and Balne
setDoc(177,
  'English Lottery Broadside: "The Richest Pine Apple Ever Seen, Value £200,000"',
  'This printed lottery broadside features a large wood-engraved pineapple growing in a barrel or tub, with prize amounts printed directly on its broad leaves: four prizes of £20,000 and four of £5,000 branching out from the central stem. The headline reads "The richest PINE APPLE EVER SEEN, Value £200,000!" and the footer states "One Ticket may gain £100,000." Tickets and shares for the State Lottery are on sale at all Lottery Offices in Town and Country, with the drawing on "8th of NEXT MONTH." The pineapple—an emblem of exotic luxury, royal hospitality, and prestige in Georgian England—here serves as a visual metaphor for the fabulous wealth the lottery promises. Printed by Gye and Balne, 38 Gracechurch-Street.',
  { ...lotteryBase, issueDate: '1810-01-01', creator: 'Gye and Balne (printers)', notes: 'English State Lottery broadside, ca. 1800–1820s. Gye and Balne, Printers, 38 Gracechurch-Street.' }
);

// 0178 - "In Pursuance of the Act 4 Geo. IV. Cap. 60" – official notice, 1826
setDoc(178,
  'English Lottery Official Notice: "Lotteries are about to Cease" (Act 4 Geo. IV, 1826)',
  'This printed official broadside bears the royal coat of arms and is headed "IN PURSUANCE OF THE ACT 4 Geo. IV. Cap. 60." It declares: "Notice is hereby Given, that Lotteries are about to Cease; and that only Two more will be allowed in this Kingdom, viz. one on the 3d of MAY next, and another immediately after that Day. When these Two are drawn, all Lotteries must then be discontinued BY ORDER Of Government." A lower section notes that the current scheme—"the Last but One that will ever be Drawn in England"—contains six prizes of £21,000, to be decided in one day, 3rd May. Printed by Whiting, Printer, Lombard Street. A handwritten annotation reads "1826." The Act 4 George IV Cap. 60 (1823) provided the legislative authority for abolishing the English State Lottery, which had operated since 1694 as a major source of government revenue.',
  { ...lotteryBase, issueDate: '1826-05-03', creator: 'Whiting (printer), issued by order of Government', notes: 'Official government notice broadside for penultimate English State Lottery drawing, 3 May 1826. Under Act 4 Geo. IV Cap. 60. Whiting, Printer, Lombard Street.' }
);

// --- FRENCH ROYAL TONTINE, 1759–1760: rows 179–180 ---

const tontine1759base = {
  type: 'Bond',
  subjectCountry: 'France',
  issuingCountry: 'France',
  creator: "Joseph Micault d'Harvelay (Garde du Trésor royal)",
  issueDate: '1759-12-01',
  currency: 'FRF',
  language: 'French',
  numberPages: 2,
  period: '18th Century or before',
  notes: "French Royal Tontine certificate, 10th Tontine (created by Royal Edict, December 1759). Subscriber: Jean Marie Joseph François Gueneau de Vauzette. Nominee: Victoire Sophie de Noailles. Registered at the Contrôle général des Finances, Paris, 22 September 1760.",
};

setDoc(179,
  'French Royal Tontine Certificate, 10th Tontine, 1759 (Page 1 of 2)',
  'This handwritten and printed tontine certificate is issued under the "10ème Tontine créée par Édit du mois de Décembre 1759" (10th Royal Tontine created by Edict of December 1759). The subscriber is "Jean Marie Joseph François Gueneau de Vauzette, Bourgeois de Paris," who has paid "Deux cent Livres" (200 livres) for a share, with the annuity contingent on the life of "Victoire Sophie de Noailles," described as sixty years old and widow of a member of the household of Alexandre de Bourbon, Comte de Toulouse. Annuities are to be paid twice yearly at the Hôtel des Marchands et Échevins (Merchants\' and Aldermen\'s Hall) of Paris, as established by the Royal Edict and the regulations of the Ferme générale. Signed by Joseph Micault d\'Harvelay, Conseiller du Roi and Garde du Trésor royal. Royal tontines were a major mechanism of French state borrowing throughout the 18th century, raising capital by pooling life annuities that increased as fellow subscribers died.',
  tontine1759base
);

setDoc(180,
  'French Royal Tontine Certificate, 10th Tontine, 1759 (Page 2 of 2)',
  'This page shows the reverse of a French Royal Tontine certificate from the 10th Tontine (Edict of December 1759), bearing the official registration endorsement: "Enregistrée au Contrôle général des Finances par nous Chevalier, Conseiller du Roi en ses Conseils, Garde des Registres du Contrôle général des Finances, commis à cet effet. A Paris, le vingt 2e jour de Septembre mil sept cent soixante." (Registered at the Contrôle général des Finances... Paris, the 22nd of September 1760.) Signed "Berrotin." The registration stamp confirms the certificate\'s validity and the French royal government\'s financial obligation to pay the tontine annuity. The show-through from the recto reveals the subscriber\'s name (Gueneau de Vauzette) and nominee (Victoire Sophie de Noailles). France employed ten royal tontines between 1689 and 1759 as instruments of sovereign debt; this certificate belongs to the last tontine issued before the financial crises of the Seven Years\' War and the Revolution.',
  tontine1759base
);

// --- AMERICAN CONTINENTAL LOTTERY DOCUMENTS, 1776–1778: rows 181–187 ---

// 0181: Subscriber list, Winchester VA, page 1 of 3
setDoc(181,
  'United States Continental Lottery Subscriber List, Winchester, Virginia, 1777 (Page 1 of 3)',
  'This handwritten manuscript is the first page of a United States Continental Lottery ticket subscriber list compiled in Winchester, Virginia, recording names, ticket quantities, and assigned lot numbers (prefixed "52m"). Subscribers include John Hanly (2 tickets), John Norvell (4), William C. Holliday (3), Albert Roberts (1), John Cox (3), Edward M. Quinn (3), John DuHeld (2), Benjamin Williams, Ann Brownson, Elizabeth McFarland, Col. Isaac Lake (10 tickets), and others—totaling at least 86 tickets carried over to the next page. The tickets were sold on behalf of Josiah Watson Esq. of Alexandria, Virginia, who distributed United States Lottery tickets to local agents in the Shenandoah Valley for sale to the public. This list is part of a small archive documenting how Continental Lottery ticket sales were organized through a decentralized network of agents during the Revolutionary War.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'William Holliday (lottery agent, Winchester, VA)',
    issueDate: '1777-07-24',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'United States (Continental) Lottery subscriber list, Winchester, Virginia, ca. July 1777. Agent: William Holliday; distributor: Josiah Watson of Alexandria; commissioner: James Searle, Philadelphia. Page 1 of 3-page document set.',
  }
);

// 0182: Subscriber list page 2 + Holliday letter
setDoc(182,
  'United States Continental Lottery: Holliday Letter to James Searle, Winchester, 1777 (Page 2 of 3)',
  'This manuscript document is the second page of a United States Continental Lottery subscriber list, continuing the tally of Winchester-area ticket purchasers and totaling 100 tickets sold at 10 dollars each for $1,000 (with an arithmetic subtotal noted). Appended below is a letter dated "Winchester, July 24, 1777," from William Holliday to "Mr. James Searle": "I herewith remit you $1000 by the bearer Mr. Samuel Gilkeson agreeable to the directions of Josiah Watson Esq. of Alexandria, who sent me 100 Tickets of the United States Lottery on mine to Sell, which I have accordingly done as pr. foregoing List. I could have sold more if I had them. I am Sir Your most Ob. Hble Serv. Wm Holliday." James Searle was a prominent Philadelphia merchant, Continental Congress delegate, and one of the lottery commissioners who oversaw the first national lottery in American history, authorized by Congress on November 18, 1776 to help finance the Revolutionary War.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'William Holliday (lottery agent)',
    issueDate: '1777-07-24',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'United States Continental Lottery remittance letter from William Holliday (Winchester, VA) to James Searle (Philadelphia), July 24, 1777. $1,000 remitted via Samuel Gilkeson. Page 2 of 3.',
  }
);

// 0183: Reverse/address panel of the Holliday letter
setDoc(183,
  'United States Continental Lottery: Holliday Letter Address Panel, 1777 (Page 3 of 3)',
  'This image shows the folded exterior address panel of the William Holliday remittance letter dated July 24, 1777. The address reads: "Mr. James Searle / Philadelphia / fav\'d of / Mr. Gilkeson." A red wax seal impression is visible at the center fold. Marginal endorsements read: "July 24th, 1777 / Wm. Holliday / Winchester." This format—the letter folded into its own envelope with the address on the outside and sealed with wax—was standard for 18th-century correspondence before the adoption of separate envelopes. The letter was hand-carried ("favored of") by Samuel Gilkeson, who also transported the $1,000 remittance from the Winchester lottery ticket sales to James Searle in Philadelphia, demonstrating the personal courier networks on which wartime financial transactions depended.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'William Holliday (lottery agent)',
    issueDate: '1777-07-24',
    currency: 'USD',
    language: 'English',
    numberPages: 3,
    period: '18th Century or before',
    notes: 'Address panel and exterior of the Holliday-to-Searle Continental Lottery remittance letter, Winchester, VA, July 24, 1777. Page 3 of 3.',
  }
);

// 0184: Robert Traill subscriber list, Easton, Dec 27, 1778
setDoc(184,
  'United States Continental Lottery Ticket Sales List, Robert Traill, Easton, 1778 (Page 1 of 2)',
  'This handwritten manuscript is headed "United States Lottery, Class the Second, List of Tickets sold by Robert Traill of Easton." It records 25 lottery ticket numbers (prefixed "gom") and concludes: "25 Tickets, 20 Dollars each, 500 Dollars." Signed "Rob. Traill, Easton, Decmr. 27, 1778." The United States Continental Lottery was authorized by resolution of Congress on November 18, 1776 to help finance the Revolutionary War, and was organized in successive classes (First, Second, Third) with tickets priced at $20 each. Robert Traill operated as a local lottery agent in Easton, Pennsylvania, distributing and selling tickets on behalf of the lottery commissioners in Philadelphia. This list is evidence of how the Continental Lottery reached small-town markets throughout Pennsylvania through a network of named local agents.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Robert Traill (lottery agent, Easton, PA)',
    issueDate: '1778-12-27',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'United States Continental Lottery, Class the Second, ticket sales list by Robert Traill, Easton, PA, December 27, 1778. 25 tickets, $500 total. Page 1 of 2.',
  }
);

// 0185: Companion lottery ticket sales list (worn)
setDoc(185,
  'United States Continental Lottery Ticket Sales List (Companion Document), ca. 1778 (Page 2 of 2)',
  'This handwritten manuscript is a companion ticket sales list related to the Robert Traill document (goetzmann0184), recording additional United States Continental Lottery ticket numbers (prefixed "gom") for Class the Second. The document is heavily worn with some text obscured by damage, but appears to record approximately 25 tickets at $10 per share (2 shares each), totaling approximately $500, with a signature and location note at the bottom now largely illegible. Together with the Traill list, this document provides evidence of the multi-agent distribution network through which the Continental Congress\'s wartime lottery was sold across Pennsylvania and surrounding states in 1778.',
  {
    type: 'Receipt',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Unknown (lottery agent)',
    issueDate: '1778-01-01',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'United States Continental Lottery, Class the Second, companion ticket sales list, ca. 1778. Heavily worn; agent name illegible. Page 2 of 2.',
  }
);

// 0186: US Lottery ticket, Class the Third, No. 1A.m.655
setDoc(186,
  'United States Continental Lottery Ticket, Class the Third, No. 1A.m.655, 1776 (Page 1 of 2)',
  'This printed lottery ticket reads: "United States Lottery No. 1A.m.655 – CLASS the THIRD. THIS TICKET entitles the Bearer to receive such PRIZE as may be drawn against its Number, according to a Resolution of CONGRESS, passed at Philadelphia, November 18, 1776." Signed by "Jo: Budden" with a class indicator "S." at lower left. The Continental Congress authorized this lottery—widely regarded as the first national lottery in American history—on November 18, 1776, to help finance the Revolutionary War. The lottery was organized in classes with tickets sold at $20 each, and administered by commissioners including James Searle, John Bayard, and Daniel Roberdeau of Philadelphia. Jo: Budden served as a lottery manager responsible for signing the tickets. A later collector annotation notes this ticket as "Very rare," reflecting the extreme scarcity of surviving Continental Lottery tickets.',
  {
    type: 'Ticket',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Jo: Budden (lottery manager)',
    issueDate: '1776-11-18',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'United States (Continental) Lottery ticket, Class the Third, No. 1A.m.655. Authorized by Resolution of Congress, Philadelphia, November 18, 1776. Signed by Jo: Budden. Page 1 of 2.',
  }
);

// 0187: Reverse of US Lottery ticket – Asa Lathrop
setDoc(187,
  'United States Continental Lottery Ticket Reverse – Asa Lathrop Inscription, 1776 (Page 2 of 2)',
  'This image shows the reverse of United States Continental Lottery ticket No. 1A.m.655, Class the Third (goetzmann0186), bearing a handwritten ownership inscription: "Asa Lathrop." A later collector annotation reads "33648 [Continental Lottery]" and "Very rare!" with the lot number "45" in the upper left corner. Asa Lathrop was the original ticket holder; as a bearer instrument, the ticket could be transferred by endorsement or delivery. The annotation identifying this as a "Continental Lottery" ticket and noting its rarity reflects the scholarly consensus that very few issued Continental Lottery tickets survive in any collection. Together with the subscriber lists (goetzmann0181–0185), this ticket provides a rare material link to the first national American lottery, which raised funds for the Continental Army during the Revolutionary War.',
  {
    type: 'Ticket',
    subjectCountry: 'United States',
    issuingCountry: 'United States',
    creator: 'Jo: Budden (lottery manager)',
    issueDate: '1776-11-18',
    currency: 'USD',
    language: 'English',
    numberPages: 2,
    period: '18th Century or before',
    notes: 'Reverse of United States Continental Lottery ticket No. 1A.m.655, with ownership inscription of Asa Lathrop and collector annotation: "Very rare." Page 2 of 2.',
  }
);

// --- ANGLO-AMERICAN TONTINE PLAN, 1789: rows 188–190 ---

const tontine1789base = {
  type: 'Prospectus',
  subjectCountry: 'Great Britain',
  issuingCountry: 'Great Britain',
  creator: 'Messrs. Lockhart (bankers); trustees Francis Baring, Edmund Boehm, Thomas Henchman',
  issueDate: '1789-01-01',
  currency: 'GBP',
  language: 'English',
  numberPages: 3,
  period: '18th Century or before',
  notes: 'Anglo-American Tontine prospectus and solicitation letter, 1789. Backed by $600,000 in US federal funds. Trustees: Francis Baring, Edmund Boehm, Thomas Henchman. Bankers: Messrs. Lockhart, Pall-Mall. Solicitor: James Fonnereau.',
};

setDoc(188,
  'Anglo-American Tontine Plan, 1789: Prospectus (Page 1 of 3)',
  'This printed prospectus describes "A PLAN OF A TONTINE to consist of 1000 Shares of £.100 each, to be divided into Seven Classes, as follow; the Dividend of each Class to increase by Survivorship." Seven age classes range from lives under 10 (170 shares at £6 2s per £100) to lives above 60 (100 shares at £8 3s per £100), totaling capital of £6,679. Subscribers are to pay £50 per £100 subscribed by November 1789 to Messrs. Lockhart, Bankers, Pall-Mall, with the remaining 50% due January 1790 when each subscriber must give the name, age, and description of their nominee. To secure the annuities, the prospectus pledges $600,000 in United States government funds (equal to £135,000 sterling at 6% per annum) vested in the names of trustees Francis Baring, Edmund Boehm, and Thomas Henchman of London—plus £50,000 in Three per Cent Consolidated Annuities of Great Britain. This Anglo-American tontine, backed directly by US federal bonds, exemplifies the transatlantic financial connections between Britain and the newly independent United States formed immediately after the Revolutionary War.',
  tontine1789base
);

setDoc(189,
  'Anglo-American Tontine Plan, 1789: Solicitation Letter from James Fonnereau (Page 2 of 3)',
  'This manuscript letter, dated "London 31 Aug. 1789," solicits investment in the Anglo-American Tontine plan. The writer, James Fonnereau, refers the recipient to the enclosed printed prospectus as an "explanation of the Tontine advertised in one of the public prints by authority of the Trustees of Messrs. [Lockhart]" and urges: "If the Gentlemen or any of your friends are inclined to take any shares in the Tontine, we would be very happy to receive your kind intentions." Signed by James Fonnereau. Fonnereau was a London merchant with ties to Anglo-American commercial networks. This letter—sent together with the printed tontine plan as a solicitation package—represents the direct marketing phase of a Georgian-era financial product: an 18th-century equivalent of a private placement memorandum, distributed by personal letter to potential investors.',
  { ...tontine1789base, type: 'Letter', creator: 'James Fonnereau' }
);

setDoc(190,
  'Anglo-American Tontine Plan, 1789: Fonnereau Letter Address Panel (Page 3 of 3)',
  'This image shows the folded exterior address panel of the James Fonnereau tontine solicitation letter of August 31, 1789. The address is directed to a Mr. or Mrs. Bourgeois [?] via "Col. Nov. Lockart" in London. A red wax seal impression is visible at the center fold, and manuscript endorsements note "London Aug 31—1 Sept 1789." The document format—a letter folded and sealed into its own envelope—was standard for late-18th-century private correspondence. The address panel reveals the social network through which the Anglo-American Tontine was marketed: a London banker or agent dispatching personally addressed solicitation letters with the printed prospectus to individual subscribers, on behalf of the tontine\'s trustees. The tontine instrument, in which survivors\' annuity income grew as other subscribers died, was widely used for private investment in both Britain and France through the 18th century.',
  { ...tontine1789base, type: 'Letter', creator: 'James Fonnereau' }
);

// --- THOMAS PAINE, "DECLINE AND FALL OF THE ENGLISH SYSTEM OF FINANCE," PHILADELPHIA 1796: rows 191–202 ---

const paineBase = {
  type: 'Pamphlet',
  subjectCountry: 'Great Britain',
  issuingCountry: 'United States',
  creator: 'Thomas Paine',
  issueDate: '1796-01-01',
  currency: '',
  language: 'English',
  numberPages: 12,
  period: '18th Century or before',
  notes: 'Thomas Paine, "The Decline and Fall of the English System of Finance." Philadelphia: Printed by John Ormrod for Benjamin Franklin Bache, 1796. First American edition.',
};

setDoc(191,
  'Paine, The Decline and Fall of the English System of Finance – Cover (Page 1 of 12)',
  'This image shows the front cover wrapper of Thomas Paine\'s pamphlet "The Decline and Fall of the English System of Finance," printed in Philadelphia in 1796. The worn, water-stained paper cover bears the title in plain typography: "THE DECLINE AND FALL OF THE ENGLISH SYSTEM OF FINANCE." Library call numbers "W0.30731" and "3408" are written at the top, indicating a Yale University holding. Paine wrote this work in France in 1796 to argue that the English system of funded national debt and paper bank notes was on a mathematically calculable path to collapse—a prediction partially vindicated the following year when the Bank of England suspended cash payments in 1797.',
  paineBase
);

setDoc(192,
  'Paine, The Decline and Fall of the English System of Finance – Title Page (Page 2 of 12)',
  'This page spread shows the inner title page and opening text of Thomas Paine\'s 1796 pamphlet. The title page reads: "THE DECLINE AND FALL OF THE ENGLISH SYSTEM OF FINANCE. By THOMAS PAINE, AUTHOR OF COMMON SENSE, RIGHTS OF MAN, AGE OF REASON, &c. PHILADELPHIA, PRINTED BY JOHN ORMROD, No. 41 CHESTNUT STREET, FOR BENJ. FRANKLIN BACHE, No. 112 HIGH STREET. 1796." An epigraph reads: "On the verge, nay even in the gulph of bankruptcy"—Debates in Parliament. The text opens: "NOTHING, they say, is more certain than death, and nothing more uncertain than the time of dying: yet we can always fix a period beyond which man cannot live." Paine uses this mortality analogy to introduce his central argument: the English funding system has a calculable terminal date—its collapse is as certain as death, if unpredictable in its precise timing.',
  paineBase
);

setDoc(193,
  'Paine, The Decline and Fall of the English System of Finance – Pages 2–3 (Page 3 of 12)',
  'This page spread (pp. 2–3) develops Paine\'s opening argument about the inevitability of the English funding system\'s collapse. He identifies the core mechanism: every new war requires new loans at the same or greater interest, so the national debt multiplies geometrically with each successive conflict. Paine draws a direct parallel with the paper money systems of America (Continental currency) and France (assignats), arguing that the English Bank of England system is the same species of financial experiment, only more slowly revealing its fatal nature. He argues that the system—far from being a foundation of credit—is in fact a machine for manufacturing paper money, which must eventually depreciate to nothing, as every excess paper currency does when it exceeds the gold and silver that backs it.',
  paineBase
);

setDoc(194,
  'Paine, The Decline and Fall of the English System of Finance – Pages 4–5 (Page 4 of 12)',
  'This page spread (pp. 4–5) presents Paine\'s core mathematical argument. He traces the English national debt through six wars since the funding system began ca. 1697: the War of the Grand Alliance, the War of the Spanish Succession, the War of the Austrian Succession, the Seven Years\' War, the American War (1775–83), and the current war with France (begun 1793). For each war, Paine demonstrates that the new debt added exceeds the preceding debt, and that the interest payments absorb an ever-larger share of government revenue. He identifies a fixed common ratio between successive war debt totals and uses this ratio to project the future trajectory of the debt—establishing it as a geometric, not arithmetic, progression that must eventually exceed the capacity of any tax base to service it.',
  paineBase
);

setDoc(195,
  'Paine, The Decline and Fall of the English System of Finance – Pages 6–7 (Page 5 of 12)',
  'This page spread (pp. 6–7) presents Paine\'s comparative table showing the dramatic difference between the expenses of the first five English wars since the funding system began (totaling £424 million) and his projection for the next five wars (totaling £3,042 million)—more than seven times as great. He demonstrates that at the same geometric ratio, the English government could not possibly raise or service such debt from taxation, and that the system must collapse before five more wars are fought. Paine acknowledges the war may end sooner than his projection, but insists the mathematical ratio makes systemic collapse inevitable—"as certain as the operation of time itself"—regardless of when any particular war concludes.',
  paineBase
);

setDoc(196,
  'Paine, The Decline and Fall of the English System of Finance – Pages 8–9 (Page 6 of 12)',
  'This page spread (pp. 8–9) addresses the possibility of using a sinking fund or debt retirement to arrest the funding system\'s decline, and demonstrates why Paine considers both options futile. Every reform attempt requires new borrowing to fund the transition, and any sinking fund accumulation is overwhelmed by the geometric growth of debt added in each new war. He compares the situation to France and America: the same mechanism of paper-money inflation played out in both, though at different speeds. He argues the critical distinction is that the English system has been sustained longer by commercial strength and the creditor class\'s willingness to hold government paper, but these factors can only delay—not prevent—the inevitable depreciation.',
  paineBase
);

setDoc(197,
  'Paine, The Decline and Fall of the English System of Finance – Pages 10–11 (Page 7 of 12)',
  'This page spread (pp. 10–11) discusses the mechanics of the Bank of England and the progressive displacement of specie by paper money. Paine argues that as the funding system issues bank notes in excess of the country\'s genuine gold and silver reserves, notes gradually displace coin in circulation and inflate the prices of all goods. Those who receive fixed incomes (laborers, rentiers, annuity-holders) find their purchasing power falling relative to those who produce or trade goods. He directly compares the English situation to America and France with their paper currencies, insisting the only difference is the pace of inflation—not its ultimate consequence. He predicts that when the Bank of England is forced to suspend cash payments, the hidden inflation will become suddenly visible and prices will rapidly escalate.',
  paineBase
);

setDoc(198,
  'Paine, The Decline and Fall of the English System of Finance – Pages 12–13 (Page 8 of 12)',
  'This page spread (pp. 12–13) elaborates Paine\'s comparison of the English, French, and American paper money systems. He argues they are identical in kind: all involve governments issuing paper in quantities exceeding their gold and silver reserves, with inevitable depreciation. The English system is "more complicated" and has lasted longer due to the country\'s commercial strength and established credit relationships—but the mathematical law of compound debt accumulation means it will meet the same end. Paine identifies the moment of collapse as when the Bank of England must choose between refusing to pay cash (admitting insolvency) or attempting to pay cash and exhausting all specie reserves. Either path leads to the same outcome: the destruction of the funding system and the paper credit built upon it.',
  paineBase
);

setDoc(199,
  'Paine, The Decline and Fall of the English System of Finance – Pages 14–15 (Page 9 of 12)',
  'This page spread (pp. 14–15) focuses on the relationship between paper money, specie, and commodity prices. Paine argues that as long as paper and gold circulate together at par, prices appear stable—but this is illusory, because paper is always depreciating relative to the real value of goods. When the paper is finally separated from gold (i.e., when cash payments are suspended), the hidden inflation becomes immediately visible and prices jump to reflect the true quantity of paper in circulation. He draws on American price experience during the Revolution, noting that while Continental currency and silver circulated together at nominal par, people believed gold and silver were as plentiful as paper—until the moment of separation revealed the true ratio.',
  paineBase
);

setDoc(200,
  'Paine, The Decline and Fall of the English System of Finance – Pages 16–17 (Page 10 of 12)',
  'This page spread (pp. 16–17) addresses the mechanism of bank note circulation and the impossibility of universal cash redemption. Paine notes that England has approximately £20 million in bank notes in circulation, but neither the Bank of England nor any country bank holds enough gold and silver to redeem them simultaneously. He argues that the entire edifice of commercial credit—bills of exchange, merchants\' drafts, bank notes—rests on the common assumption that no one will actually demand cash all at once. He identifies this as a structural fragility known to "every Shopkeeper, merchant, tradesman" in London: the bank\'s ability to pay cash depends entirely on its customers not exercising that right en masse.',
  paineBase
);

setDoc(201,
  'Paine, The Decline and Fall of the English System of Finance – Pages 18–19 (Page 11 of 12)',
  'This page spread (pp. 18–19) moves toward Paine\'s conclusion. He argues that the funding system, now in the last stage of its existence, will produce a failure "total or partial" when the present war ends. A failure of the French funding system produced the French Revolution; a failure of the American system produced the US Constitution. Paine argues that a failure of the English system will similarly force fundamental political change. He maintains cool analytical distance, insisting he is not wishing for England\'s ruin but merely tracing the inevitable consequences of the geometric ratio he has identified. He also notes that the longer the war continues and the more loans are added, the more rapidly the terminal point approaches, and the more catastrophic the final failure will be.',
  paineBase
);

setDoc(202,
  'Paine, The Decline and Fall of the English System of Finance – Pages 20–21 (Page 12 of 12)',
  'This final page spread (pp. 20–21) concludes Thomas Paine\'s 1796 analysis. Paine argues that the funding system accelerates its own ruin: every new emission of paper money to service the debt further inflates prices, depresses the real value of wages and fixed incomes, and undermines commercial confidence—until the system fails. He identifies the root cause as the substitution of paper for gold and silver in excess of what the real productive economy can sustain, and argues that no financial policy can permanently save a government that has relied on borrowing beyond that capacity. The pamphlet ends without a formal peroration, relying on mathematical inevitability as its rhetorical conclusion. Within one year of publication, the Bank of England suspended cash payments (February 1797), partially vindicating Paine\'s central prediction.',
  paineBase
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 170–202 (English Lottery broadsides continued, French Tontine 1759, Continental Lottery documents 1776–1778, Anglo-American Tontine 1789, Thomas Paine pamphlet 1796).');
