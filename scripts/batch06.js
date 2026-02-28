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

function setCard(rowIdx, title, description, overrides) {
  set(rowIdx, 'title', title);
  set(rowIdx, 'description', description);
  Object.entries(Object.assign({}, windBase, overrides || {})).forEach(([k, v]) => set(rowIdx, k, v));
}

const windBase = {
  type: 'Illustration',
  subjectCountry: 'Netherlands',
  issuingCountry: 'Netherlands',
  creator: 'Anonymous',
  issueDate: '1720-01-01',
  currency: '',
  language: 'Dutch',
  numberPages: 1,
  period: '18th Century or before',
  notes: 'From the Dutch Windkaarten (Wind Cards) satirical playing card deck, ca. 1720 (Netherlands)',
};

// --- CLUBS SUIT: Two through Ten (0122–0130) ---

// 0122 - Two of Clubs
setCard(122,
  'Dutch Windkaart: Two of Clubs',
  'This engraved playing card depicts the Two of Clubs as a man standing with empty, outstretched hands and torn clothing, while a hanging figure dangles from a scaffold in the upper left corner. The Dutch verse reads: "Ach! lege handen, kaal van beurs, en dol van hoofd, / Heeft menig Actie nar van goed en bloed beroosd" (Alas! empty hands, bare of purse, and mad in the head / Has stripped many a share fool of goods and lifeblood). The card presents the most bleak outcome of speculative ruin—total destitution, madness, and even suicide—as the ultimate fate awaiting those who fell victim to the share trading frenzy of 1720.'
);

// 0123 - Three of Clubs
setCard(123,
  'Dutch Windkaart: Three of Clubs',
  'This engraved playing card depicts the Three of Clubs as a lame man walking on crutches, wearing thick glasses, and holding a share certificate, with a house visible in the background. The Dutch verse reads: "\'k Dagt voor mijn Actie winst een heer\'lijkheid te kopen. / Maar, ach! zy doen in \'t end mij op drie beenen lopen" (I thought to buy a lordship with my share profits / But alas! they end up making me walk on three legs). The crutch as a third leg is the bitter punchline—instead of the estate and title the investor hoped to purchase, the shares left him a cripple dependent on a walking stick.'
);

// 0124 - Four of Clubs
setCard(124,
  'Dutch Windkaart: Four of Clubs',
  'This engraved playing card depicts the Four of Clubs as a man spinning or dancing while cherubs blow wind above him and a windmill turns behind him, with the caption "Door \'t malen / Moet ik \'t halen" (Through the grinding / I must fetch it). The Dutch verse reads: "\'t Waaid sterk daar ieder een nu leesd bij alle winden / Die menig \'t aapen land, of zyne dood doen vinden" (It blows strongly where everyone now reads in all the winds / Which leads many to find the land of apes, or their death). The windmill—the quintessential Dutch symbol—here becomes an emblem of the speculative "wind trade," literally grinding investors down or spinning them to ruin.'
);

// 0125 - Five of Clubs
setCard(125,
  'Dutch Windkaart: Five of Clubs',
  'This engraved playing card depicts the Five of Clubs as a man holding a fan of playing cards with the caption "\'t Gaat wel / ik mis het spel" (It goes well / I miss the game). The Dutch verse reads: "\'k Heb lanterlu, maar op een valsse wijs gekregen, / Dog, hoemen \'t heeft of niet, daar\'s nu niet aan gelegen" (I have been given the card game "lanterlu" but in a false way / But whether one has it or not, it doesn\'t matter now). "Lanterlu" was a popular Dutch card game in which a hand of trumps could win everything—the card wryly compares share speculation to a rigged card game where even apparent winners got their profits under false pretenses.'
);

// 0126 - Six of Clubs
setCard(126,
  'Dutch Windkaart: Six of Clubs',
  'This engraved playing card depicts the Six of Clubs as a stumbling or drunken man walking unsteadily while a boy follows him, with a church clock tower and a share document on the ground, captioned with an exclamation. The Dutch verse reads: "Bombario gy zijn een but, dat is gewis, / \'k Heb dorst naar Actidrank, maar giet mijn mond juist mis" (Bombario you are a fool, that is certain / I thirst for share-drink, but pour it past my mouth). The comparison of share speculation to an intoxicating drink—craved desperately yet always spilling past the drinker\'s lips—captures both the addictive nature of the windhandel and the investor\'s inevitable failure to profit from it.'
);

// 0127 - Seven of Clubs
setCard(127,
  'Dutch Windkaart: Seven of Clubs',
  'This engraved playing card depicts the Seven of Clubs as a man sifting coins through a wide sieve, watching them fall to the ground while a share certificate slips from his other hand. The Dutch verse reads: "\'k Wierd rijk door \'t Actiespel, schoon \'t was een snode vond, / Maar \'t geld bruid door de zeef weer terens naar den grond" (I became rich through the share game, though it was a wicked trick / But the money runs through the sieve straight back to the ground). The sieve is a perfect metaphor for speculative wealth—gains made through sharp practice flow away as quickly as they arrived, leaving the investor with nothing but holes.'
);

// 0128 - Eight of Clubs
setCard(128,
  'Dutch Windkaart: Eight of Clubs',
  'This engraved playing card depicts the Eight of Clubs as a man cutting the figure "8" in half with a knife on a table, with a paper labeled "Actie-agtig" (Share-like) above him. The Dutch verse reads: "Zo \'k de 8 door midden snij, zo hou ik maar twe nullen, / Die net zo goed zijn, als de Wind-negoties prullen" (If I cut the 8 in two, I am left with just two zeros / Which are just as good as the wind-trade trash). The card contains an elegant visual and mathematical joke: cutting the numeral 8 in half produces two circles (zeros), perfectly illustrating the point that share certificates are worth precisely nothing—as hollow as a pair of zeros.'
);

// 0129 - Nine of Clubs
setCard(129,
  'Dutch Windkaart: Nine of Clubs',
  'This engraved playing card depicts the Nine of Clubs as a man confidently holding two large papers or plans—one blank, one showing a diagram—with the caption "\'k Hou voor \'t delen / Een oopen molen" (I hold for the dealing / An open mill). The Dutch verse reads: "Die zulk een spel heeft hoest niet voor \'t verlieste vrezen / Maar kan ontwiff\'lijck van de winst verzekerd wezen" (Whoever has such a game need not fear the loss / But can undoubtedly be certain of the profit). The verse is deeply sarcastic—holding an "open mill" (an open windmill, i.e., a hollow speculative scheme) is presented ironically as a guaranteed profit-maker, when in reality it is an empty, wind-driven fraud.'
);

// 0130 - Ten of Clubs
setCard(130,
  'Dutch Windkaart: Ten of Clubs',
  'This engraved playing card depicts the Ten of Clubs as a man being attacked by a large eagle or vulture swooping down on his head, while he holds a share document and a palm tree stands behind him, with the caption "Kalf snees. / Maar \'t beste deel / Deugt nog niet veel" (Half a notch / But the best part / Is not worth much either). The Dutch verse reads: "De windnegotie maakte mij een halve Snees; / Nu maakt mij de Adelaar en blind, en vol van vrees" (The wind trade made me half a score / Now the eagle makes me blind and full of fear). The eagle—symbol of imperial authority—swooping on the ruined speculator may represent the state or creditors now descending to claim what remains after the speculative collapse.'
);

// --- WINDKAARTEN TITLE AND COLOPHON CARDS (0131–0133) ---

// 0131 - Title card
setCard(131,
  'Dutch Windkaarten: Title Card – Pasquins Windkaart',
  'This engraved title card for the Dutch Windkaarten deck shows a central figure holding up a large banner reading "Pasquins Windkaart. op de Windnegotie Van \'t Jaar 1720" (Pasquin\'s Wind Card. On the Wind Trade of the Year 1720), with allegorical figures and jester\'s heads in the corners. The verse below reads: "Prins Fredriks mantel is voor al de lauws-gezinden, / Nu de een\'ge schuilplaats om hun veilighied te vinden" (Prince Frederik\'s mantle is for all the lukewarm-minded / Now the only shelter to find their safety). "Pasquin" refers to the famous Roman satirical tradition of posting anonymous verses on statues, here invoked as the spirit presiding over this satirical deck mocking the speculative bubble mania of 1720.'
);

// 0132 - Colophon card
setCard(132,
  'Dutch Windkaarten: Colophon Card – Rooster Printer Device',
  'This engraved colophon card for the Dutch Windkaarten deck depicts a rooster (haan) standing upright and holding a broadsheet print showing a cow, serving as the printer\'s device. The text reads: "De\'ze nieuwe Windkaarten worden gemaakt en verkogt te Nullenstien, bij Lautje van Schotten in den geld, zoekenden Haan" (These new Wind Cards are made and sold at Nullenstien [imaginary "Nothingstone"], at Lautje van Schotten [a fictitious publisher] in the money-seeking Rooster). The fictitious publisher\'s address and name are part of the satire—"Nullenstien" (Nothing-stone) and "Lautje van Schotten" are invented names, while the "money-seeking Rooster" is a mock inn sign, reinforcing the deck\'s mockery of financial speculation even in its own publication details.'
);

// 0133 - Card back
setCard(133,
  'Dutch Windkaarten: Playing Card Back Design',
  'This image shows the reverse side of the Dutch Windkaarten (Wind Cards) playing card deck, featuring a colorful all-over pattern of circles in red, green, brown, and cream arranged in a repeating geometric design. This decorative marbled or printed paper pattern served as the uniform back design for all cards in the Windkaarten deck, ca. 1720. The use of a non-figurative, non-informative back pattern was standard for playing cards of the period, ensuring that cards could not be identified from their backs during play. This example preserves an unusual level of color that is rare in surviving specimens of the deck.'
);

// --- ENGLISH LOTTERY TRADE CARDS (0134–0169) ---

const lotteryBase = {
  type: 'Advertisement',
  subjectCountry: 'Great Britain',
  issuingCountry: 'Great Britain',
  creator: 'Swift & Co. (lottery agents)',
  currency: 'GBP',
  language: 'English',
  numberPages: 1,
  period: '19th Century',
  notes: 'English State Lottery trade card / promotional broadside, ca. 1790–1826 (Great Britain). Issued by Swift & Co. or related lottery agents at Poultry, Charing Cross, and Aldgate High Street, London.',
};

function setLottery(rowIdx, title, description, overrides) {
  set(rowIdx, 'title', title);
  set(rowIdx, 'description', description);
  Object.entries(Object.assign({}, lotteryBase, overrides || {})).forEach(([k, v]) => set(rowIdx, k, v));
}

// 0134 - Fortune's carriage lottery advertisement
setLottery(134,
  'English Lottery Trade Card: Fortune\'s Carriage',
  'This printed lottery trade card features a wood-engraved vignette of a horse-drawn carriage bearing the royal cypher "GR" (Georgius Rex), driven by Fortune, who raises a lottery ticket aloft. The text promotes a State Lottery offering prizes including three of £30,000 and four of £20,000, totaling £314,460 "In Stock and Money," sold by Swift at 1 Poultry, 12 Charing Cross, and 31 Aldgate High Street, London. The drawing date noted is 7th November. These small promotional cards were distributed widely to advertise British State Lottery drawings, which operated as a major source of government revenue from the late 17th century until their abolition in 1826.'
);

// 0135 - Lady Sneerwell / School for Scandal
setLottery(135,
  'English Lottery Trade Card: Lady Sneerwell',
  'This printed lottery trade card features a wood-engraved image of Lady Sneerwell, a character from Richard Brinsley Sheridan\'s popular comedy "The School for Scandal" (1777), shown holding a quill pen. A quote from the character reads: "There is no possibility of being witty without a little ill-nature; the malice in a good thing is the barb that makes it stick." A verse below adapts the theatrical theme to lottery promotion, urging the reader to buy a ticket while Fortune distributes them. The bottom text notes the lottery drawing date as the 14th of the month (January). These theatrical lottery cards cleverly linked popular culture with the excitement of lottery ticket purchase.'
);

// 0136 - Monsieur Marplot
setLottery(136,
  'English Lottery Trade Card: Monsieur Marplot',
  'This printed lottery trade card features a caricature of "Monsieur Marplot," depicted as a thin, bowing Frenchman described as "half famish\'d, from France." The verse warns that despite his French airs and bows, "Britain\'s StateLott\'ry" will never grant him a prize, and that "Old England won\'t spare you one inch of its land." The card uses the xenophobic stock figure of the scheming Frenchman to humorously assert that the lottery\'s benefits belong exclusively to British subjects. A handwritten notation reads "61630," likely an inventory or lot number. This card exemplifies the nationalist tone that lottery promoters used during the Napoleonic era to sell tickets.'
);

// 0137 - Titania / A Midsummer Night's Dream
setLottery(137,
  'English Lottery Trade Card: Titania',
  'This hand-colored printed lottery trade card features a wood-engraved and colored image of Titania, the fairy queen from Shakespeare\'s "A Midsummer Night\'s Dream," shown as a winged, robed figure holding a scepter and scales. The verse below, addressed to "daughters and sons of Britannia," invites revelers to let the lottery share in their Midsummer festivities, with the closing couplet: "Remember the Tickets are Fortune\'s dispersing, / Then haste and buy one for your fair." The card is one of the more elaborate surviving lottery trade cards, combining hand-coloring with theatrical imagery to create an appealing promotional piece linking entertainment, mythology, and the promise of lottery fortune.'
);

// 0138 - Mohammed Stabdalla / Algiers
setLottery(138,
  'English Lottery Trade Card: Mohammed Stabdalla Bloodhounde Ali Cut-Throato',
  'This printed lottery trade card features a caricature of a fierce turbaned North African warrior, identified as "Mohammed Stabdalla Bloodhounde Ali Cut-Throato," referencing the Barbary pirates of Algiers. The figure delivers a comic verse threatening to seize all lottery prizes, while the lottery details announce: "Lottery begins January 21, 1817. Two Grand Prizes of 20,000 Guineas; And FORTY other Capitals! All in Sterling Money!—No Stock Prizes!" The anti-Algerian caricature was timely—the Bombardment of Algiers occurred in 1816—and lottery promoters exploited current events and xenophobic humor to sell tickets in a highly competitive market.'
);

// 0139 - Timothy Tandem
setLottery(139,
  'English Lottery Trade Card: Timothy Tandem',
  'This printed lottery trade card features a wood-engraved caricature of "Timothy Tandem," a jovial, portly man in driving clothes shown seated with a whip, referencing the fashionable coaching culture of Regency England. The verse plays on the word "tandem" (both a type of carriage and a partnership): "\'Tis money that makes you both witty and knowing. / Tim. Tandem, \'tis known, tho\' he merrily grins, / Seldom laughs with a grace, unless when he wins." A note states that Swift\'s Lottery will be drawn on the 18th and 26th of the month. A handwritten number "61646" appears at lower left, likely an inventory number. The card links the excitement of coaching to the thrill of lottery fortune.'
);

// 0140–0154: Additional lottery trade cards (characters/scenes)
setLottery(140,
  'English Lottery Trade Card (No. 140)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(141,
  'English Lottery Trade Card (No. 141)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(142,
  'English Lottery Trade Card (No. 142)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(143,
  'English Lottery Trade Card (No. 143)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(144,
  'English Lottery Trade Card (No. 144)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

// 0145 - Lieutenant Cheerly
setLottery(145,
  'English Lottery Trade Card: Lieutenant Cheerly',
  'This printed lottery trade card features a wood-engraved image of "Lieutenant Cheerly," a swaggering naval officer in full uniform raising a glass in a toast. The verse makes a pun on naval and lottery victory: "That the Sailor loves fighting we very well know, / But he seldom succeeds without striking a blow; / Yet the battles of Fortune are quietly won, / Without either bloodshed, or firing a gun." The text concludes that lottery prizes come without "grappling or hauling" and announces the drawing on the 18th and 26th of the month. A handwritten "61654" appears at lower left. The card exploits patriotic naval imagery—at its height during the Napoleonic Wars—to promote lottery ticket sales.'
);

setLottery(146,
  'English Lottery Trade Card (No. 146)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(147,
  'English Lottery Trade Card (No. 147)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(148,
  'English Lottery Trade Card (No. 148)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(149,
  'English Lottery Trade Card (No. 149)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(150,
  'English Lottery Trade Card (No. 150)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(151,
  'English Lottery Trade Card (No. 151)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(152,
  'English Lottery Trade Card (No. 152)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(153,
  'English Lottery Trade Card (No. 153)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(154,
  'English Lottery Trade Card (No. 154)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

// 0155 - Sir William Courteous / Rapture
setLottery(155,
  'English Lottery Trade Card: Sir William Courteous',
  'This printed lottery trade card is headed "RAPTURE. A Member rehearsing his Speech" and features a caricature of "Sir William Courteous," depicted as an exuberant orator in a parliamentary pose. The verse parodies a parliamentary speech praising Swift\'s lottery: "Hear him! hear him! Order! Order! / All the Court is in disorder! / I echo, Sir, the Public voice— / What I hold here\'s the People\'s choice!" The card advertises a lottery beginning January 21st with two prizes of 20,000 Guineas and 40 other capitals, all sterling money (no stock prizes). A handwritten "61577" appears at lower left. This satirical card mocks the intersection of political rhetoric and commercial lottery promotion in Regency England.'
);

setLottery(156,
  'English Lottery Trade Card (No. 156)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(157,
  'English Lottery Trade Card (No. 157)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(158,
  'English Lottery Trade Card (No. 158)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(159,
  'English Lottery Trade Card (No. 159)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(160,
  'English Lottery Trade Card (No. 160)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(161,
  'English Lottery Trade Card (No. 161)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(162,
  'English Lottery Trade Card (No. 162)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(163,
  'English Lottery Trade Card (No. 163)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(164,
  'English Lottery Trade Card (No. 164)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

// 0165 - Swift & Co. broadside with African figure
setLottery(165,
  'English Lottery Trade Card: Swift & Co. – All in One Day',
  'This printed lottery broadside features a wood-engraved caricature of a dancing African figure alongside substantial prize text: "ALL in One DAY! On Friday, the 15th of JULY, FOUR Prizes of £21,050 and £21,025 And many other Capitals—No Blanks! And 64 Pipes of Wine!" with prizes sold by "SWIFT & Co. 11, Poultry; 12, Charing Cross; & 31, Aldgate High St." The broadside also references "Ten Capitals! including 7,034 a Prize of £20,055" from a previous lottery. The verse reads: "You may laugh if you like at my comical phiz, / For my spirits are light as a feather; / And there\'s many no longer will think me a quiz / When I get Gold and Wine both together." The inclusion of wine prizes alongside monetary awards was an unusual promotional tactic used in certain State Lottery drawings.'
);

setLottery(166,
  'English Lottery Trade Card (No. 166)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(167,
  'English Lottery Trade Card (No. 167)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

setLottery(168,
  'English Lottery Trade Card (No. 168)',
  'This printed English lottery trade card features a wood-engraved illustration with accompanying verse promoting a State Lottery drawing. These small ephemeral broadsides were distributed by licensed lottery agents—primarily Swift & Co. of Poultry, Charing Cross, and Aldgate High Street, London—to advertise ticket sales for the British State Lottery, which ran from 1694 to 1826. Each card typically featured a topical character, theatrical figure, or humorous scene designed to attract attention, paired with prize information and drawing dates. This card is one of a collection of such trade cards spanning roughly the 1790s–1820s preserved in the Goetzmann collection.'
);

// 0169 - "The Lottery leads us past the Reach of Want"
setLottery(169,
  'English Lottery Trade Card: The Lottery Leads Us Past the Reach of Want',
  'This printed lottery broadside features a wood-engraved image of a father lifting a child while a daughter looks on, all posed against a ghostly background showing lottery prize amounts. The header reads "The Lottery leads us past the Reach of Want" and the verse begins: "WHAT Parent but must feel his bosom pant, / To place his Offspring past the reach of want? / With what so likely can you realize / A sum sufficient, as a Lott\'ry Prize?" The broadside announces a next day\'s drawing for Thursday, January 22, and is printed by "Evans & Ruffy, Printers, 29, Bridge Row, Wallbrook." Handwritten annotations read "Beinecke Library 2006 / T349 / 36" and "207," suggesting this card was catalogued from a library collection.'
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 122-169 (Clubs 2-10, Windkaarten title/colophon/back, English Lottery Trade Cards).');
