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

function setCard(rowIdx, title, description) {
  set(rowIdx, 'title', title);
  set(rowIdx, 'description', description);
  Object.entries(windBase).forEach(([k, v]) => set(rowIdx, k, v));
}

// --- HEARTS SUIT: Ace through Ten (0101–0110) ---

// 0101 - Ace of Hearts
setCard(101,
  'Dutch Windkaart: Ace of Hearts',
  'This engraved playing card depicts the Ace of Hearts (1) as a man fishing with a drum slung over his back, while a mermaid—the "Zuidzé Zesierzen" (South Sea Siren)—rises from the water holding a mirror. The Dutch verse reads: "De Zuidzé Zesierzen dagt ik vergeefs te vangen / Zy knapt het geld, en laat de lege hengel hangen" (The South Sea Siren I thought in vain to catch / She snaps up the money and leaves the empty fishing rod hanging). The card personifies the South Sea Company as an alluring but deceptive mermaid who seizes investors\' money and leaves them holding nothing.'
);

// 0102 - Two of Hearts
setCard(102,
  'Dutch Windkaart: Two of Hearts',
  'This engraved playing card shows the Two of Hearts as a man carrying a basket of goods and raising a burning torch, with the caption "Zy drijven op en onder" (They drift up and down). The Dutch verse reads: "Mijn harten hobbèlende op \'t verdwynend Bubbel water, / Zijn als al \'t Actiewerk, zo vals gelijk een Sater" (My hearts bobbing on the disappearing Bubble water / Are like all the share trading, as false as a satyr). The card compares speculative share trading to goods bobbing helplessly on vanishing bubble water—ultimately as false and treacherous as a mythological satyr.'
);

// 0103 - Three of Hearts
setCard(103,
  'Dutch Windkaart: Three of Hearts',
  'This engraved playing card depicts the Three of Hearts as a stout man in working clothes holding up a coin in one hand and a paper in the other, with the caption "Een ryke vrek / Is dubbeld gek" (A rich miser / Is doubly mad). The Dutch verse reads: "Hoe hoog mijn kapitaal door de Acties is gerezen / \'k Ben door mijn gierigheid niet rijker als voor dezen" (However high my capital has risen through the shares / Through my greed I am no richer than before). The card moralizes that greed defeats its own purpose—a miser who profits from shares but cannot enjoy or deploy his gains is doubly foolish.'
);

// 0104 - Four of Hearts
setCard(104,
  'Dutch Windkaart: Four of Hearts',
  'This engraved playing card depicts the Four of Hearts as a well-dressed man standing at the seashore with the caption "\'t Gaat alles naar de zé" (Everything goes to the sea), tipping a small chest as coins fall into the water where a dragon or sea-serpent lurks. The Dutch verse reads: "Om al mijn geld door de Acties niet te missen, / Gooi ik \'t in zé, tot lokäas voor de vissen" (To keep from losing all my money through the shares / I throw it in the sea as bait for the fish). The card wryly suggests that throwing one\'s money into the sea is no worse than investing it in share companies, since both result in total loss.'
);

// 0105 - Five of Hearts
setCard(105,
  'Dutch Windkaart: Five of Hearts',
  'This engraved playing card shows the Five of Hearts as a jester in motley costume dancing and holding a five-pointed star, with a bill on his cap reading "Vijf Sin te koop" (Five senses for sale). The Dutch verse reads: "Ik ben niet gekker als al de and\'ren, die met hopen, / Ook hun vijf zinnen voor een Actiebrief verkopen" (I am no madder than all the others who with hope / Also sell their five senses for a share certificate). The caption "Daar ieder tragt om geld te winnen, Geef ik om geld ook mijn vijf sinnen" (Since everyone tries to win money, I give for money also my five senses) frames universal speculative folly as a kind of mass jester-like madness.'
);

// 0106 - Six of Hearts
setCard(106,
  'Dutch Windkaart: Six of Hearts',
  'This engraved playing card depicts the Six of Hearts as a man carrying a balance scale—weighing paper shares against coin—with a monkey at his feet and a rising or setting sun behind him, captioned "Ryne Crenaar / Toont alles klaar" (Rhine merchant / Shows everything clearly). The Dutch verse reads: "Door \'t Actie werk te dol te vatten bij der hand, / Gaat men de Linie door naar \'t gloeyend Aapenland" (By taking the share trading too madly by the hand / One crosses the line into the glowing Land of Apes). The "Land of Apes" (Aapenland) is a satirical image of the foolish tropical country one metaphorically reaches when greed drives all reason away.'
);

// 0107 - Seven of Hearts
setCard(107,
  'Dutch Windkaart: Seven of Hearts',
  'This engraved playing card depicts the Seven of Hearts as an astrologer in a long robe gazing through a cross-shaped instrument at the stars, with a dog running in the background and the caption "Astrologist / der Windnegotie" (Astrologer of the Wind Trade). The Dutch verse reads: "Om Actie voordeel kijk ik ijv\'rig in de starren, / Maar hoe ik meerder zie, hoe \'k meerder raak aan \'t warren" (To find share profit I eagerly look at the stars / But the more I see, the more confused I become). The card mocks those who tried to use astrology or market divination to predict share prices, satirizing the pseudo-scientific pretensions of speculative investment advice.'
);

// 0108 - Eight of Hearts
setCard(108,
  'Dutch Windkaart: Eight of Hearts',
  'This engraved playing card depicts the Eight of Hearts as an old woman seated among palm trees watching a fox rummage in a chest and a monkey beside her, with the caption "\'k leer Vossen en aapen / Om geld te schrapen" (I teach foxes and monkeys / To scrape together money). The Dutch verse reads: "\'t Is vrugteloos op winst gedagt, / Daar de Aep en Vos het voordeel wagt" (It is futile to think of profit / Where the Ape and Fox await the advantage). The fox and monkey are traditional symbols of cunning and mimicry, representing the dishonest brokers and imitators who exploited naive investors in the share market.'
);

// 0109 - Nine of Hearts
setCard(109,
  'Dutch Windkaart: Nine of Hearts',
  'This engraved playing card depicts the Nine of Hearts as a man who has fallen beneath a bell-and-rope construction—possibly a pillory or tollgate—with papers flying and a lantern dropped on the ground, captioned "Ai my wat smak" (Oh my, what a fall). The Dutch verse reads: "Door \'t breken van de koord ben ik bedroefd gevallen, / dog \'t is mijn loon; ik moest niet de Actieklok niet mallen" (By the breaking of the rope I have sadly fallen / But it is my reward; I should not have played with the share-bell). The "Actieklok" (share-bell) likely refers to the exchange bell that signaled trading sessions, and the broken rope is the investor\'s downfall, brought on by his own foolish participation.'
);

// 0110 - Ten of Hearts
setCard(110,
  'Dutch Windkaart: Ten of Hearts',
  'This engraved playing card depicts the Ten of Hearts as a man sitting on a log and blowing soap bubbles with a long pipe, a mirror lying beside him, with the caption "Windbellen om Nul / te koop" (Wind-bubbles for nothing / for sale). The Dutch verse reads: "\'t Is Wind en Nul en anders niet, / Gelijk men klaar op \'t einde ziet" (It is Wind and Nothing and nothing else / As one clearly sees at the end). The image distills the Windkaarten deck\'s central message: speculative shares are nothing but wind and bubbles, worth zero—as anyone could see once the bubble had burst.'
);

// --- DIAMONDS SUIT: Ace through Ten (0111–0120) ---

// 0111 - Ace of Diamonds
setCard(111,
  'Dutch Windkaart: Ace of Diamonds',
  'This engraved playing card depicts the Ace of Diamonds as the biblical Samson tearing apart a lion with his bare hands, captioned "Actie raadzel" (Share riddle). The Dutch verse reads: "Nog blinder was \'t geheim van de Acties, dan voorheen / \'t Diepzinnig raadzel van den dapp\'ren Samson scheen" (Even more blind was the secret of the Shares / than the deep riddle of brave Samson seemed). Samson\'s famous riddle ("Out of the eater came something to eat") parallels the inscrutable mystery of share values, suggesting that the inner workings of speculative companies were as opaque and dangerous as a lion\'s carcass.'
);

// 0112 - Two of Diamonds
setCard(112,
  'Dutch Windkaart: Two of Diamonds',
  'This engraved playing card depicts the Two of Diamonds as a peddler wearing thick spectacles and carrying a case of eyeglasses for sale, with a bat flying overhead and an owl perched on a post, captioned "Koop brillen voor blinde Actionisten" (Buy glasses for blind shareholders). The Dutch verse reads: "Koop Brillen voor de blind gehuilde Actionisten, / Die eerst in de uil vlegt hun geld en goed verkwisten" (Buy glasses for the blindly weeping shareholders / Who first in the owls\' flight squandered their money and goods). The bat and owl—both animals associated with blindness and foolishness—stand for investors who acted in the dark and lost everything before seeing the truth.'
);

// 0113 - Three of Diamonds
setCard(113,
  'Dutch Windkaart: Three of Diamonds',
  'This engraved playing card depicts the Three of Diamonds as an armored knight bearing a coat of arms with crossed spades and an anchor, with the caption "Hoe dol de wind ook gilden; / Hy stuit op deze schilden" (However madly the wind may scream; / It halts at these shields). The Dutch verse reads: "Dit trits der steden houd de Koopgod in zijn stand, / Omhelsd Neptuin, en stut het vallend Vaderland" (This trio of cities keeps the God of Commerce in his position / Embracing Neptune, and props up the falling Fatherland). The card presents a counterpoint of civic virtue and stability—the three cities (perhaps Amsterdam, Rotterdam, and Middelburg) standing firm against the speculative storm.'
);

// 0114 - Four of Diamonds
setCard(114,
  'Dutch Windkaart: Four of Diamonds',
  'This engraved playing card depicts the Four of Diamonds as a mariner holding a compass and a flag with a ship visible at sea in the background, captioned "\'k Zoek Zuid nog West. / Naar \'t oosten \'t best" (I seek neither South nor West. / Eastward is best). The Dutch verse reads: "Het Zuid en \'t West heeft zo veel kwaad gedaan, / Dat nooit hun naam wéér op \'t Kompas moest staan" (The South and the West have done so much harm / That their names should never again appear on the Compass). The card explicitly condemns the South Sea Company and the West Indies companies, calling for their names to be struck from the navigator\'s compass as disgraced directions.'
);

// 0115 - Five of Diamonds
setCard(115,
  'Dutch Windkaart: Five of Diamonds',
  'This engraved playing card depicts the Five of Diamonds as a man leaning over a table examining papers while a surveyor with instruments works in the background, captioned "Kromme Cinquen" (Crooked Fives). The Dutch verse reads: "De kromme Cinquen, daar ik mij op durf vertrouwen, / Die zullen, vrees ik, mij in \'t eind\' nog eens berouwen" (The crooked fives that I dare trust myself to / I fear will in the end cause me regret once more). "Kromme Cinquen" refers to the manipulated or falsified accounts in share subscription books, suggesting that dishonest accounting in company prospectuses would ultimately bring regret to trusting investors.'
);

// 0116 - Six of Diamonds
setCard(116,
  'Dutch Windkaart: Six of Diamonds',
  'This engraved playing card depicts the Six of Diamonds as a gentleman in fine clothes holding a prospectus or share document, with horse reins lying slack on the ground and a carriage wheel visible behind him, captioned "Daar vlieqt koets en Paarden" (There fly carriage and horses). The Dutch verse reads: "Mijn koets en paarden zijn mij door den wind ontglipt. / Terwyl de teugel ook op \'t laast mijn hand ontslipt" (My carriage and horses have slipped away from me in the wind / While the reins too have finally slipped from my hand). The lost carriage and horses represent the tangible wealth—the trappings of prosperity—that speculative investors squandered chasing wind-company shares.'
);

// 0117 - Seven of Diamonds
setCard(117,
  'Dutch Windkaart: Seven of Diamonds',
  'This engraved playing card depicts the Seven of Diamonds as a man holding a large, thick ledger book aloft with one hand and a paper under the other arm, captioned "\'t Duister Actieboek" (The Dark or Obscure Share Book). The Dutch verse reads: "Zo min als ik den loop der zeven hoofd Planeten / Versta; zo min kan ik het eind der Acties weten" (As little as I understand the course of the seven main planets / So little can I know the outcome of the shares). The card compares the impenetrable complexity of share company accounts to the mysterious movements of the planets—both equally beyond the ordinary person\'s comprehension and equally unpredictable.'
);

// 0118 - Eight of Diamonds
setCard(118,
  'Dutch Windkaart: Eight of Diamonds',
  'This engraved playing card depicts the Eight of Diamonds as a magistrate or clergyman in formal robes holding a paper in each outstretched hand, with the caption "\'k Ben de agtbaarheid. / Door mus druk kryt" (I am respectability. / Through [press] print I cry out). The Dutch verse reads: "Mijn Agtbaarheid is met de Zuid, / Door schraap- en woekerzugt verbruid" (My respectability has been ruined with the South [Sea Company] / Through greed and usury). The card censures members of the respectable establishment—magistrates, clergy, officials—who compromised their dignity and authority by participating in or enabling the speculative bubble.'
);

// 0119 - Nine of Diamonds
setCard(119,
  'Dutch Windkaart: Nine of Diamonds',
  'This engraved playing card depicts the Nine of Diamonds as a man bowing with his hat in hand, holding a paper printed with rows of pawns or chess pieces, captioned "Voor den Koning" (For the King). The Dutch verse reads: "Sta ruin; want zook eens mag den Actie-koning raken, / Lag ik met Zuid en West om mijn fortuin te maken" (Stand ruined; for even the share-king may fall / I relied on South and West to make my fortune). The chess imagery—pawns before a king—suggests the expendable rank-and-file investors sacrificed in the speculative game played by the great South Sea and West Indies companies.'
);

// 0120 - Ten of Diamonds
setCard(120,
  'Dutch Windkaart: Ten of Diamonds',
  'This engraved playing card depicts the Ten of Diamonds as a man in worn clothing offering his ring for sale while holding a share certificate, with the caption "Wie koopt er mijn laasten Ring" (Who buys my last ring). The Dutch verse reads: "Wie stenen koopt bij nagt laat zig altoos bedotten, / Zo ging \'t in Quinquempoix ook met al de Actie zotten" (Whoever buys stones by night always lets himself be fooled / So it went in Quinquampoix with all the share fools too). The man reduced to selling his last piece of jewelry encapsulates the complete financial ruin of the ordinary investor, while the reference to Quinquampoix links this fate directly to John Law\'s Parisian speculative market.'
);

// --- CLUBS SUIT: Ace (0121) ---

// 0121 - Ace of Clubs
setCard(121,
  'Dutch Windkaart: Ace of Clubs',
  'This engraved playing card depicts the Ace of Clubs as an allegorical pilgrim figure robed in a star-patterned cloak with a sun emblem on his chest, holding a staff and extending a blank paper, captioned "Pelgrim van de waarheid" (Pilgrim of Truth). The Dutch verse reads: "Ik zoek naar waarheid, en naar wijsheid; maar, helaas! / Waar dat ik kom, ik vind de wereld even dwaas" (I seek after truth and after wisdom; but alas! / Wherever I come, I find the world equally foolish). The Pilgrim of Truth—a classical allegorical figure—wanders the world in search of wisdom but finds only universal folly, a damning verdict on the speculative madness that gripped Europe in 1720.'
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 101-121 (Hearts Ace-10, Diamonds Ace-10, Clubs Ace).');
