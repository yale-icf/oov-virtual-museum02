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

// 0079 - King of Spades
setCard(79,
  'Dutch Windkaart: King of Spades',
  'This engraved playing card depicts the King of Spades (Heer) from a Dutch satirical card deck mocking the speculative bubble mania of 1720. A gentleman in fashionable dress stands holding a balance scale, weighing a shovel and sheet of paper against a purse of gold coins. The Dutch verse reads: "De schop, en\'t blad papier in de eene schaal gelegen, / Kan, door de braversshoop, een goud beurs overwegen" (The shovel and a sheet of paper in one pan of the scale / Can, through a gambler\'s luck, outweigh a purse of gold). The card satirizes speculative investment in stock company shares ("Acties"), suggesting that worthless paper backed only by luck could appear to rival real wealth.'
);

// 0080 - Queen of Spades
setCard(80,
  'Dutch Windkaart: Queen of Spades',
  'This engraved playing card depicts the Queen of Spades (Vrouw) from a Dutch satirical card deck ridiculing the 1720 stock market frenzy. A noblewoman stands with one hand on a spade driven into the ground while her other hand releases a shower of falling playing cards. The Dutch verse reads: "Men zet de schop vrij neer, ons spelen is verbruid / De blinde razernij van Quinquampoix heeft uit" (We put the shovel firmly down, our playing is over / The blind madness of Quinquampoix is done). The reference to "Quinquampoix" alludes to Rue Quincampoix in Paris, the street where John Law\'s Mississippi Company shares were traded.'
);

// 0081 - Jack of Spades
setCard(81,
  'Dutch Windkaart: Jack of Spades',
  'This engraved playing card shows the Jack of Spades (Knegt) as a young drummer marching with scattered papers flying around him, captioned "Prins Fredriks mars, naar Viauen" (Prince Frederik\'s march, to Vianen). The Dutch verse reads: "Ik slaprins Fredriks mars, om met al de Actie gekken / Als wind zoldaten, naar Viauen toe te trekken" (I beat Prince Frederik\'s march, to lead all the share-mad fools / Like wind soldiers, toward Vianen). The card satirizes the frenzied speculation that drove investors like soldiers following a futile march, referencing real Dutch events and persons of the period.'
);

// 0082 - King of Hearts
setCard(82,
  'Dutch Windkaart: King of Hearts',
  'This engraved playing card depicts the King of Hearts (Heer) labeled "Lauwer Koning" (Laurel King), likely a satirical reference to John Law, the Scottish financier behind France\'s Mississippi Company scheme. The man stands crowned with laurels while workers dig in the background. The Dutch verse reads: "\'k Heb de Acties eerst bedagt, maar een kwade Actiekans; / Smit al mijn lof omver, en schent mijn lauwer krans" (I first conceived the shares, but a bad share-chance / Overturns all my praise and dishonors my laurel wreath). The card mocks the collapse of Law\'s system, which contributed to the speculative bubble fever of 1720 across Europe.'
);

// 0083 - Queen of Hearts
setCard(83,
  'Dutch Windkaart: Queen of Hearts',
  'This engraved playing card shows the Queen of Hearts (Vrouw) holding a rolled architectural plan, with the motto "De Actie ketel distileerd, / Dat het papier in Goud verkeerd" (The stock kettle distills / That paper turns into gold) inscribed beside her. The Dutch verse reads: "Mijn lauwerman volleerd om ieder te bedotten, / Stookt goud uit klad papier, en maakt \'t Heel-Al rol zotten" (My laureate perfects the art of fooling everyone / Distills gold from scrap paper and drives the whole world mad). The card satirizes the alchemical pretensions of stock promoters who claimed to transform worthless paper shares into real wealth.'
);

// 0084 - Jack of Hearts
setCard(84,
  'Dutch Windkaart: Jack of Hearts',
  'This engraved playing card depicts the Jack of Hearts (Knegt of Boef, meaning "Knave or Rogue") carrying a spade on his shoulder with the label "Uit Wanhoop" (Out of Despair) beside him. The Dutch verse reads: "Mijn boeve hart, \'t geen door een reeks van Actiestreken. / My deed naar voordeel zien, word nu van spijt doorsteken" (My roguish heart, which through a series of share tricks / Made me see profit, is now pierced by regret). The card portrays the common speculator left ruined after the collapse of the share trading schemes, reduced to menial labor and overcome with despair.'
);

// 0085 - King of Diamonds
setCard(85,
  'Dutch Windkaart: King of Diamonds',
  'This engraved playing card shows the King of Diamonds (Heer) standing in an open doorway with the motto "Ik ben beschermd. / Daar ieder kermd" (I am protected / While everyone else laments). The Dutch verse reads: "\'k Ben voor de Zuid en West hun windgeblaas beschut / Terwyl een laauw Geest \'t Heel Al helpt m den dut" (I am sheltered from the wind-blowing of the South and West [companies] / While a lukewarm Spirit sends the whole world to sleep). The card directly references the South Sea Company (Britain) and the West Indies trading companies, mocking those who thought themselves immune from the bubble\'s collapse.'
);

// 0086 - Queen of Diamonds
setCard(86,
  'Dutch Windkaart: Queen of Diamonds',
  'This engraved playing card depicts the Queen of Diamonds (Vrouw) holding a long clay pipe that produces a large cloud of smoke, with an oven behind her and diamond-shaped papers flying through the air. The Dutch verse reads: "Door sterk te blazen breekt het glas, / Net als \'t met de Actie-bubbels was" (By blowing hard the glass breaks / Just as it was with the stock bubbles). The card is a visual pun on the word "bubble"—both the glassblower\'s fragile creation and the financial bubble—suggesting that excessive speculative "blowing" inevitably causes collapse.'
);

// 0087 - Jack of Diamonds
setCard(87,
  'Dutch Windkaart: Jack of Diamonds',
  'This engraved playing card shows the Jack of Diamonds (Knegt) as a pipe-smoking man holding a share certificate being consumed by fire from the sun, with the caption "\'t Papier zo duur / Werd rook door \'t vuur" (The paper so dear / Became smoke through the fire). The Dutch verse reads: "Mijn Actiebrief verteerd, hij kan geen zon verdragen; / \'t Is vuur voor kleine tijds, als al de Bubbelv lagen" (My share certificate consumed, it cannot bear the sun / It is fire for a brief moment, like all the Bubble papers). The card satirizes the ephemeral nature of share certificates, which like paper ignited by sunlight turned to smoke and worthless ash.'
);

// 0088 - King of Clubs
setCard(88,
  'Dutch Windkaart: King of Clubs',
  'This engraved playing card depicts the King of Clubs (Heer) with the caption "Eerst gelukkig, / Nu drukkig" (First lucky / Now miserable) beside him, while a monkey operates a funnel behind him and soap bubbles float in the air. The Dutch verse reads: "\'k Ben Directeur geweest, maar tot mijn ongeluk, / Het geld droop door de zak, en liet mij niets als druk" (I was a Director, but to my misfortune / The money dripped through the bag and left me nothing but misery). The card mocks company directors who profited briefly from the bubble before its collapse, leaving them as destitute as the investors they had exploited.'
);

// 0089 - Queen of Clubs
setCard(89,
  'Dutch Windkaart: Queen of Clubs',
  'This engraved playing card shows the Queen of Clubs (Vrouw) with the caption "Met Ikarus, en Faéton, / Verdwynt mijn man zijn Actie zon" (With Icarus and Phaeton / My husband\'s share-sun disappears), while a falling Icarus figure appears in the background. The Dutch verse reads: "Schoon \'t lot van Ikarus mijn man valt tot een deel, / Ik hou nog moet, zo lang ik met de klavers speel" (Although the fate of Icarus befalls my husband to some degree / I still have courage as long as I play with clubs). The card uses the classical myths of Icarus and Phaeton—who flew too close to the sun and fell—as metaphors for financial overreach and speculation gone wrong.'
);

// 0090 - Jack of Clubs
setCard(90,
  'Dutch Windkaart: Jack of Clubs',
  'This engraved playing card depicts the Jack of Clubs (Knegt) as a man freed from chains, with the caption "\'t Geluk brengt mij / Uit slavernij" (Luck brings me / Out of slavery), while a bowing beggar stoops beside him. The Dutch verse reads: "Mijn Actiestar heeft mij van knegt tot Heer gemaakt, / Schoon menig Heer daar door tot beed\'len is geraakt" (My share-star has made me from servant to lord / Although many a lord has thereby been reduced to begging). The card satirizes the social inversion caused by stock speculation, in which servants could become wealthy while noblemen lost everything.'
);

// 0091 - Ace of Spades
setCard(91,
  'Dutch Windkaart: Ace of Spades',
  'This engraved playing card depicts the Ace of Spades (1) as a common laborer sweeping the street with a broom and spade, surrounded by discarded share certificates, with the caption "Men maakt zig \'t ruils kwyt / Het stinkt zelf daar het leid" (One rids oneself of the exchange / It stinks even where it lies). The Dutch verse reads: "Weg met dit kladpapier, \'t geen ieder een verlegen / Op straat werpt, wy\'lik\'t meen in lethes-poel te vegen" (Away with this scrap paper which everyone, troubled, throws in the street / I think I\'d rather sweep it into Lethe\'s pool). The reference to Lethe, the river of forgetfulness in the underworld, expresses a wish to erase the memory of the whole financial catastrophe.'
);

// 0092 - Two of Spades
setCard(92,
  'Dutch Windkaart: Two of Spades',
  'This engraved playing card shows the Two of Spades as a mythological winged figure burning share certificates beside a fire, with the caption "Door my word\'t eind gezien / Van al de Compagnien" (Through me the end is seen / Of all the Companies). The Dutch verse reads: "Ik ben\'t die \'t al ondekt, en\'t klad papier verbrand, / De vonken doof, en help den Koopman weer in stand" (I am the one who discovers all and burns the scrap paper / Quenches the sparks, and helps the merchant recover). The card presents a figure of truth or revelation who exposes the fraud of the bubble companies and helps commerce return to solid footing.'
);

// 0093 - Three of Spades
setCard(93,
  'Dutch Windkaart: Three of Spades',
  'This engraved playing card depicts the Three of Spades as three allegorical women floating in the air upon a single spade or shovel, representing the three great financial bubbles of 1720. The Dutch verse reads: "Zie hoe de Zuid, en West, en Missi-sippi zweven / Op eene schop, wyl zy van lugt- en winden leven" (See how the South [Sea], and West [Indies Company], and Mississippi float / On a shovel, while they live on air and winds). The card explicitly names England\'s South Sea Company, a West Indies trading company, and John Law\'s Mississippi Company as empty fantasies sustained by nothing but hot air and speculation.'
);

// 0094 - Four of Spades
setCard(94,
  'Dutch Windkaart: Four of Spades',
  'This engraved playing card depicts the Four of Spades as a gentleman dangling an upside-down cat while another animal lies fallen on the ground nearby. The Dutch verse reads: "Hoe schreeuwd de Zuid ze-dog zo schrikk\'lijk, daar mijn kat / Naauw maauwd, of heeft hij ligt meer van de zweep gehad" (How the South Sea dog screams so terribly, while my cat / Has barely mewed, or has perhaps gotten more of the whip). The card uses a domestic scene—a cat and a shrieking dog—to satirize the contrasting fortunes of investors in the South Sea Company versus Dutch bubble schemes, with the implication that Dutch speculators suffered more quietly.'
);

// 0095 - Five of Spades
setCard(95,
  'Dutch Windkaart: Five of Spades',
  'This engraved playing card depicts the Five of Spades as a man holding up a spade decorated with the Amsterdam coat of arms (three X\'s) while cherubs blow trumpets overhead. The Dutch verse reads: "Zie hoe de vijfde stad van Holland, de Acties plet, / Waarom een lauwerkrans mij werd op \'t hoofd gezet" (See how the fifth city of Holland crushes the shares / For which a laurel wreath was placed on my head). The Amsterdam coat of arms identifies this as a boastful reference to Amsterdam\'s civic identity, with the city claiming credit—and a victor\'s laurel—for its role in resisting or ending the bubble speculation.'
);

// 0096 - Six of Spades
setCard(96,
  'Dutch Windkaart: Six of Spades',
  'This engraved playing card depicts the Six of Spades as a man sitting on a chest with the caption "Een Sesje min of meer / Scheeld weinig aan mijn eer" (A six more or less / Makes little difference to my honor). The Dutch verse reads: "Mijn rotte kist doet mij thans met sesje pronken; / \'k Behouw het geld wel, maar mijn eer, die is verzonken" (My rotted chest makes me now flaunt with a six / I keep the money, but my honor has sunk). The card satirizes a speculator who salvaged some money from the bubble crash but has permanently lost his reputation—the rotten chest on which he sits symbolizes the decayed moral foundation of his ill-gotten gains.'
);

// 0097 - Seven of Spades
setCard(97,
  'Dutch Windkaart: Seven of Spades',
  'This engraved playing card depicts the Seven of Spades as a man leaning over a counting table, with the caption "Drie en vier / Betaald altier" (Three and four / Paid here in full). The Dutch verse reads: "Zie de een en twe, en vijf en zes gaan vrij, maar ach! / De drie en vier betaald in Quinquempoix \'t gelag" (See the one and two, and five and six go free, but alas! / The three and four paid the bill at Quinquampoix). The card references the famous trading address of Rue Quincampoix in Paris, where John Law\'s bubble was centered, suggesting that unlucky investors paid a heavy price at that address while others escaped unscathed.'
);

// 0098 - Eight of Spades
setCard(98,
  'Dutch Windkaart: Eight of Spades',
  'This engraved playing card depicts the Eight of Spades as a man holding a white bird (falcon) and a lantern beside an hourglass-like cage. The Dutch verse reads: "Diogenes lantaarn heb ik bij dag van noden; / \'k Vong de agtings Witte Valk die and\'ren is ontvloden" (I need Diogenes\' lantern by day / I caught the honorable White Falcon that others have fled). The reference to Diogenes—who searched for an honest man by lantern in daylight—suggests that honest dealing has become virtually impossible to find in the speculative market, while the "Witte Valk" (White Falcon) may refer to a specific company or investment opportunity that proved elusive to most investors.'
);

// 0099 - Nine of Spades
setCard(99,
  'Dutch Windkaart: Nine of Spades',
  'This engraved playing card depicts the Nine of Spades as a jester or Fool holding a mirror in one hand and a large satirical broadsheet in the other, with a snake at his feet and a spade nearby. The Dutch verse reads: "De dwaasheid dogt de kunst te delven in het graf, / Maar vrouw Voorzigtigheid keerd deze rampen af" (Folly thought to bury the art in the grave / But Lady Prudence averted these disasters). The card personifies speculative folly as a jester (dwaasheid = foolishness) contrasted with the virtue of Prudence (Voorzigtigheid), suggesting that only caution and wisdom could avert the financial ruin that speculation had wrought.'
);

// 0100 - Ten of Spades
setCard(100,
  'Dutch Windkaart: Ten of Spades',
  'This engraved playing card depicts the Ten of Spades as a man in a plain coat holding a coin aloft and carrying a portfolio of share papers under his arm. The Dutch verse reads: "Qui donne moiun sou, zo zult ge aanstonds aanschouwen / Wat grafstee, dat de tijd voor de Actieprins zal bouwen" (Whoever gives even half a sou, you will soon behold / What gravestone time will build for the share prince). The bilingual caption—mixing French ("Qui donne") with Dutch—alludes to the cross-border speculative market, and predicts a tombstone (grafsteen) as the ultimate monument to the Actieprins (share prince), a term for the bubble profiteer.'
);

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 79-100 (Dutch Windkaarten, court cards and Spades suit).');
