const XLSX = require('xlsx');
const wb = XLSX.readFile('oov_data_new.xlsx');
const ws = wb.Sheets[wb.SheetNames[0]];

const rows = [
  [715,'goetzmann0715.jpg','/images/goetzmann0715.jpg',
   'Republic of Poland, Series III Premium Dollar Loan Bond (Obligacja Serji III Premjowej Pozyczki Dolarowej), $5 / 44.57 Zlotych, No. 1222277, Warsaw, 1931',
   'Bearer bond (Obligacja) of the Republic of Poland (Rzeczpospolita Polska) for the Series III Premium Dollar Loan (Serja III Premjowa Pozyczka Dolarowa), No. 1222277. Denomination: 5 US dollars (Stanow Zjednoczonych Ameryki) = 44.57 Polish zlotych. Issued under the law of 2 January 1930. Interest-bearing bond at 4% per annum, payable semi-annually on 1 February and 1 August. Eligible for premium lottery draws held publicly at the Ministry of Finance. Redeemable on 1 February 1941. Warsaw, 1 February 1931. Signed by the Director of the Department of Monetary Affairs, the Head of the Ministry of Finance, and the Commission of State Debt Control. Blue decorative border. Printed by Polska Wytwornia Papierow Wartosciowych, Warsaw.',
   'bond',
   'Republic of Poland; Rzeczpospolita Polska; Dollar Loan; Obligacja; Series III; Premium Loan; $5; 44.57 zlotych; No. 1222277; Warsaw; 1931; Pozyczka Dolarowa',
   'Poland','Poland',
   'Rzeczpospolita Polska / Ministerstwo Skarbu','1931-02-01','United States dollar; Polish zloty','Polish','1','Early 20th century','','','','',''],

  [716,'goetzmann0716.jpg','/images/goetzmann0716.jpg',
   'Republic of Poland, Series III Premium Dollar Loan Bond (Obligacja Serji III Premjowej Pozyczki Dolarowej), $5 / 44.57 Zlotych, No. 1450669, Warsaw, 1931, with coupon stubs',
   'Bearer bond (Obligacja) of the Republic of Poland for the Series III Premium Dollar Loan, No. 1450669. Denomination: 5 US dollars = 44.57 Polish zlotych. Same issue as No. 1222277 (goetzmann0715). Three remaining coupon stubs (coupons 16, 18, 20) attached at bottom, indicating earlier coupons have been redeemed. Warsaw, 1 February 1931. Signed by authorized officers of the Ministry of Finance.',
   'bond',
   'Republic of Poland; Rzeczpospolita Polska; Dollar Loan; Obligacja; Series III; Premium Loan; $5; 44.57 zlotych; No. 1450669; Warsaw; 1931; coupon stubs; Pozyczka Dolarowa',
   'Poland','Poland',
   'Rzeczpospolita Polska / Ministerstwo Skarbu','1931-02-01','United States dollar; Polish zloty','Polish','1','Early 20th century','','','','',''],

  [717,'goetzmann0717.jpg','/images/goetzmann0717.jpg',
   'Conversion of the External Debt of Portugal, Provisional Certificate of Deferred Stock, Three Pence, No. 4463, London, ca. 1853',
   'Provisional Certificate of Deferred Stock for the Conversion of the External Debt of Portugal, issued under the Decree of 18 December 1852. No. 4463. In conformity with Article 25 of the Regulations approved by a Decree of Her Most Faithful Majesty the Queen of Portugal (23 March 1853), the Financial Agency in London is authorised to carry out the Conversion. Holder entitled to Three Pence of Three Per Cent. Deferred Stock, bearing interest from 1 January 1863. Exchangeable on presentation with others making up a sufficient amount for bonds of £50, £100, £200, or £500. Signed by the Financial Agent of the Portuguese Government in London. Printed by Whiting, London.',
   'provisional certificate',
   'Conversion of External Debt of Portugal; Provisional Certificate; Deferred Stock; 3%; Three Pence; No. 4463; London; 1852; 1853; Portuguese government debt; Financial Agent; Whiting',
   'Portugal','United Kingdom',
   'Financial Agency of the Portuguese Government in London','ca. 1853','British pound sterling','English','1','19th century','','','','',''],

  [718,'goetzmann0718.jpg','/images/goetzmann0718.jpg',
   'Titulo Especial sem Juro do Fundo Externo Portuguez 3a Serie (Portuguese External Fund, 3rd Series, Non-Interest-Bearing Title), No. 303:486, Lisbon, 1902',
   'Special non-interest-bearing title (Titulo Especial sem Juro) of the Portuguese External Fund (Fundo Externo Portuguez), 3rd Series, No. 303:486. Denomination: Reis 308,000 = Libras 6-12-8 = Francos 166.67 = Marcos 135.34 = Florins 79.34. Issued under the Decree of 11 May 1902. Part of an emission of 477,517 obligations issued without interest, amortizable by lottery, in connection with the conversion of the 2nd Series of 477,517 obligations (3% bonds). Lisbon, Junta do Credito Publico, 22 December 1902. Signed by the Minister of Finance and the President of the Junta do Credito Publico. Black 1% stamp/seal visible. Text in Portuguese.',
   'bond',
   'Titulo Especial sem Juro; Fundo Externo Portuguez; Portuguese External Fund; 3rd Series; No. 303:486; Reis 308000; Lisbon; 1902; conversion; Junta do Credito Publico; non-interest-bearing',
   'Portugal','Portugal',
   'Junta do Credito Publico, Lisboa','1902-12-22','Portuguese real; British pound sterling; French franc; German mark; Dutch guilder','Portuguese','1','Early 20th century','','','','',''],

  [719,'goetzmann0719.jpg','/images/goetzmann0719.jpg',
   'Hope & Co. / Ketwich & Voombergh, Certificate of 6% Russian Funds in Bank Assignats (Certificaat van 6 per Cents Russische Fondsen), 1,000 Rubles, No. 1326, Amsterdam, 1825',
   'Bilingual Dutch/French certificate (Certificaat / Certificat d\'Inscription Russe) for an inscription of 1,000 Rubles in 6% Russian Funds payable in Bank Assignats (Russische Fondsen in Bank-Assignatien / Fonds Russes en Assignations de Banque), No. 1326. The inscription is registered at St. Petersburg in the Grand Book of the Public Debt of Russia (Grand-Livre) in the name of the Administration Office (Administratie-Kantoor) established at Amsterdam under the direction of Hope en Comp., Ketwich en Voombergh, and Wed. W. Borski. Interest received by the holder against surrender of coupons. Amsterdam, 11 May 1825. Notarially witnessed. Ten semi-annual coupons attached valid until 14 July 1829.',
   'certificate',
   'Hope & Co.; Ketwich & Voombergh; Wed. W. Borski; Russian Funds; 6% Bank Assignats; 1000 rubles; No. 1326; Amsterdam; 1825; certificaat; certificat inscription russe; Dutch; French',
   'Russia','Netherlands',
   'Hope en Comp. / Ketwich en Voombergh / Wed. W. Borski, Amsterdam','1825-05-11','Russian ruble (assignat); Dutch guilder','Dutch; French','1','19th century','','','','',''],

  [720,'goetzmann0720.jpg','/images/goetzmann0720.jpg',
   'USSR, State Loan for Reconstruction and Development of the National Economy, 1946, 25 Rubles, Series 012594, No. 33',
   'Bearer bond (Obligatsiya) of the USSR for the State Loan for Reconstruction and Development of the National Economy of the USSR (Gosudarstvennyy Zaym Vosstanovleniya i Razvitiya Narodnogo Khozyaystva SSSR), 1946. Denomination: 25 Rubles (Dvadtsat pyat rubley). Series 012594, No. 33, Category (Razryad) 98. Green decorative border with Soviet state emblem (hammer, sickle, wheat sheaves). Text in Russian.',
   'bond',
   'USSR; CCCP; Soviet Union; State Loan; Reconstruction; National Economy; 1946; 25 rubles; Series 012594; No. 33; Gosudarstvennyy Zaym; bearer bond',
   'Russia','Russia',
   'USSR / Narodniy Komissariat Finansov','1946','Soviet ruble','Russian','1','Mid 20th century','','','','',''],

  [721,'goetzmann0721.jpg','/images/goetzmann0721.jpg',
   'Russian Provisional Government, Internal 4.5% Prize-Winning Loan of 1917, 200 Rubles, Series 15800, No. 32, Second Category',
   'Bearer lottery bond (Bilet) of the Russian Provisional Government for the Internal 4.5% Prize-Winning Loan of 1917 (Gosudarstvennyy Vnutrenniy 4.5% Vygryshnyy Zaym 1917 goda). Issued based on resolutions of the Provisional Government of 11-13 August 1917. Denomination: 200 Rubles Nominal (V Dvesti Rubley Naritsatelnykh). Second category (Razryad vtoroy). Series 15800, No. 32. Brown/orange decorative design with allegorical figure. The last coupon is dated 16 January 1928. Interest at 4.5% per annum. Signed by the Director of the State Debt Department.',
   'bond',
   'Russian Provisional Government; 4.5% Prize-Winning Loan; 1917; 200 rubles; Series 15800; No. 32; Vygryshnyy Zaym; lottery bond; bearer bond; Russia; August 1917',
   'Russia','Russia',
   'Russian Provisional Government / Upravlenie Gosudarstvennym Dolgom','1917','Russian ruble','Russian','1','Early 20th century','','','','',''],

  [722,'goetzmann0722.jpg','/images/goetzmann0722.jpg',
   'City of Baku, 5% Loan of 1910, 189 Rubles / 504 Francs / £20 Bearer Bond, No. 20783',
   'Bearer bond (Obligatsiya / Bond) of the City of Baku for the 5% Loan of the City of Baku, 1910 (5% Zaym Goroda Baku 1910 goda). No. 20783. Denomination: 189 rubles = 504 francs = 20 pounds sterling. Total loan: 25,999,973 rubles = 2,887,140 francs sterling = 71,999,928 francs. Interest at 5% per annum. Bilingual Russian and English. Signed by the City of Baku authorities and the Baku City Bank. Ornate decorative border. The English side reads: Bond for One Hundred and Eighty-Nine Roubles, bearing FIVE per cent interest per annum, Bond to Bearer, 5% Loan of the City of Baku, 1910.',
   'bond',
   'City of Baku; Zaym Goroda Baku; 5% loan; 1910; 189 rubles; 504 francs; £20; No. 20783; bearer bond; bilingual Russian English; Azerbaijan; Baku City Bank',
   'Russia','Russia',
   'City of Baku / Bakinskiy Gorodskoy Bank','1910','Russian ruble; French franc; British pound sterling','Russian; English','1','Early 20th century','','','','',''],

  [723,'goetzmann0723.jpg','/images/goetzmann0723.jpg',
   'Imperial Russian Three Per Cent Loan 1859, £100 Sterling Inscription, No. 41150, Great Book of the Public Debt of Russia, London agents Thomson, Bonar & Co.',
   'Inscription certificate of the Imperial Russian Three Per Cent Loan 1859 for One Hundred Pounds Sterling (£100), No. 41150. Inscription recorded in the Great Book of the Public Debt of Russia, Book 9, Folio 1945. Annual interest of 3% paid half-yearly in London by J. Thomson, T. Bonar & Co.; in Berlin by F. Mart. Magnus in Thalers. Dividend warrants attached until 19 April 1869, after which a talon will issue new warrants. Conditions include provisions from the Imperial Ukase of 20 March 1859. Countersigned by Thomson, Bonar & Co. and F. Mart. Magnus, Contractors. Danish fiscal revenue stamps affixed.',
   'bond inscription certificate',
   'Imperial Russian Three Per Cent Loan; 1859; £100; No. 41150; Great Book of Public Debt; Thomson Bonar & Co.; F. Mart. Magnus; London; Berlin; 3%; annual interest; Danish revenue stamps',
   'Russia','United Kingdom',
   'Imperial Russian Government / Thomson, Bonar & Co., London','1859','British pound sterling; German thaler','English','1','19th century','','','','',''],

  [724,'goetzmann0724.jpg','/images/goetzmann0724.jpg',
   'Citta di Bari delle Puglie, Prestito a Premi (Prize Loan), 100 Lire Bond Redeemable at 150 Lire, Serie 317, No. 049, Bari, 1869',
   'Prize loan bond (Obbligazione di Lire 100 al Portatore Rimborsabile con Lire 150) of the Citta di Bari delle Puglie (City of Bari in Puglia), issued by the Compagnia Assicuratrice Milano. Serie 317, No. 049. Deliberated by ordinanza of 31 December 1867, approved by Royal Decree of 11 June 1868. Non-interest-bearing; prizes and repayment awarded by lottery. Lottery extraction plan (Piano delle Estrazioni) printed on bond face, showing three categories (Casella A, B, C) over 20 years. Total emission: Lire 5,000,000. Bari delle Puglie, 11 May 1869. Signed by Il Sindaco. Right side: blank prize annotation sheet (Annotazione dei Premi). Green decorative border.',
   'bond',
   'Citta di Bari; Bari delle Puglie; Prestito a Premi; prize loan; 100 lire; 150 lire; Serie 317; No. 049; lottery; Compagnia Assicuratrice Milano; 1869; Italy; municipal bond',
   'Italy','Italy',
   'Citta di Bari delle Puglie / Compagnia Assicuratrice Milano','1869-05-11','Italian lira','Italian','1','19th century','','','','',''],

  [725,'goetzmann0725.jpg','/images/goetzmann0725.jpg',
   'Caisse Autonome des Monopoles du Royaume de Roumanie, 7% External Gold Bond, Emprunt de Stabilisation et de Developpement 1929, No. 161,517, Frs. 2,352.90 / US $100',
   'External gold bearer bond (Obligation Exterieure or 7%) of the Caisse Autonome des Monopoles du Royaume de Roumanie (Autonomous Monopolies Fund of the Kingdom of Romania / Casa Autonoma a Monopolurilor Regatului Romaniei), No. 161,517. From the Emprunt de Stabilisation et de Developpement de 1929 (Stabilization and Development Loan of 1929). Denomination: Frs.E 2,352.90 or US $100. Dated 1 February 1929. Interest at 7% per annum. Coupon sheet (coupons 31-60) attached on right, each bearing portrait of Romanian royalty. Signed by the Director and Governor of the Caisse Autonome. Blue decorative border with Romanian royal coat of arms.',
   'bond',
   'Caisse Autonome des Monopoles; Royaume de Roumanie; Romania; 7% external gold bond; Emprunt de Stabilisation; 1929; No. 161517; Frs. 2352.90; US $100; bearer bond; coupon sheet',
   'Romania','Romania',
   'Caisse Autonome des Monopoles du Royaume de Roumanie','1929-02-01','French franc; United States dollar','French; Romanian','1','Early 20th century','','','','',''],

  [726,'goetzmann0726.jpg','/images/goetzmann0726.jpg',
   'Caisse Autonome des Monopoles du Royaume de Roumanie, 7% External Gold Bond 1929 (verso) - Coupon Sheet Continuation and Bond Back Labels',
   'Reverse side (verso) of the Caisse Autonome des Monopoles du Royaume de Roumanie 7% External Gold Bond 1929 (see goetzmann0725 for recto). Left portion: continuation coupon sheet. Right portion shows two documents: the face of the bond in lighter blue with bird vignette and the back of the bond in tan/yellow showing the bond identification and Romanian coat of arms. Denomination Frs. 2,352.90 or US $100 confirmed on back labels.',
   'bond',
   'Caisse Autonome des Monopoles; Royaume de Roumanie; Romania; 7% external gold bond; 1929; verso; coupon sheet; back label; Frs. 2352.90; US $100',
   'Romania','Romania',
   'Caisse Autonome des Monopoles du Royaume de Roumanie','1929-02-01','French franc; United States dollar','French; Romanian','1','Early 20th century','','','','',''],

  [727,'goetzmann0727.jpg','/images/goetzmann0727.jpg',
   'Bons Representatifs des Annuites Arrieres de la Dette Publique Ottomane Serie A 1928, Quadrilingual Bearer Bond, Ltqs 22 / £20 / Fr. 500, No. 175,406',
   'Quadrilingual bearer bond (Titre de UN Bon au Porteur / Obligation of ONE Bond to Bearer / Gutschein uber EINEN Gutschein an den Inhaber / Hamile Muharrer BIR Senetttir) of the Conseil de la Dette Publique Repartie de l\'Ancien Empire Ottoman, representing Arrear Annuities of the Ottoman Public Debt "Serie A" 1928. No. 175,406. Denomination: Ltqs 22 (Turkish pounds) or £20 or Francs 500. Text in French, English, German, and Ottoman Turkish (Arabic script). Coupon sheet attached on right. Issued by the Conseil de la Dette Publique, signed by its President. The bond is guaranteed by the contributing States of the former Ottoman Empire.',
   'bond',
   'Bons Representatifs; Annuites Arrieres; Dette Publique Ottomane; Serie A 1928; Ltqs 22; £20; Fr. 500; No. 175406; quadrilingual; French; English; German; Ottoman Turkish; bearer bond',
   'Turkey','France',
   'Conseil de la Dette Publique Repartie de l\'Ancien Empire Ottoman','1928','Turkish lira; British pound sterling; French franc','French; English; German; Ottoman Turkish','1','Early 20th century','','','','',''],

  [728,'goetzmann0728.jpg','/images/goetzmann0728.jpg',
   'Bons Representatifs des Annuites Arrieres de la Dette Publique Ottomane Serie A 1928, Bond No. 175,406 (verso) - Contract Extracts in Four Languages and Coupon Sheet',
   'Reverse side (verso) of the Ottoman Public Debt Serie A 1928 bond No. 175,406 (see goetzmann0727 for recto). Left portion: continuation coupon sheet with Ottoman Turkish coupons. Right portion: large text area printing "Extraits du Contrat" (French), "Extracts from the Contract" (English), "Auszuge aus dem Vertrage" (German), and Ottoman Turkish equivalent in four columns, detailing the terms of the contract of 13 June 1928 between the Turkish Republic and Ottoman debt holders. Amortization/repayment schedule tables at bottom. Multiple red Conseil de la Dette stamps affixed.',
   'bond',
   'Bons Representatifs; Annuites Arrieres; Dette Publique Ottomane; Serie A 1928; verso; contract extracts; amortization schedule; quadrilingual; French; English; German; Ottoman Turkish',
   'Turkey','France',
   'Conseil de la Dette Publique Repartie de l\'Ancien Empire Ottoman','1928','Turkish lira; British pound sterling; French franc','French; English; German; Ottoman Turkish','1','Early 20th century','','','','',''],

  [729,'goetzmann0729.jpg','/images/goetzmann0729.jpg',
   'Principality of Bulgaria, 4.5% Gold Loan of 1907, 500 Francs Bearer Bond, No. 279,186, Sofia, 1907',
   'Bearer gold bond (Obligation au Porteur / Bond to Bearer / Obligation auf den Inhaber) of the Principality of Bulgaria (Knyazhestvo Balgariya) for the Bulgarian 4.5% Gold Loan of 1907. No. 279,186. Denomination: 500 Francs (Five Hundred Francs / Funf Hundert Francs). Total: 290,000 bonds. Total loan: 145,000,000 gold francs = 163,700,000 marks = £5,800,000 = 147,753,600 Austrian crowns = 140 million Bulgarian leva. Quadrilingual: Bulgarian, French, German, and English. Sofia, 14 September 1907. Signed by the Bulgarian Minister of Finance and other authorized officials. Decorative red and black border.',
   'bond',
   'Principality of Bulgaria; Knyazhestvo Balgariya; 4.5% gold loan; 1907; 500 francs; No. 279186; Sofia; bearer bond; quadrilingual; Bulgarian; French; German; English',
   'Bulgaria','Bulgaria',
   'Principality of Bulgaria / Ministerstvo na Finansite','1907-09-14','French franc; German mark; British pound sterling; Austrian crown; Bulgarian lev','Bulgarian; French; German; English','1','Early 20th century','','','','',''],

  [730,'goetzmann0730.jpg','/images/goetzmann0730.jpg',
   'Principality of Bulgaria, 4.5% Gold Loan of 1907 (verso) - Conditions of Loan and Table of Amortisation in Four Languages',
   'Reverse side (verso) of the Principality of Bulgaria 4.5% Gold Loan of 1907 bond (see goetzmann0729 for recto). Left portion: conditions of the loan in Bulgarian (Usloviya na Zayma), German (Bedingungen der Anleihe), French (Conditions de l\'Emprunt), and English (Conditions of the Loan). Right portion: amortization table (Tablitsa za Porishenie / Tilgungs-Plan / Tableau d\'Amortissement / Table of Amortisation) in all four languages, with lottery draw dates and bond numbers.',
   'bond',
   'Principality of Bulgaria; 4.5% gold loan; 1907; verso; conditions of loan; amortization table; Tilgungs-Plan; quadrilingual; Bulgarian; French; German; English',
   'Bulgaria','Bulgaria',
   'Principality of Bulgaria / Ministerstvo na Finansite','1907-09-14','French franc; German mark; British pound sterling; Austrian crown; Bulgarian lev','Bulgarian; French; German; English','1','Early 20th century','','','','',''],

  [731,'goetzmann0731.jpg','/images/goetzmann0731.jpg',
   'Ottoman Empire, Unified Converted Debt / Dette Convertie Unifiee de l\'Empire Ottoman, 4% Bearer Bond, Fr. 500 / £20, No. 4,176,790, Constantinople, 1903',
   'Bearer bond (Titre au Porteur de Fr. 500 / Bond to Bearer for £20) of the Unified Converted Debt of the Ottoman Empire (Dette Convertie Unifiee de l\'Empire Ottoman / Osmanli Imparatorlugu Umumiyesi Borcu). Total capital: Fr. 744,068,000 = £29,762,520. Issue of 1,488,126 obligations of Fr. 500 / £20 each. No. 4,176,790. Annual interest: 20 Francs (£0.16.0), payable on 1/14 March and 1/14 September. Constantinople, 1/14 September 1903. Trilingual: French, English, and Ottoman Turkish (Arabic script). Issued by the Banque Imperiale Ottomane. Coupon/talon sheet attached on right. Red/salmon and green decorative border with Ottoman tughra.',
   'bond',
   'Ottoman Empire; Unified Converted Debt; Dette Convertie Unifiee; 4%; Fr. 500; £20; No. 4176790; Constantinople; 1903; Banque Imperiale Ottomane; trilingual; French; English; Ottoman Turkish; tughra',
   'Turkey','Turkey',
   'Banque Imperiale Ottomane / Administration de la Dette Publique Ottomane','1903-09-14','French franc; British pound sterling; Ottoman lira','French; English; Ottoman Turkish','1','Early 20th century','','','','',''],

  [732,'goetzmann0732.jpg','/images/goetzmann0732.jpg',
   'Ottoman Empire, Unified Converted Debt 4%, Bond No. 4,176,790 (verso) - Amortization Table and Bond Back Labels in Three Languages',
   'Reverse side (verso) of the Ottoman Unified Converted Debt 4% bond No. 4,176,790 (see goetzmann0731 for recto). Left: continuation coupon sheet in Ottoman Turkish. Right: large area showing the Tableau d\'Amortissement of the 1,488,126 Obligations 4% de la Dette Convertie Unifiee de l\'Empire Ottoman (amortization table), plus terms in French and English. Bottom portion shows the back labels in three languages: French (Dette Convertie Unifiee de l\'Empire Ottoman, 1 Obligation de Fr.500, No. 4,176,790, Titre de Fr.500), English (Unified Converted Debt of the Ottoman Empire, 1 Bond of £20, No. 4,176,790, Certificate for £20), and Ottoman Turkish equivalent.',
   'bond',
   'Ottoman Empire; Unified Converted Debt; 4%; verso; amortization table; Tableau d\'Amortissement; Fr. 500; £20; No. 4176790; trilingual; French; English; Ottoman Turkish; back labels',
   'Turkey','Turkey',
   'Banque Imperiale Ottomane / Administration de la Dette Publique Ottomane','1903-09-14','French franc; British pound sterling; Ottoman lira','French; English; Ottoman Turkish','1','Early 20th century','','','','',''],

  // Row 733: image file missing from archive - skipped

  [734,'goetzmann0734.jpg','/images/goetzmann0734.jpg',
   'Principality of Bulgaria, 5% Gold Loan of 1902, 500 Francs Bearer Bond, No. 130,763, Sofia, 1902',
   'Bearer gold bond (Obligation au Porteur / Bond to Bearer / Obligation auf den Inhaber / Obligatsiya) of the Principality of Bulgaria (Knyazhestvo Balgariya) for the Bulgarian Government 5% Gold Loan of 1902 (Bulgarski 5% Darzhaven Zaem v Zlato ot 1902 godini). No. 130,763. Denomination: Five Hundred Francs / Cinq Cents Francs / Funf Hundert Francs. Total: 212,000 bonds. Total loan approximately 56,580,000 francs. Quadrilingual: Bulgarian, French, German, and English. Sofia, 1902. Signed by the Bulgarian Finance Minister and authorized officials. Blue and black decorative border with Bulgarian royal coat of arms.',
   'bond',
   'Principality of Bulgaria; Knyazhestvo Balgariya; 5% gold loan; 1902; 500 francs; No. 130763; Sofia; bearer bond; quadrilingual; Bulgarian; French; German; English',
   'Bulgaria','Bulgaria',
   'Principality of Bulgaria / Ministerstvo na Finansite','1902','French franc; German mark; British pound sterling; Austrian crown; Bulgarian lev','Bulgarian; French; German; English','1','Early 20th century','','','','','']
];

rows.forEach(row => {
  const rowIndex = row[0];
  row.forEach((val, colIndex) => {
    const cellAddr = XLSX.utils.encode_cell({r: rowIndex, c: colIndex});
    ws[cellAddr] = {v: val, t: typeof val === 'number' ? 'n' : 's'};
  });
});

XLSX.writeFile(wb, 'oov_data_new.xlsx');
console.log('Done - rows 715-734 written (733 skipped - image missing).');
