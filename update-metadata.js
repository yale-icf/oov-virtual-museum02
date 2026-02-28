const XLSX = require('xlsx');
const path = require('path');

const SPREADSHEET = path.join(__dirname, 'financial_documents_template.xlsx');

// Complete metadata for all 109 empty documents
// type uses \u001d as separator for multiple values
const SEP = '\u001d';

const metadata = {
  'goetzmann0630.jpg': {
    title: 'Certificate of 6% Russian Bonds in Bank Assignations, 1827',
    description: 'Dutch/French bilingual certificate for 6% Russian government bonds in bank assignations. Denomination of 1,000 Roubles. Issued in Amsterdam, September 1827, through Hope and Company.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Dutch, French',
    issueDate: '1827'
  },
  'goetzmann0631.jpg': {
    title: 'Certificate of Russian Government Bonds, 1,000 Roubles, 1835',
    description: 'Certificate for Russian government bonds, denomination of 1,000 Roubles in assignations at 6% interest. Issued 1835 through Stadnitski & van Heukelom consortium.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Dutch, French',
    issueDate: '1835'
  },
  'goetzmann0632.jpg': {
    title: 'England 3½% War Loan 1932, £100 Certificate (Specimen)',
    description: 'Specimen certificate for the British 3½% War Loan of 1932, denomination of £100. Administered through Amsterdam.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United Kingdom',
    issuingCountry: 'Netherlands',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English, Dutch',
    issueDate: '1932'
  },
  'goetzmann0633.jpg': {
    title: 'Austrian Housing Lottery Bond (Bau-Los), 1,200 Kronen, 1921',
    description: 'Austrian housing lottery bond (Bau-Los) with a denomination of 1,200 Kronen, issued in 1921.',
    type: ['Bond', 'Debt', 'Lottery'].join(SEP),
    subjectCountry: 'Austria',
    Period: '20th Century',
    currency: 'Kronen',
    language: 'German',
    issueDate: '1921'
  },
  'goetzmann0634.jpg': {
    title: '5% Loan of the City of Baku, 189 Roubles, 1910',
    description: 'Municipal bond for the 5% Loan of the City of Baku, denomination of 189 Roubles. Multilingual text in Russian, English, and French. Issued during the Russian Empire era.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Azerbaijan',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian, English, French',
    issueDate: '1910'
  },
  'goetzmann0635.jpg': {
    title: 'Republic of Bolivia Internal Loan, 100 Pesos, 1827',
    description: 'Bolivian government internal loan bond for 100 Pesos, issued under the 1826 law.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bolivia',
    Period: '19th Century',
    currency: 'Pesos',
    language: 'Spanish',
    issueDate: '1827'
  },
  'goetzmann0636.jpg': {
    title: 'Bolivian Government Loan Certificate, 1872',
    description: 'Certificate for the Bolivian Government Loan of 1872.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bolivia',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1872'
  },
  'goetzmann0637.jpg': {
    title: 'Republic of Bolivia Public Fund Bond, 1844',
    description: 'Bolivian public fund bond issued in 1844.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bolivia',
    Period: '19th Century',
    currency: 'Pesos',
    language: 'Spanish',
    issueDate: '1844'
  },
  'goetzmann0638.jpg': {
    title: '8% Public Works Loan Bond, Kingdom of Serbs, Croats and Slovenes',
    description: 'Bond for the 8% Public Works Loan of the Kingdom of Serbs, Croats and Slovenes (Yugoslavia). Multilingual text in Serbian, German, Hungarian, French, and English.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Yugoslavia',
    Period: '20th Century',
    language: 'Serbian, German, Hungarian, French, English'
  },
  'goetzmann0639.jpg': {
    title: '4½% Government Bond, 1,000 Kronen, Austria-Hungary',
    description: 'Austro-Hungarian government bond at 4½% interest, denomination of 1,000 Kronen.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Austria',
    Period: '20th Century',
    currency: 'Kronen',
    language: 'German'
  },
  'goetzmann0640.jpg': {
    title: 'Socialist Republic of Serbia Bond, 100 Dinara',
    description: 'Bond of the Socialist Republic of Serbia, denomination of 100 Dinara. Features coupons. Yugoslav era.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Serbia',
    Period: '20th Century',
    currency: 'Dinara',
    language: 'Serbian'
  },
  'goetzmann0641.jpg': {
    title: 'Socialist Republic of Bosnia and Herzegovina Employment Loan Bond, 10,000 Dinara',
    description: 'Employment loan bond of the Socialist Republic of Bosnia and Herzegovina, denomination of 10,000 Dinara. Features photograph of workers. Yugoslav era.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bosnia and Herzegovina',
    Period: '20th Century',
    currency: 'Dinara',
    language: 'Serbian, Bosnian'
  },
  'goetzmann0642.jpg': {
    title: 'Prospectus for Conversion of Imperial Brazilian 5% Loans, 1865-1886',
    description: 'Prospectus for the conversion and redemption of Imperial Brazilian loans at 5%, covering the period 1865-1886. Issue of £20,000,000 at 4%.',
    type: ['Prospectus', 'Debt'].join(SEP),
    subjectCountry: 'Brazil',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English'
  },
  'goetzmann0643.jpg': {
    title: 'Hungarian Crown Lands 4½% Annuity Loan Bond, 1913',
    description: 'Bond for the Hungarian Crown Lands 4½% Annuity Loan, issued in 1913. Denominated in Korona, German Imperial Marks, Francs, and Sterling.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Hungary',
    Period: '20th Century',
    currency: 'Korona',
    language: 'Hungarian, German, French, English',
    issueDate: '1913'
  },
  'goetzmann0644.jpg': {
    title: 'City of Budapest 4% Municipal Loan Bond, 1911',
    description: 'Municipal bond for the City of Budapest at 4% interest, issued in 1911. Features panoramic illustration of Budapest and the Danube River. Denominations in Korona and Francs.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Hungary',
    Period: '20th Century',
    currency: 'Korona',
    language: 'Hungarian, French',
    issueDate: '1911'
  },
  'goetzmann0645.jpg': {
    title: 'Province of Buenos Aires 5% Consolidation Gold Loan, £20, 1915',
    description: 'Bond for the Province of Buenos Aires 5% Consolidation Gold Loan, denomination of £20 or 504 Francs, issued in 1915.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Argentina',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English, French',
    issueDate: '1915'
  },
  'goetzmann0646.jpg': {
    title: 'United Incandescent Lamp and Electrical Company Share Certificate, 1,000 Pengő',
    description: 'Share certificate of the United Incandescent Lamp and Electrical Company, denomination of 1,000 Pengő. Hungarian company, 1930s-1940s era.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Hungary',
    Period: '20th Century',
    currency: 'Pengő',
    language: 'Hungarian'
  },
  'goetzmann0647.jpg': {
    title: 'Companhia da Zambezia Provisional Share Certificate, 1900',
    description: 'Provisional share certificate of the Companhia da Zambezia, issued in 1900. Capital of 2.7 billion Reis / 15 million Francs / £600,000. Issued in Lisbon. Mozambican colonial-era company.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Mozambique',
    issuingCountry: 'Portugal',
    Period: '20th Century',
    currency: 'Reis',
    language: 'Portuguese',
    issueDate: '1900'
  },
  'goetzmann0648.jpg': {
    title: 'Principality of Bulgaria State Loan Bond, 500 Francs, 1892',
    description: 'State loan bond of the Principality of Bulgaria, denomination of 500 Francs, issued in 1892.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '19th Century',
    currency: 'Francs',
    language: 'Bulgarian, French',
    issueDate: '1892'
  },
  'goetzmann0650.jpg': {
    title: 'Chrysler Corporation Share Certificate, 10 Shares (Specimen)',
    description: 'Specimen share certificate of Chrysler Corporation for 10 shares. Administered through the Netherlands.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'United States',
    issuingCountry: 'Netherlands',
    Period: '20th Century',
    currency: 'USD',
    language: 'English'
  },
  'goetzmann0651.jpg': {
    title: 'Republic of Colombia 4% Funding Certificate, 1934',
    description: 'Funding certificate of the Republic of Colombia at 4% interest, issued in 1934.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Colombia',
    Period: '20th Century',
    currency: 'USD',
    language: 'English',
    issueDate: '1934'
  },
  'goetzmann0652.jpg': {
    title: 'Republic of Cuba 4½% Gold Bond (Specimen), 1949',
    description: 'Specimen gold bond of the Republic of Cuba at 4½% interest, issued in 1949.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Cuba',
    Period: '20th Century',
    currency: 'USD',
    language: 'English',
    issueDate: '1949'
  },
  'goetzmann0653.jpg': {
    title: 'Czechoslovak State Bond, $100',
    description: 'State bond of Czechoslovakia, denomination of $100.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Czechoslovakia',
    Period: '20th Century',
    currency: 'USD',
    language: 'English'
  },
  'goetzmann0654.jpg': {
    title: 'Egyptian Credit Foncier Share Certificate, £20 / 500 Francs, 1904',
    description: 'Share certificate of the Egyptian Credit Foncier, denomination of £20 or 500 Francs, issued in 1904.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Egypt',
    Period: '20th Century',
    currency: 'GBP',
    language: 'French, English',
    issueDate: '1904'
  },
  'goetzmann0655.jpg': {
    title: 'Consolidated £3 Per Cent Annuities Transfer Receipt, 1860',
    description: 'Transfer receipt for British Consolidated £3 per cent annuities (Consols), dated 1860.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United Kingdom',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1860'
  },
  'goetzmann0656.jpg': {
    title: 'Consolidated £3 Per Cent Annuities Acceptance and Dividends, 1831',
    description: 'Acceptance and dividend document for British Consolidated £3 per cent annuities (Consols), dated 1831.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United Kingdom',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1831'
  },
  'goetzmann0657.jpg': {
    title: 'Land Bank of Estonia (Eesti Maapank) 4% Mortgage Bond, 250 Krooni, 1927',
    description: 'Mortgage bond of the Land Bank of Estonia (Eesti Maapank) at 4% interest, denomination of 250 Krooni, issued in 1927.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Estonia',
    Period: '20th Century',
    currency: 'Krooni',
    language: 'Estonian',
    issueDate: '1927'
  },
  'goetzmann0658.jpg': {
    title: 'Kingdom of Serbia 5% Gold Loan Bond, 500 Francs, 1913',
    description: 'Gold loan bond of the Kingdom of Serbia at 5% interest, denomination of 500 Francs, issued in 1913.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Serbia',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Serbian, French',
    issueDate: '1913'
  },
  'goetzmann0659.jpg': {
    title: 'City of Berlin Municipal Bond (Schuldverschreibung), 50,000 Mark, 1923',
    description: 'Municipal bond (Schuldverschreibung) of the City of Berlin, denomination of 50,000 Mark, issued during the hyperinflation period of 1923.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Germany',
    Period: '20th Century',
    currency: 'Mark',
    language: 'German',
    issueDate: '1923'
  },
  'goetzmann0660.jpg': {
    title: 'Kingdom of Greece 2½% Gold Loan Bond, 2,500 Drachmai, 1898',
    description: 'Gold loan bond of the Kingdom of Greece at 2½% interest, denomination of 2,500 Drachmai, issued in 1898.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Greece',
    Period: '19th Century',
    currency: 'Drachmai',
    language: 'Greek, French',
    issueDate: '1898'
  },
  'goetzmann0661.jpg': {
    title: 'The Native Guano Company Limited Share Certificate, £5, 1881',
    description: 'Share certificate of The Native Guano Company Limited, denomination of £5, dated November 15, 1881. Adelaide (Australia) registered.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'United Kingdom',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1881'
  },
  'goetzmann0662.jpg': {
    title: 'Hungarian Mortgage Credit Bank Prize Bond, 100 Korona',
    description: 'Prize bond of the Hungarian Mortgage Credit Bank (Magyar Jelzálog-Hitelbank), denomination of 100 Korona. Multilingual text in Hungarian, French, and German.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Hungary',
    Period: '20th Century',
    currency: 'Korona',
    language: 'Hungarian, French, German'
  },
  'goetzmann0663.jpg': {
    title: 'India £3 Per Cent Stock Transfer Receipt, 1864',
    description: 'Transfer receipt for India £3 per cent stock, dated August 17, 1864. Transfer from Fred Binmer Tindell to Catherine Mary Tindell through the Bank of England.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'India',
    issuingCountry: 'United Kingdom',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1864'
  },
  'goetzmann0664.jpg': {
    title: 'India £3 Per Cent Stock Transfer Receipt, 1885',
    description: 'Transfer receipt for India £3 per cent stock, dated June 1885. Also involving Fred Binmer Tindell.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'India',
    issuingCountry: 'United Kingdom',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1885'
  },
  'goetzmann0665.jpg': {
    title: 'Principality of Bulgaria 6% State Loan Bond, 500 Francs, 1892',
    description: 'State loan bond of the Principality of Bulgaria at 6% interest, denomination of 500 Francs, issued in 1892. Stamped "RECOUVRE" (recovered/collected).',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '19th Century',
    currency: 'Francs',
    language: 'Bulgarian, French',
    issueDate: '1892'
  },
  'goetzmann0666.jpg': {
    title: 'Principality of Bulgaria 6% State Loan, Two Obligations, 500 Francs, 1892',
    description: 'Two obligations of the Principality of Bulgaria 6% State Loan, denomination of 500 Francs each, issued in 1892.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '19th Century',
    currency: 'Francs',
    language: 'Bulgarian, French',
    issueDate: '1892'
  },
  'goetzmann0667.jpg': {
    title: 'Imperial Russian Government 3⅜% Conversion Bond Coupon Sheet, 150 Roubles',
    description: 'Coupon sheet for the Imperial Russian Government 3⅜% Conversion Bond, denomination of 150 Roubles.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French'
  },
  'goetzmann0668.jpg': {
    title: 'Imperial Russian Government 3⅞% Conversion Bond, 150 Roubles, 1899',
    description: 'Conversion bond of the Imperial Russian Government at 3⅞% interest, denomination of 150 Roubles, dated March 1899.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French',
    issueDate: '1899'
  },
  'goetzmann0669.jpg': {
    title: 'Imperial Russian Government 4% State Loan Bond, 1902',
    description: 'State loan bond of the Imperial Russian Government at 4% interest, denomination of 1,000 Imperial German Marks with multiple currency equivalents, issued in 1902.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Marks',
    language: 'Russian',
    issueDate: '1902'
  },
  'goetzmann0670.jpg': {
    title: 'Imperial Russian Government 4% Gold Loan Bond, 125 Gold Roubles, 1894',
    description: 'Gold loan bond of the Imperial Russian Government at 4% interest, denomination of 125 Gold Roubles, issued in 1894.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1894'
  },
  'goetzmann0671.jpg': {
    title: 'Imperial Russian Government 4% Gold Loan Bond, 125 Gold Roubles, 1894',
    description: 'Gold loan bond of the Imperial Russian Government at 4% interest, denomination of 125 Gold Roubles / 500 Francs, issued in 1894. Different serial number from goetzmann0670.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1894'
  },
  'goetzmann0672.jpg': {
    title: 'Imperial Russian Government Nikolaev Railway Bond, 125 Roubles, 1867',
    description: 'Bond for the Imperial Russian Government Nikolaev Railway, denomination of 125 Roubles / 500 Francs / £20 Sterling. Part of an issue of 600,000 obligations, dated 1867.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French',
    issueDate: '1867'
  },
  'goetzmann0673.jpg': {
    title: 'Imperial Russian State Nobles\' Land Bank 3½% Mortgage Bond, 150 Roubles',
    description: 'Mortgage bond (Zakladnoi List) of the Imperial Russian State Nobles\' Land Bank at 3½% interest, denomination of 150 Roubles.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian'
  },
  'goetzmann0674.jpg': {
    title: 'Moscow-Kiev-Voronezh Railway 4½% Bond, 187.50 Roubles, 1914',
    description: 'Bond for the Moscow-Kiev-Voronezh Railway Company at 4½% interest, denomination of 187 Roubles 50 Kopecks / 500 Francs, issued in 1914.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1914'
  },
  'goetzmann0675.jpg': {
    title: 'Imperial Russian Government Nikolaev Railway Bond, Five Obligations, 625 Roubles, 1867',
    description: 'Five obligations for the Imperial Russian Government Nikolaev Railway, total denomination of 625 Roubles / 2,500 Francs / £100 Sterling / 4,150 Gold Guilders, dated 1867.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French',
    issueDate: '1867'
  },
  'goetzmann0676.jpg': {
    title: 'Imperial Russian Government Consolidated 4% Railway Bond, 2nd Series, 125 Gold Roubles',
    description: 'Consolidated railway bond of the Imperial Russian Government at 4% interest, 2nd Series, denomination of 125 Gold Roubles.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian'
  },
  'goetzmann0677.jpg': {
    title: 'Russian Freedom Loan (Zaem Svobody), 5% Bond, 100 Roubles, 1917',
    description: 'Freedom Loan bond issued by the Russian Provisional Government after the February Revolution, at 5% interest, denomination of 100 Roubles, dated 1917. Features the Tauride Palace.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1917'
  },
  'goetzmann0678.jpg': {
    title: 'Imperial Russian Government Nikolaev Railway Bond, 125 Roubles, 1869',
    description: 'Bond for the Imperial Russian Government Nikolaev Railway, denomination of 125 Roubles with multiple currency equivalents, dated 1869.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French',
    issueDate: '1869'
  },
  'goetzmann0679.jpg': {
    title: 'Student Loan Marketing Association (Sallie Mae) Yield Curve Note, Due 1991',
    description: 'Registered Yield Curve Note of the Student Loan Marketing Association (Sallie Mae), due 1991. Number R 2217. Dated March 6, 1986. Held by Hare & Co.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United States',
    Period: '20th Century',
    currency: 'USD',
    language: 'English',
    issueDate: '1986'
  },
  'goetzmann0688.jpg': {
    title: 'Student Loan Marketing Association (Sallie Mae) Yield Curve Note, 1986',
    description: 'Yield Curve Note of the Student Loan Marketing Association (Sallie Mae), issued by Kidder, Peabody & Co. in 1986. Modern American financial instrument.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United States',
    Period: '20th Century',
    currency: 'USD',
    language: 'English',
    issueDate: '1986'
  },
  'goetzmann0689.jpg': {
    title: 'Bond Certificate with Coupon Sheet (Verso)',
    description: 'Reverse side of a bond certificate showing coupon sheet with $40 denominations. Includes handwritten details and revenue stamps.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'United States',
    Period: '20th Century',
    currency: 'USD',
    language: 'English'
  },
  'goetzmann0690.jpg': {
    title: 'Republic of Peru Guaranteed Loan, National Pisco to Yca Railway Company, £100',
    description: 'Guaranteed loan bond of the Republic of Peru for the National Pisco to Yca Railway Company, denomination of £100.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Peru',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English'
  },
  'goetzmann0691.jpg': {
    title: 'Chilean Eastern Central Railway Company First Mortgage Bond, £20',
    description: 'First mortgage bond of the Chilean Eastern Central Railway Company (Compagnie du Chemin de Fer de l\'Est Central), denomination of £20. Text in English and French.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Chile',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, French'
  },
  'goetzmann0692.jpg': {
    title: 'Chilean Eastern Central Railway Company First Mortgage Bond, £20 (Verso)',
    description: 'Reverse side of the Chilean Eastern Central Railway Company first mortgage bond.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Chile',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, French'
  },
  'goetzmann0693.jpg': {
    title: 'Illustrated European Bond or Share Certificate',
    description: 'Decorative illustrated certificate with artistic elements, figures, and landscape scenes. Possibly Italian or French, late 19th or early 20th century. Maritime imagery, possibly related to a canal venture.',
    type: 'Security',
    Period: '19th Century',
    language: 'French'
  },
  'goetzmann0694.jpg': {
    title: 'European Bond Certificate (Verso with Amortization Tables)',
    description: 'Reverse side of a decorative European bond certificate with amortization tables and coupons.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    Period: '19th Century'
  },
  'goetzmann0695.jpg': {
    title: 'Bond Certificate (Verso with French Amortization Table)',
    description: 'Reverse side of a bond certificate showing Tableau d\'Amortissement (amortization table) and Tilgungs-Plan in French and German.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    Period: '19th Century',
    language: 'French, German'
  },
  'goetzmann0696.jpg': {
    title: 'Kingdom of Serbia Bond, 500 Francs / 500 Dinara',
    description: 'Bond of the Kingdom of Serbia, denomination of 500 Francs / 500 Dinara. Text in Serbian Cyrillic and German.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Serbia',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Serbian, German'
  },
  'goetzmann0697.jpg': {
    title: 'Chinese Imperial Government 4½% Gold Loan Bond, £100 Sterling, 1898',
    description: 'Gold loan bond of the Chinese Imperial Government at 4½% interest, denomination of £100 Sterling, issued in 1898. Cancelled with hole punch. Text in English and German.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, German',
    issueDate: '1898'
  },
  'goetzmann0698.jpg': {
    title: 'Chinese Imperial Government 4½% Gold Loan Bond, £100, 1898 (Interior Page)',
    description: 'Interior page of the Chinese Imperial Government 4½% Gold Loan Bond of 1898.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, German',
    issueDate: '1898'
  },
  'goetzmann0699.jpg': {
    title: 'Chinese Imperial Government 4½% Gold Loan Coupon Sheet, 1898',
    description: 'Coupon sheet for the Chinese Imperial Government 4½% Gold Loan Bond of 1898.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, German',
    issueDate: '1898'
  },
  'goetzmann0701.jpg': {
    title: 'Republic of Honduras Government Bond, 500 Francs',
    description: 'Government bond of the Republic of Honduras, denomination of 500 Francs.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Honduras',
    Period: '19th Century',
    currency: 'Francs',
    language: 'French'
  },
  'goetzmann0702.jpg': {
    title: 'Honduras Government Loan, £100, 1869',
    description: 'Government loan bond of Honduras, denomination of £100, issued in 1869.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Honduras',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1869'
  },
  // === Batch 8: 0718, 0738, 0966-0971, 0974-0975 ===
  'goetzmann0718.jpg': {
    title: 'Portuguese External Fund Special Bond (Título Especial Sem Juro), 3rd Series',
    description: 'Portuguese government bond from the External Fund (Fundo Externo Portuguez), 3rd Series. Title of 1 Obligation, denomination of 308,000 Reis. Features ornate border and Portuguese coat of arms. Handwritten date July 1928.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Portugal',
    Period: '20th Century',
    currency: 'Reis',
    language: 'Portuguese',
    issueDate: '1928'
  },
  'goetzmann0733.jpg': {
    title: '[No thumbnail available]',
    description: 'Thumbnail image not available for analysis.',
    type: '',
    subjectCountry: '',
    Period: ''
  },
  'goetzmann0738.jpg': {
    title: 'Russian Bond Reverse — Amortization Tables (Tilgungs-Plan)',
    description: 'Reverse side of a Russian or Eastern European bond showing amortization tables (Tilgungs-Plan / Tableau d\'Amortissement) in Russian, German, and French. Features credit institution payment schedules.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    language: 'Russian, German, French'
  },
  'goetzmann0966.jpg': {
    title: 'Bulgarian Government 5% Gold Loan of 1902, 500 Francs (Page 1)',
    description: 'Bond from the Principality of Bulgaria for the 5% State Gold Loan of 1902. Obligation of 500 Francs. Bond number 008761. Secured on tobacco duties. 212,000 bonds issued, repayable over 50 years. Text in Bulgarian, French, German, and English.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Bulgarian, French, German, English',
    issueDate: '1902'
  },
  'goetzmann0967.jpg': {
    title: 'Bulgarian Government 5% Gold Loan of 1902 — Conditions & Amortization (Page 2)',
    description: 'Reverse side of the Bulgarian 5% Gold Loan of 1902 showing loan conditions and amortization tables in Bulgarian, French, German, and English.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Bulgarian, French, German, English',
    issueDate: '1902'
  },
  'goetzmann0968.jpg': {
    title: 'Bulgarian Government 5% Gold Loan of 1902 — Talon (Page 3)',
    description: 'Talon (coupon sheet stub) for the Bulgarian 5% Gold Loan of 1902. Text in Bulgarian, French, and English. Labeled "Royaume de Bulgarie" (Kingdom of Bulgaria).',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Bulgarian, French, English',
    issueDate: '1902'
  },
  'goetzmann0969.jpg': {
    title: 'Bulgarian Government 3½% Gold Loan of 1902 — Talon (Page 4)',
    description: 'Talon for the Bulgarian 3½% Gold Loan of 1902. Text in Bulgarian, German, and English. Shows "Königreich Bulgarien" (Kingdom of Bulgaria).',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Bulgaria',
    Period: '20th Century',
    currency: 'Francs',
    language: 'Bulgarian, German, English',
    issueDate: '1902'
  },
  'goetzmann0970.jpg': {
    title: 'Chinese Imperial Government 4½% Gold Loan of 1898, £25 Sterling (Page 1)',
    description: 'Bond from the Chinese Imperial Government for the 4½% Gold Loan of 1898. Denomination of £25 Sterling. Total issue of £16,000,000. Bond number 024187. Issued through the Deutsch-Asiatische Bank. Features ornate red design with Chinese imperial motifs. Text in English and German.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, German',
    issueDate: '1898'
  },
  'goetzmann0971.jpg': {
    title: 'Chinese Imperial Government 4½% Gold Loan of 1898 — Reverse (Page 2)',
    description: 'Reverse side of the Chinese Imperial Government 4½% Gold Loan of 1898 showing extracts from the agreement and table of drawings/amortization plan. Attached coupons numbered in red.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '19th Century',
    currency: 'GBP',
    language: 'English, German',
    issueDate: '1898'
  },
  'goetzmann0974.jpg': {
    title: 'French Colonial Exposition 1931 Lottery Bond (Bon à Lot), 60 Francs (Page 1)',
    description: 'Lottery bond (Bon à Lot) of 60 Francs issued for the 1931 International Colonial Exposition in Paris. 2,300,000 bonds issued. Series 015, Number 08632. Issued by Crédit Foncier de France. Includes attached travel voucher stubs for railways, maritime transport, and aviation.',
    type: ['Bond', 'Debt', 'Lottery'].join(SEP),
    subjectCountry: 'France',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French',
    issueDate: '1931'
  },
  'goetzmann0975.jpg': {
    title: 'French Colonial Exposition 1931 Lottery Bond — Prize Tables (Page 2)',
    description: 'Reverse side of the Colonial Exposition 1931 lottery bond showing Tableau des Tirages et des Lots (prize schedule) and travel discount benefits.',
    type: ['Bond', 'Debt', 'Lottery'].join(SEP),
    subjectCountry: 'France',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French',
    issueDate: '1931'
  },
  // === Batch 9: 0980-0984 ===
  'goetzmann0980.jpg': {
    title: 'Grand Russian Railway Company 3% Bond, Third Issue, 125 Rubles (Page 1)',
    description: 'Bond of the Grand Russian Railway Company (Glavnoye Obshchestvo Rossiyskikh Zheleznykh Dorog) at 3% interest, Third Issue. Denomination of 125 Rubles Silver. Headquartered in St. Petersburg. Green colored bond with ornate border.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian'
  },
  'goetzmann0981.jpg': {
    title: 'Grand Russian Railway Company 3% Bond — Reverse (Page 2)',
    description: 'Reverse side of the Grand Russian Railway Company 3% bond showing text in French, German, and English. One obligation at 125 Rubles / 500 Francs / £20. Includes amortization table.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian, French, German, English'
  },
  'goetzmann0982.jpg': {
    title: 'Grand Russian Railway Company 3% Bond — Coupon Sheet (Page 3)',
    description: 'Coupon sheet for the Grand Russian Railway Company 3% bond. Green paper with grid of numbered coupons in Russian.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian'
  },
  'goetzmann0983.jpg': {
    title: 'Grand Russian Railway Company 3% Bond — Coupon Sheet (Page 4)',
    description: 'Additional coupon sheet for the Grand Russian Railway Company 3% bond. Text in French and English: "Obligations 3% de la Grande Société des Chemins de Fer Russes / 3% Bonds of the Grand Russian Railway Company."',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Francs',
    language: 'French, English'
  },
  'goetzmann0984.jpg': {
    title: 'Province of Nova Scotia Government Redeemable Stock, £1,000, 1914',
    description: 'Government redeemable stock of the Province of Nova Scotia (Dominion of Canada), bearing interest at 3½% per annum. Denomination of £1,000. Number C 1515. Dated October 19, 1914. Transferable at the National Provincial Bank of England, London. Stamped "CANCELLED." Pink/salmon colored.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Canada',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1914'
  },
  // === Batch 10: 0985-0989 ===
  'goetzmann0985.jpg': {
    title: 'Province of Nova Scotia Government Redeemable Stock — Coupon Sheet (Page 2)',
    description: 'Coupon sheet for the Province of Nova Scotia government redeemable stock. Green coupons with £17.10.0 amounts. Number C 1515.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Canada',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1914'
  },
  'goetzmann0986.jpg': {
    title: 'Province of Nova Scotia Government Redeemable Stock — Coupon Sheet (Page 3)',
    description: 'Additional coupon sheet for the Province of Nova Scotia government redeemable stock. Red/pink coupons with references to the National Provincial Bank of England.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Canada',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1914'
  },
  'goetzmann0987.jpg': {
    title: 'Province of Nova Scotia Government Redeemable Stock — Certificate Cover (Page 4)',
    description: 'Cover/outer wrapper of the Province of Nova Scotia government redeemable stock certificate showing two side-by-side decorative titles: "Dominion of Canada, Province of Nova Scotia, Government Redeemable Stock, £1,000." Red/salmon design.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Canada',
    Period: '20th Century',
    currency: 'GBP',
    language: 'English',
    issueDate: '1914'
  },
  'goetzmann0988.jpg': {
    title: 'Banque Industrielle de Chine (Industrial Bank of China), Ordinary Share, 500 Francs (Page 1)',
    description: 'Ordinary share certificate of the Banque Industrielle de Chine (Industrial Bank of China). Société Anonyme with capital of 150,000,000 Francs. Headquarters in Paris. Share of 500 Francs to bearer. Number 258044. Beautiful Art Deco/Chinese-style design with pagodas, dragons, and cityscape.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'China',
    issuingCountry: 'France',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French, Chinese'
  },
  'goetzmann0989.jpg': {
    title: 'Banque Industrielle de Chine — Statutes & Coupons (Page 2)',
    description: 'Reverse side of the Banque Industrielle de Chine share showing Extraits des Statuts (bylaws extracts) in French. Bottom has dividend coupons with Chinese characters (中法實業銀行 息票).',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'China',
    issuingCountry: 'France',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French, Chinese'
  },
  // === Batch 11: 0990-0993, 0996-0997 ===
  'goetzmann0990.jpg': {
    title: 'Chemins de Fer Économiques de l\'Est Égyptien, 3½% Bond, £20, 1897 (Page 1)',
    description: 'Bond of the Compagnie des Chemins de Fer Économiques de l\'Est Égyptien (Economic Railways of Eastern Egypt). 3½% bond to bearer, denomination of £20 Sterling / 504 Francs. Headquarters in Cairo. Capital of £200,000. Issue of 12,500 bonds. Blue ornate design. Dated July 1, 1897.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Egypt',
    Period: '19th Century',
    currency: 'GBP',
    language: 'French, English',
    issueDate: '1897'
  },
  'goetzmann0991.jpg': {
    title: 'Chemins de Fer Économiques de l\'Est Égyptien — Guarantees & Amortization (Page 2)',
    description: 'Reverse of the Eastern Egyptian Railway bond showing guarantees and sinking fund amortization table in French and English. Mentions railway concession in provinces of Charkieh, Dakahlieh, and Kalioubieh.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Egypt',
    Period: '19th Century',
    currency: 'GBP',
    language: 'French, English',
    issueDate: '1897'
  },
  'goetzmann0992.jpg': {
    title: 'Chemins de Fer Économiques de l\'Est Égyptien — Coupon Sheet (Page 3)',
    description: 'Coupon sheet for the Eastern Egyptian Railway 3½% bond. Blue coupons with obligation numbers.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Egypt',
    Period: '19th Century',
    currency: 'GBP',
    language: 'French, English',
    issueDate: '1897'
  },
  'goetzmann0993.jpg': {
    title: 'Chemins de Fer Économiques de l\'Est Égyptien — Coupon Sheet (Page 4)',
    description: 'Additional coupon sheet for the Eastern Egyptian Railway bond. Coupons for £0.7.0 each, payable in Cairo, Alexandria, Paris, London, Amsterdam, and Brussels.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Egypt',
    Period: '19th Century',
    currency: 'GBP',
    language: 'French, English',
    issueDate: '1897'
  },
  'goetzmann0996.jpg': {
    title: 'Count Casimir Esterházy Partial Bond (Partial-Schuldverschreibung), 20 Gulden, 1847 (Page 1)',
    description: 'Partial bond (Partial-Schuldverschreibung) of Count Casimir Esterházy von Galántha. Denomination of 20 Gulden in Convention Coin. Number F 14725. Issued through the Esterházy Central Treasury. Features coat of arms and repayment tables.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Austria',
    Period: '19th Century',
    currency: 'Gulden',
    language: 'German',
    issueDate: '1847'
  },
  'goetzmann0997.jpg': {
    title: 'Count Casimir Esterházy Partial Bond — Repayment Plan (Page 2)',
    description: 'Reverse side of the Esterházy partial bond showing the repayment plan (Rückzahlungs-Plan) for 6,000 partial bonds. Dated Vienna, December 1, 1847. Signed and issued through the Esterházy Central Treasury.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Austria',
    Period: '19th Century',
    currency: 'Gulden',
    language: 'German',
    issueDate: '1847'
  },
  // === Batch 12: 0998-1002 ===
  'goetzmann0998.jpg': {
    title: 'USSR Third State Loan for Restoration and Development of National Economy, 100 Rubles (Page 1)',
    description: 'Bond for the Third State Loan for the Restoration and Development of the National Economy of the USSR. Denomination of 100 Rubles. Features Soviet coat of arms and industrial imagery (dam, factories, tractors). Post-WWII reconstruction era.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1948'
  },
  'goetzmann0999.jpg': {
    title: 'USSR Third State Loan — Conditions (Page 2)',
    description: 'Reverse side of the USSR Third State Loan bond showing the terms and conditions in Russian, including lottery drawing details.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1948'
  },
  'goetzmann1000.jpg': {
    title: 'USSR State Internal Prize Loan of 1982, 50 Rubles (Page 1)',
    description: 'State Internal Prize Loan bond of the USSR from 1982. Denomination of 50 Rubles. Features Soviet coat of arms. Green/brown design.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1982'
  },
  'goetzmann1001.jpg': {
    title: 'USSR State Internal Prize Loan of 1982 — Reverse (Page 2)',
    description: 'Reverse side of the USSR State Internal Prize Loan of 1982 showing conditions and prize structure.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '20th Century',
    currency: 'Rubles',
    language: 'Russian',
    issueDate: '1982'
  },
  'goetzmann1002.jpg': {
    title: 'Republic of China Ministry of Finance Second Prize Bond, 5 Yuan',
    description: 'Prize bond of the Republic of China Ministry of Finance, Second Issue. Denomination of 5 Yuan. Purple design with ornate circular patterns. Text in traditional Chinese characters.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '20th Century',
    currency: 'Yuan',
    language: 'Chinese'
  },
  // === Batch 13: 1003-1007 ===
  'goetzmann1003.jpg': {
    title: 'Second Nationalist Government Lottery Loan, $5 Canton Currency, 1926',
    description: 'Second Nationalist Government Lottery Loan of the Fifteenth Year of the Republic of China (1926). Note for Five Dollars, Canton Currency. Repayment from revenue of the City of Wuchow. Green certificate.',
    type: ['Bond', 'Debt', 'Lottery'].join(SEP),
    subjectCountry: 'China',
    Period: '20th Century',
    currency: 'Dollars (Canton)',
    language: 'English, Chinese',
    issueDate: '1926'
  },
  'goetzmann1004.jpg': {
    title: 'Imperial Russian Consolidated 4% Railway Bond, 1st Series, 125 Gold Rubles (Page 1)',
    description: 'Consolidated railway bond of the Imperial Russian Government at 4% interest, 1st Series. Denomination of 125 Gold Rubles. Brown/chocolate design with double-headed eagle. Revenue stamp attached.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'Russian'
  },
  'goetzmann1005.jpg': {
    title: 'Imperial Russian Consolidated 4% Railway Bond — Reverse (Page 2)',
    description: 'Reverse of the Imperial Russian Consolidated 4% Railway Bond showing text in French, German, and English. Bond of 125 Rubles Gold = 500 Francs = 404 Marks = £19 Sterling. Payment locations in St. Petersburg, Paris, London, Berlin, Brussels, and Amsterdam.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    Period: '19th Century',
    currency: 'Rubles',
    language: 'French, German, English'
  },
  'goetzmann1006.jpg': {
    title: 'Shanghai Pudong Qiangsheng Taxi Co., Ltd., Share Certificate, 10 Shares (100 RMB) (Page 1)',
    description: 'Share certificate of Shanghai Pu Dong Qiang Sheng Taxi Co., Ltd. (上海浦東強生出租汽車股份有限公司). Certificate for 10 shares with total value of 100 RMB. Green design with illustrations of a taxi and bus. Dated February 12, 1992.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '20th Century',
    currency: 'RMB',
    language: 'Chinese, English',
    issueDate: '1992'
  },
  'goetzmann1007.jpg': {
    title: 'Shanghai Pudong Qiangsheng Taxi Co., Ltd. — Reverse (Page 2)',
    description: 'Reverse side of the Shanghai Pudong Qiangsheng Taxi share certificate showing explanatory notes (說明), terms and conditions, and share transfer record table. Red official seal.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'China',
    Period: '20th Century',
    currency: 'RMB',
    language: 'Chinese',
    issueDate: '1992'
  },
  // === Batch 14: 1008-1011, 1022-1026 ===
  'goetzmann1008.jpg': {
    title: 'L\'Uranium S.A. (Belgian Congo) — Statutes Extract',
    description: 'Extract from the statutes (Extrait des Statuts) of L\'Uranium, a société anonyme based in the Congo with registration in Brussels. Mentions share capital and administrative structure. Part of a share certificate.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Belgium',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French'
  },
  'goetzmann1009.jpg': {
    title: 'L\'Uranium S.A. — Dividend Coupon Sheet',
    description: 'Dividend coupon sheet numbered 1-30 in decorative design, accompanying the L\'Uranium S.A. share certificate.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Belgium',
    Period: '20th Century',
    currency: 'Francs',
    language: 'French'
  },
  'goetzmann1010.jpg': {
    title: 'Receipt for Grand Russian Railway Company 3% Bond Talons, 3rd Issue 1881',
    description: 'Receipt from Lippmann, Rosenthal & Co., Amsterdam, for talons of 3% bonds of the Grand Russian Railway Company (Groote Russische Spoorweg Maatschappij), 3rd issue 1881. Bond number 111791. Dated Amsterdam, May 12, 1924. Transferred from the Boedel, Königsberg estate.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    currency: 'GBP',
    language: 'Dutch',
    issueDate: '1881'
  },
  'goetzmann1011.jpg': {
    title: 'Receipt for Grand Russian Railway Bond Talons — Reverse',
    description: 'Reverse side of the receipt for Grand Russian Railway Company bond talons. Mostly blank with return address text in Dutch.',
    type: ['Bond', 'Debt', 'Security'].join(SEP),
    subjectCountry: 'Russia',
    issuingCountry: 'Netherlands',
    Period: '19th Century',
    language: 'Dutch',
    issueDate: '1881'
  },
  'goetzmann1022.jpg': {
    title: 'Dutch Company Certificate, Leidijk van den Woerdans, Utrecht, 25 Guilders',
    description: 'Handwritten Dutch certificate related to the Maatschappij Leidijk van den Woerdans, based in Utrecht, Netherlands. Denomination of 25 Guilders. Features columns of numerical records alongside the certificate text. Early 20th century.',
    type: ['Equity', 'Security'].join(SEP),
    subjectCountry: 'Netherlands',
    Period: '20th Century',
    currency: 'Guilders',
    language: 'Dutch'
  },
  'goetzmann1023.jpg': {
    title: 'Historical Manuscript/Parchment Document (Page 1)',
    description: 'Very old parchment/vellum document with handwritten text in 17th-century script. Aged and damaged with torn edges. Possibly a deed, bond, or early financial document. Extremely difficult to read due to age and deterioration.',
    type: 'Document',
    Period: '18th Century or before',
    language: 'English'
  },
  'goetzmann1024.jpg': {
    title: 'Historical Manuscript/Parchment Document — Reverse (Page 2)',
    description: 'Reverse side of the historical parchment document showing various handwritten endorsements and records. Heavily worn and aged.',
    type: 'Document',
    Period: '18th Century or before',
    language: 'English'
  },
  'goetzmann1025.jpg': {
    title: 'Ming Dynasty Treasure Note (Da Ming Baochao), 1 Guan (Page 1)',
    description: 'Ming Dynasty paper currency note (大明通行寳鈔, Da Ming Tongxing Baochao). Denomination of 1 Guan (壹貫, one string of cash). Issued by the Ministry of Revenue (戶部) during the Hongwu era (~1368-1398). Features traditional border design with dragons and official seals in red. Extraordinarily rare piece of early Chinese paper money.',
    type: ['Currency', 'Banknote'].join(SEP),
    subjectCountry: 'China',
    Period: '18th Century or before',
    currency: 'Guan',
    language: 'Chinese',
    issueDate: 'c. 1375'
  },
  'goetzmann1026.jpg': {
    title: 'Ming Dynasty Treasure Note (Da Ming Baochao) — Reverse (Page 2)',
    description: 'Reverse side of the Ming Dynasty treasure note showing faded printing and a red official seal. Very worn and aged.',
    type: ['Currency', 'Banknote'].join(SEP),
    subjectCountry: 'China',
    Period: '18th Century or before',
    currency: 'Guan',
    language: 'Chinese',
    issueDate: 'c. 1375'
  }
};

// Read the spreadsheet
const wb = XLSX.readFile(SPREADSHEET);
const ws = wb.Sheets[wb.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(ws);

let updated = 0;
let skipped = 0;
let notFound = 0;

for (const row of data) {
  const fn = row.filename;
  if (metadata[fn]) {
    const m = metadata[fn];
    // Only update if the row is currently empty
    if (!row.title && !row.description && !row.type && !row.subjectCountry && !row.Period) {
      if (m.title && m.title !== '[No thumbnail available]') {
        row.title = m.title;
        row.description = m.description || '';
        row.type = m.type || '';
        row.subjectCountry = m.subjectCountry || '';
        if (m.issuingCountry) row.issuingCountry = m.issuingCountry;
        row.Period = m.Period || '';
        if (m.currency) row.currency = m.currency;
        if (m.language) row.language = m.language;
        if (m.issueDate) row.issueDate = m.issueDate;
        updated++;
      } else {
        skipped++;
      }
    }
  }
}

// Write back
const newWs = XLSX.utils.json_to_sheet(data);
wb.Sheets[wb.SheetNames[0]] = newWs;
XLSX.writeFile(wb, SPREADSHEET);

console.log('Updated:', updated, 'documents');
console.log('Skipped (no data):', skipped);
console.log('Total empty docs with metadata entries:', Object.keys(metadata).length);
