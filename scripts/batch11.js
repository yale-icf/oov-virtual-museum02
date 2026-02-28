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

// --- 24 Geo. II Acts of Parliament (pages 13-18, rows 286-291) ---

const geoII24Base = {
  type: 'Legislative Document',
  subjectCountry: 'United Kingdom',
  issuingCountry: 'United Kingdom',
  creator: 'Parliament of Great Britain',
  issueDate: '1750-01-01',
  currency: '',
  language: 'English',
  numberPages: 18,
  period: '18th Century or before',
  notes: 'Acts of Parliament, 24 George II, 1750 (Great Britain)',
};

setDoc(286, 'Acts of Parliament, 24 Geo. II (1750) (Page 13 of 18)',
  'This page (printed p. 90) continues the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750). It contains statutory text of one or more acts regulating trade, public finance, or colonial administration, printed in the standard double-column folio format used for Georgian era parliamentary publications.',
  geoII24Base);

setDoc(287, 'Acts of Parliament, 24 Geo. II (1750) (Page 14 of 18)',
  'This page (printed p. 91) continues the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750). It contains statutory text printed in the characteristic double-column folio layout of eighteenth-century British parliamentary printing.',
  geoII24Base);

setDoc(288, 'Acts of Parliament, 24 Geo. II (1750) (Page 15 of 18)',
  'This page (printed p. 92) continues the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750). The double-column text presents the standard legal language of Georgian-era statutes, organized by act and section.',
  geoII24Base);

setDoc(289, 'Acts of Parliament, 24 Geo. II (1750) (Page 16 of 18)',
  'This page (printed p. 93) continues the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750). As the compilation nears its end, this page contains the concluding provisions of the acts included in this session.',
  geoII24Base);

setDoc(290, 'Acts of Parliament, 24 Geo. II (1750) (Page 17 of 18)',
  'This page (printed p. 94) is among the final pages of the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750), containing concluding statutory text before the terminal page.',
  geoII24Base);

setDoc(291, 'Acts of Parliament, 24 Geo. II (1750) (Page 18 of 18)',
  'This final page (printed p. 95) of the compilation of Acts of Parliament passed in the twenty-fourth year of the reign of King George II (1750) concludes with the word "FINIS," marking the end of the volume. It represents the last page of this official parliamentary publication, which compiled the statutes enacted during the 1750 session of the Parliament of Great Britain.',
  geoII24Base);

// --- Individual documents (rows 292-335) ---

setDoc(292, 'Baltimore and Ohio Railroad Company, Common Stock',
  'This engraved stock certificate represents 100 shares of Common Stock in the Baltimore and Ohio Railroad Company, Certificate No. C243475, issued to Evans Stillman & Co. on June 19, 1934. The Baltimore and Ohio Railroad, chartered in 1827, was the first commercial railroad in the United States, connecting Baltimore to the Ohio River. By 1934 the B&O had become one of the great trunk railroads of the eastern United States, though it faced severe financial strain during the Great Depression.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Baltimore and Ohio Railroad Company', issueDate: '1934-06-19', currency: 'USD', language: 'English', numberPages: 1, period: '20th Century', notes: 'Baltimore and Ohio Railroad Company, Common Stock, 100 shares, No. C243475, Evans Stillman & Co., June 19, 1934' });

setDoc(293, 'Baltimore and Ohio Rail Road Company, 6% Preferred Stock',
  'This engraved stock certificate represents shares of 6% Preferred Stock in the Baltimore and Ohio Rail Road Company, Certificate No. 466, issued circa September 1863. The B&O Railroad, the oldest commercial railroad in America, issued preferred stock during the Civil War era to fund operations and repairs as military activity along its lines—which ran through the border states of Maryland and Virginia—caused enormous damage to its infrastructure.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Baltimore and Ohio Rail Road Company', issueDate: '1863-09-07', currency: 'USD', language: 'English', numberPages: 1, period: '19th Century', notes: 'Baltimore and Ohio Rail Road Company, 6% Preferred Stock, No. 466, ca. September 1863' });

setDoc(294, 'Banca Românească Share Certificate, Emisiunea VIII-A (1938)',
  'This engraved Romanian share certificate represents one share of 500 lei par value in Banca Românească (Romanian Bank), Emisiunea VIII-A (8th-A Emission), issued in Bucharest in 1938, when the company\'s total capital stood at 350,000,000 lei. Banca Românească, founded in 1911, was one of Romania\'s leading commercial banks, playing a central role in financing Romanian industrial and commercial development during the interwar period. The 1938 issuance occurred under the royal dictatorship of Carol II, in the last years before Romania was drawn into World War II.',
  { type: 'Stock Certificate', subjectCountry: 'Romania', issuingCountry: 'Romania', creator: 'Banca Românească', issueDate: '1938-01-01', currency: 'ROL', language: 'Romanian', numberPages: 1, period: '20th Century', notes: 'Banca Românească, 500 lei share, Emisiunea VIII-A, Capital 350,000,000 lei, Bucharest 1938' });

setDoc(295, 'Banca Românească Share Certificate (1920)',
  'This engraved Romanian share certificate represents one share of 500 lei par value in Banca Românească (Romanian Bank), No. 011666, with a capital of 100,000,000 lei, issued in Bucharest in 1920. This emission dates from the immediate postwar period, when Romania—greatly enlarged after World War I by the acquisition of Transylvania, Bessarabia, Bukovina, and southern Dobruja—was reorganizing its banking sector to serve the expanded national economy. Banca Românească was one of the principal institutions financing the industrialization of Greater Romania.',
  { type: 'Stock Certificate', subjectCountry: 'Romania', issuingCountry: 'Romania', creator: 'Banca Românească', issueDate: '1920-01-01', currency: 'ROL', language: 'Romanian', numberPages: 1, period: '20th Century', notes: 'Banca Românească, 500 lei share No. 011666, Capital 100,000,000 lei, Bucharest 1920' });

setDoc(296, 'Banco Territorial de Cuba (Crédit Foncier Cubain), Acción Beneficiaria',
  'This certificate represents an Acción Beneficiaria (Beneficiary Share) No. 22752 of the Banco Territorial de Cuba, also known as the Crédit Foncier Cubain, issued in Havana on March 1, 1911. The Banco Territorial de Cuba was a French-backed mortgage bank financing land transactions and real estate in post-independence Cuba. Beneficiary shares (acciones beneficiarias) were a founders\' share form common in Latin American financial institutions, entitling holders to a portion of profits without a fixed par value.',
  { type: 'Stock Certificate', subjectCountry: 'Cuba', issuingCountry: 'Cuba', creator: 'Banco Territorial de Cuba', issueDate: '1911-03-01', currency: 'CUP', language: 'Spanish', numberPages: 1, period: '20th Century', notes: 'Banco Territorial de Cuba (Crédit Foncier Cubain), Acción Beneficiaria No. 22752, Havana, March 1, 1911' });

setDoc(297, 'Bank of Roumania Limited, Share Warrant',
  'This ornately printed share warrant, No. 5205, certifies that the bearer is the owner of five shares (Nos. 17771–17775) in the Bank of Roumania Limited, each share worth £6, issued in London in May 1903. The Bank of Roumania Limited was a British-registered bank financing commerce and industry in Romania, part of the wave of Western European capital flowing into Eastern European developing economies at the turn of the twentieth century.',
  { type: 'Share Warrant', subjectCountry: 'Romania', issuingCountry: 'United Kingdom', creator: 'Bank of Roumania Limited', issueDate: '1903-05-01', currency: 'GBP', language: 'English', numberPages: 1, period: '20th Century', notes: 'Bank of Roumania Limited, Share Warrant No. 5205, 5 shares (Nos. 17771-17775), £6 each, London, May 1903' });

setDoc(298, 'Bank Małopolski, Share Certificate',
  'This share certificate represents 25 shares in the Bank Małopolski (Bank of Little Poland) of Kraków, valued at 10,000 Koron / 7,000 Polish Marks, issued December 15, 1920. Bank Małopolski was a regional bank serving the formerly Austrian Galicia, a territory incorporated into the reconstituted Polish state after World War I. The dual denomination (Austrian Koron and Polish Marks) reflects the currency transition underway as Poland consolidated financial institutions inherited from three different partitioning empires.',
  { type: 'Stock Certificate', subjectCountry: 'Poland', issuingCountry: 'Poland', creator: 'Bank Małopolski', issueDate: '1920-12-15', currency: 'PLN', language: 'Polish', numberPages: 1, period: '20th Century', notes: 'Bank Małopolski, Kraków, 25 shares, 10,000 Koron / 7,000 Polish Marks, December 15, 1920' });

setDoc(299, 'Siberian Commercial Bank, Dividend Coupon Sheet',
  'This document is a sheet of five dividend coupons (6th through 10th) for shares Nos. 17747–17748 of the Siberian Commercial Bank (Сибирский Торговый Банк / Banque Commerciale de Sibérie), covering dividend years 1917 through 1921. The Siberian Commercial Bank was one of Russia\'s major regional banks, headquartered in St. Petersburg with extensive Siberian operations. These coupons—representing dividend payments spanning the Bolshevik Revolution—were likely never redeemed, as the bank was nationalized after October 1917.',
  { type: 'Dividend Coupon', subjectCountry: 'Russia', issuingCountry: 'Russia', creator: 'Siberian Commercial Bank', issueDate: '1917-01-01', currency: 'RUB', language: 'Russian', numberPages: 1, period: '20th Century', notes: 'Siberian Commercial Bank (Banque Commerciale de Sibérie), dividend coupons 6-10 for shares 17747-17748, years 1917-1921' });

setDoc(300, 'St. Petersburg Private Commercial Bank, 200 Ruble Share',
  'This share certificate represents one share of 200 Rubles in the St. Petersburg Private Commercial Bank (Санкт-Петербургский Частный Коммерческий Банк / Banque de Commerce Privée de St.-Pétersbourg), Share No. 82719, Emission of 1911. Founded in 1864, the bank was one of Russia\'s leading joint-stock commercial banks. The 1911 emission reflected the bank\'s expansion during Russia\'s rapid pre-war industrialization. After the Bolshevik Revolution, the bank was nationalized and its shares became worthless.',
  { type: 'Stock Certificate', subjectCountry: 'Russia', issuingCountry: 'Russia', creator: 'St. Petersburg Private Commercial Bank', issueDate: '1911-01-01', currency: 'RUB', language: 'Russian', numberPages: 1, period: '20th Century', notes: 'St. Petersburg Private Commercial Bank (Banque de Commerce Privée de St.-Pétersbourg), 200 Ruble share No. 82719, Emission 1911' });

setDoc(301, 'Sofiyska Banka (Banque de Sofia), Share Certificate',
  'This share certificate represents one share of 100 gold leva in the Sofiyska Banka (Banque de Sofia / Bank of Sofia), No. 035685, issued in Bulgaria in 1917. The denomination in gold leva—rather than regular currency—protected investors against the wartime inflation ravaging Bulgaria during World War I, when the country fought alongside the Central Powers. The Bank of Sofia was one of Bulgaria\'s principal private financial institutions in the early twentieth century.',
  { type: 'Stock Certificate', subjectCountry: 'Bulgaria', issuingCountry: 'Bulgaria', creator: 'Sofiyska Banka', issueDate: '1917-01-01', currency: 'BGN', language: 'Bulgarian', numberPages: 1, period: '20th Century', notes: 'Sofiyska Banka (Banque de Sofia), 100 gold leva share No. 035685, Bulgaria, 1917' });

setDoc(302, 'Banque Industrielle de Chine (中国实业银行), 500 Franc Share',
  'This engraved share certificate represents one action of 500 Francs in the Banque Industrielle de Chine (中国实业银行 / Industrial Bank of China), No. 47221, issued in Paris. Founded in 1913 with Franco-Chinese capital, the Banque Industrielle de Chine aimed to finance Chinese industrial development and had branches in major Chinese cities. The bank ultimately failed in 1921 during the post-war global financial crisis—its collapse caused a major scandal in France, where thousands of small investors lost their savings.',
  { type: 'Stock Certificate', subjectCountry: 'China', issuingCountry: 'France', creator: 'Banque Industrielle de Chine', issueDate: '1913-01-01', currency: 'FRF', language: 'French', numberPages: 1, period: '20th Century', notes: 'Banque Industrielle de Chine (中国实业银行), 500 Franc action No. 47221, Paris' });

setDoc(303, 'Banque Territoriale d\'Espagne, 500 Franc Share',
  'This engraved share certificate represents one action of 500 Francs in the Banque Territoriale d\'Espagne, No. 011,824, issued in Madrid on May 30, 1871. The Banque Territoriale d\'Espagne was a French-Spanish land mortgage bank providing long-term financing for real estate and agriculture in Spain. It was modeled on France\'s Crédit Foncier and part of a broader European movement to establish mortgage banking institutions that could mobilize capital for agricultural improvement and urban development.',
  { type: 'Stock Certificate', subjectCountry: 'Spain', issuingCountry: 'Spain', creator: 'Banque Territoriale d\'Espagne', issueDate: '1871-05-30', currency: 'FRF', language: 'French', numberPages: 1, period: '19th Century', notes: 'Banque Territoriale d\'Espagne, 500 Franc action No. 011,824, Madrid, May 30, 1871' });

setDoc(304, 'Beate Uhse Aktiengesellschaft, 1 Euro Share',
  'This stock certificate represents one share of 1 Euro par value in Beate Uhse Aktiengesellschaft, No. 000069560, issued in Flensburg, Germany in May 1999. Beate Uhse AG—founded in 1946 by former Luftwaffe test pilot Beate Uhse—was a pioneering German adult entertainment retailer. When the company went public on the Frankfurt Stock Exchange in May 1999, it became one of the first businesses in its industry to do so, attracting global media attention and representing a landmark in postwar German cultural liberalization. The company later went bankrupt in 2017.',
  { type: 'Stock Certificate', subjectCountry: 'Germany', issuingCountry: 'Germany', creator: 'Beate Uhse Aktiengesellschaft', issueDate: '1999-05-01', currency: 'EUR', language: 'German', numberPages: 1, period: '21st Century', notes: 'Beate Uhse Aktiengesellschaft, 1 Euro share No. 000069560, Flensburg, May 1999; first adult entertainment company to go public on the Frankfurt Stock Exchange' });

setDoc(305, 'Black-Sea-Kuban Railway Company, 4½% Bond',
  'This bond certificate represents one obligation of £20 (equivalent to 189 Rubles) at 4½% of the Black-Sea-Kuban Railway Company, Bond No. A13852, issued at Yekaterinador (now Krasnodar) in 1911. The Black-Sea-Kuban Railway served the fertile Kuban agricultural region of southern Russia, connecting it to Black Sea ports. The bond\'s dual denomination in British pounds and Russian rubles reflects the practice of marketing Russian railway bonds to international investors through London and other European capital markets.',
  { type: 'Bond', subjectCountry: 'Russia', issuingCountry: 'Russia', creator: 'Black-Sea-Kuban Railway Company', issueDate: '1911-01-01', currency: 'GBP', language: 'English', numberPages: 1, period: '20th Century', notes: 'Black-Sea-Kuban Railway Company, 4½% Bond £20 = 189 Rubles, No. A13852, Yekaterinador, 1911' });

setDoc(306, 'Boston and Providence Rail Road Corporation, Stock Certificate',
  'This engraved stock certificate represents ten shares in the Boston and Providence Rail Road Corporation, Certificate No. 122, issued to Haliburton Fales on August 15, 1839. The Boston and Providence Railroad, incorporated in 1831 and opened in 1835, was among the earliest railroads in the United States, connecting Boston to Providence, Rhode Island. By the time this certificate was issued, the railroad had already proven its commercial success and was attracting substantial investment.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Boston and Providence Rail Road Corporation', issueDate: '1839-08-15', currency: 'USD', language: 'English', numberPages: 1, period: '19th Century', notes: 'Boston and Providence Rail Road Corporation, 10 shares, No. 122, Haliburton Fales, August 15, 1839' });

setDoc(307, 'British Honduras Company Limited, Share Certificate',
  'This share certificate, No. 12240, represents one fully paid share of £5 in the British Honduras Company Limited, issued to John Campbell on June 1, 1863. British Honduras (present-day Belize) was a British colony economically centered on mahogany timber extraction. The British Honduras Company was a joint-stock company formed to exploit the colony\'s timber resources and promote commercial settlement, exemplifying the mid-Victorian use of the limited liability company form to finance colonial enterprise.',
  { type: 'Stock Certificate', subjectCountry: 'Belize', issuingCountry: 'United Kingdom', creator: 'British Honduras Company Limited', issueDate: '1863-06-01', currency: 'GBP', language: 'English', numberPages: 1, period: '19th Century', notes: 'British Honduras Company Limited, Share No. 12240, £5, John Campbell, June 1, 1863' });

setDoc(308, 'Calvert, Waco & Brazos Valley Railroad Company, Specimen Stock Certificate',
  'This unissued specimen stock certificate of the Calvert, Waco & Brazos Valley Railroad Company is incorporated under the laws of the State of Texas. All blanks for shareholder name, share quantity, and date are unfilled, indicating this is a printer\'s sample or archive copy rather than an issued certificate. The Calvert, Waco & Brazos Valley Railroad was a Texas short-line project typical of the hundreds of small railroad companies that sought to connect agricultural communities to trunk lines during the late nineteenth-century railroad boom.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Calvert, Waco & Brazos Valley Railroad Company', issueDate: '1890-01-01', currency: 'USD', language: 'English', numberPages: 1, period: '19th Century', notes: 'Calvert, Waco & Brazos Valley Railroad Company, blank specimen certificate, State of Texas' });

setDoc(309, 'Compagnie Universelle du Canal Interocéanique de Panama, 500 Franc Share',
  'This engraved share certificate represents one action of 500 Francs in the Compagnie Universelle du Canal Interocéanique de Panama, No. 185,229, accompanied by a 1886 billet de finance coupon. The Panama Canal Company was organized in 1879 by Ferdinand de Lesseps, who mobilized capital from hundreds of thousands of French small investors to build a sea-level canal across the Isthmus of Panama. The project collapsed in 1889 amid engineering failures and financial fraud, triggering the Panama Scandal—the greatest financial and political scandal of nineteenth-century France and the ruin of hundreds of thousands of investors.',
  { type: 'Stock Certificate', subjectCountry: 'Panama', issuingCountry: 'France', creator: 'Compagnie Universelle du Canal Interocéanique de Panama', issueDate: '1886-01-01', currency: 'FRF', language: 'French', numberPages: 1, period: '19th Century', notes: 'Compagnie Universelle du Canal Interocéanique de Panama, 500 Franc action No. 185,229, with 1886 billet de finance coupon' });

setDoc(310, 'Compagnie Universelle du Canal Maritime de Suez, 3% Obligation',
  'This engraved bond certificate represents one 3% obligation of 500 Francs, 2e Série, No. 271,266, of the Compagnie Universelle du Canal Maritime de Suez. The Suez Canal Company, founded in 1858 by Ferdinand de Lesseps, opened the Suez Canal in 1869, transforming global trade by linking the Mediterranean and Red Seas. The company issued multiple series of bonds to finance construction and maintenance of one of the most consequential infrastructure projects of the nineteenth century.',
  { type: 'Bond', subjectCountry: 'Egypt', issuingCountry: 'France', creator: 'Compagnie Universelle du Canal Maritime de Suez', issueDate: '1870-01-01', currency: 'FRF', language: 'French', numberPages: 1, period: '19th Century', notes: 'Canal Maritime de Suez, Compagnie Universelle, 2e Série, 500 Franc 3% obligation No. 271,266' });

setDoc(311, 'Caramanian Iron Corporation Limited, Share Warrant',
  'This share warrant, No. 00,689, certifies the bearer\'s ownership of ten shares (Nos. 06,881–06,890) in the Caramanian Iron Corporation Limited, issued in London and Paris on February 4, 1907. The Caramanian Iron Corporation held mining concessions in Karamania (the coastal Cilicia region of present-day southern Turkey), and was part of the wave of British and European capital flowing into Ottoman mining enterprises during the late Ottoman period, when the empire granted extensive concessions to attract foreign investment.',
  { type: 'Share Warrant', subjectCountry: 'Turkey', issuingCountry: 'United Kingdom', creator: 'Caramanian Iron Corporation Limited', issueDate: '1907-02-04', currency: 'GBP', language: 'English', numberPages: 1, period: '20th Century', notes: 'Caramanian Iron Corporation Limited, Share Warrant No. 00,689, 10 shares (06,881-06,890), London/Paris, February 4, 1907' });

setDoc(312, 'Central Trust Company, Stock Certificate',
  'This stock certificate represents 1,030 shares of the Central Trust Company of Cambridge, Massachusetts, Certificate No. 1636, issued to Edgar R. Champlin on July 21, 1929—just three months before the Wall Street Crash of October 1929. The Central Trust Company was a Massachusetts state-chartered trust company serving the Greater Boston area. Trust companies of this era operated as both commercial banks and investment institutions, making them highly vulnerable to the bank runs and asset deflation of the subsequent Depression.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Central Trust Company', issueDate: '1929-07-21', currency: 'USD', language: 'English', numberPages: 1, period: '20th Century', notes: 'Central Trust Company, Cambridge MA, 1,030 shares, No. 1636, Edgar R. Champlin, July 21, 1929' });

setDoc(313, 'Amsterdam Depositary Certificate for Bank of the United States Shares',
  'This Amsterdam depositary receipt, No. 36,013, certifies that shares of the Bank of the United States are held in deposit by the Dutch banking houses Hope & Co., Ketwich & Voombergh, and the Widow Willem Borski on behalf of the bearer, dated January 31, 1844. The Second Bank of the United States (1816–1836) had attracted major Dutch investment, and Amsterdam banking houses facilitated European participation in American securities. By 1844 the bank had lost its federal charter and was operating under a Pennsylvania charter before its final failure; these certificates represent the continued settlement of residual claims by foreign investors.',
  { type: 'Depositary Receipt', subjectCountry: 'United States', issuingCountry: 'Netherlands', creator: 'Hope & Co.; Ketwich & Voombergh; Widow Willem Borski', issueDate: '1844-01-31', currency: 'USD', language: 'Dutch', numberPages: 1, period: '19th Century', notes: 'Amsterdam depositary certificate No. 36,013 for Bank of the United States shares, Hope & Co./Ketwich & Voombergh/Widow Willem Borski, January 31, 1844' });

setDoc(314, 'Hamburg Depositary Certificate for Russian 5% State Bonds',
  'This Hamburg depositary receipt, No. 14,463, certifies that bonds of 500 Rubles at 5% of the Russian state are held in deposit by the Hamburg banking houses Sillem Benecke & Co. and H.J. Stresow on behalf of the bearer, dated April 23, 1851. Russian state bonds were widely held across Western Europe in the nineteenth century, and German banking houses in Hamburg served as intermediaries providing depositary and custody services for securities that were difficult to hold and transfer across national borders.',
  { type: 'Depositary Receipt', subjectCountry: 'Russia', issuingCountry: 'Germany', creator: 'Sillem Benecke & Co.; H.J. Stresow', issueDate: '1851-04-23', currency: 'RUB', language: 'German', numberPages: 1, period: '19th Century', notes: 'Hamburg depositary certificate No. 14,463 for 500 Rubles 5% Russian bonds, Sillem Benecke & Co./H.J. Stresow, April 23, 1851' });

setDoc(315, 'Charles Laffitte & Company Limited, Share Certificate',
  'This share certificate, No. 67906, represents one share of £20 par value (half paid) in Charles Laffitte & Company Limited, issued on January 5, 1866. Charles Laffitte was a French-born banker operating in London, associated with the distinguished French banking family of Jacques Laffitte. Charles Laffitte & Company was an Anglo-French banking and trading house involved in international finance during the mid-Victorian era of rapidly expanding global capital markets.',
  { type: 'Stock Certificate', subjectCountry: 'United Kingdom', issuingCountry: 'United Kingdom', creator: 'Charles Laffitte & Company Limited', issueDate: '1866-01-05', currency: 'GBP', language: 'English', numberPages: 1, period: '19th Century', notes: 'Charles Laffitte & Company Limited, One Share £20 (half paid), No. 67906, January 5, 1866' });

setDoc(316, 'Chicago, Rock Island & Pacific Rail Road Company, $5,000 Gold Mortgage Bond',
  'This engraved bond certificate represents a $5,000 Gold Mortgage Bond of the Chicago, Rock Island & Pacific Rail Road Company, Bond No. 1381. The Chicago, Rock Island and Pacific—known as the Rock Island Railroad—was one of the major Midwestern trunk lines, running from Chicago westward across the Great Plains. Gold mortgage bonds were secured by the railroad\'s real property and equipment and payable in gold coin, offering investors protection against currency inflation during periods of financial uncertainty.',
  { type: 'Bond', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Chicago, Rock Island & Pacific Rail Road Company', issueDate: '1880-01-01', currency: 'USD', language: 'English', numberPages: 1, period: '19th Century', notes: 'Chicago, Rock Island & Pacific Rail Road Company, Mortgage Bond No. 1381, $5,000 Gold Bond' });

setDoc(317, 'Chicago, Rock Island and Pacific Railroad Company, $1,000 Gold Bond of 2002',
  'This bond certificate represents a $1,000 Gold Bond of the Chicago, Rock Island and Pacific Railroad Company, due November 1, 2002, No. 530035. The Rock Island Railroad issued these exceptionally long-dated bonds in the mid-twentieth century as part of a major refinancing. The company, which had operated since 1847, went bankrupt in 1975 and was liquidated in 1980—more than two decades before the bonds were due to mature—making them ultimately worthless. The Rock Island\'s prolonged bankruptcy (1975–1984) became one of the most complex railroad reorganizations in US history.',
  { type: 'Bond', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Chicago, Rock Island and Pacific Railroad Company', issueDate: '1960-01-01', currency: 'USD', language: 'English', numberPages: 1, period: '20th Century', notes: 'Chicago, Rock Island and Pacific Railroad, $1,000 Gold Bond of 2002, due November 1, 2002, No. 530035' });

setDoc(318, 'Republic of China Construction Gold Bonds, Year 29 (1940), $5 US Dollar Bond',
  'This bond certificate represents a $5 US Dollar denomination bond from the Republic of China\'s Year 29 Construction Gold Bonds (建設公債), issued in 1940 during the Second Sino-Japanese War. Denominated in US dollars to attract overseas Chinese investors and foreign capital, these bonds were issued by the Nationalist Government (Kuomintang) to finance wartime infrastructure and military operations while much of coastal China was under Japanese occupation. Dollar-denominated Chinese bonds of this era reflect the internationalization of Nationalist government borrowing.',
  { type: 'Bond', subjectCountry: 'China', issuingCountry: 'China', creator: 'Republic of China, Ministry of Finance', issueDate: '1940-01-01', currency: 'USD', language: 'Chinese', numberPages: 1, period: '20th Century', notes: 'Republic of China Year 29 Construction Gold Bonds (建設公債), $5 USD bond, 1940' });

setDoc(319, 'Gouvernement Général de l\'Indo-Chine, 3.5% Obligation',
  'This engraved bond certificate represents a 3.5% obligation of 500 Francs, No. 153,376, of the Gouvernement Général de l\'Indo-Chine (General Government of French Indochina), issued on August 5, 1902. France issued obligations backed by the Indochina colonial government to finance railway construction, port development, and administrative infrastructure across Vietnam, Cambodia, and Laos. These bonds were secured by colonial revenues including customs duties and monopolies on opium, salt, and alcohol.',
  { type: 'Bond', subjectCountry: 'Vietnam', issuingCountry: 'France', creator: 'Gouvernement Général de l\'Indo-Chine', issueDate: '1902-08-05', currency: 'FRF', language: 'French', numberPages: 1, period: '20th Century', notes: 'Gouvernement Général de l\'Indo-Chine, 500 Franc 3.5% obligation No. 153,376, August 5, 1902' });

setDoc(320, 'Chinese Government Loan, £10 Bond at 4½%',
  'This bond certificate represents a £10 bond at 4½% from the Chinese Government Loan of £6,866,046 10/10, signed by Li Sihao as Minister of Finance. This was one of the reorganization or consolidation loans issued by Republican China to restructure the foreign debt obligations inherited from the Qing Dynasty. British and European banks organized syndicates to place Chinese government bonds with international investors; such bonds were typically secured by Chinese customs revenues administered by the foreign-controlled Imperial Maritime Customs Service.',
  { type: 'Bond', subjectCountry: 'China', issuingCountry: 'China', creator: 'Republic of China, Ministry of Finance', issueDate: '1920-01-01', currency: 'GBP', language: 'English', numberPages: 1, period: '20th Century', notes: 'Chinese Government Loan £6,866,046 10/10, £10 Bond at 4½%, signed by Li Sihao, Minister of Finance' });

setDoc(321, 'Chinese Rent Receipt (收租票), Guangxu Period',
  'This handwritten Chinese rent receipt (收租票) dates from the Guangxu period of the Qing Dynasty (1875–1908), issued by Bao Shan Hall (寶善堂 or similar). It records a rental payment for land or property using traditional Chinese financial accounting conventions and classical Chinese script. Such documents are primary evidence of land tenure and rental practices in late imperial China—a period of significant economic transformation as China confronted Western commercial and legal models while maintaining traditional agrarian financial customs.',
  { type: 'Receipt', subjectCountry: 'China', issuingCountry: 'China', creator: 'Bao Shan Hall', issueDate: '1890-01-01', currency: 'CNY', language: 'Chinese', numberPages: 1, period: '19th Century', notes: 'Chinese rent receipt (收租票), Guangxu period (Qing Dynasty), ca. 1875-1908' });

setDoc(322, '遂億換揭貸倉有限公司 (Sui Yi Mortgage Warehouse Co.), Share Certificate',
  'This Chinese share certificate (股份憑票) is from the 遂億換揭貸倉有限公司 (Sui Yi Huan Jie Dai Cang Youxian Gongsi / Sui Yi Mortgage Warehouse Limited Company), issued in Guangxu Year 29 (1903). The company combined mortgage lending (揭貸) with warehouse operations (倉), issuing receipts against stored goods as a form of commodity-backed credit—a hybrid of traditional Chinese pawnshop and Western joint-stock company forms. This represents the institutional experimentation of the late Qing reform era, when Chinese entrepreneurs adapted Western corporate structures to indigenous financial practices.',
  { type: 'Stock Certificate', subjectCountry: 'China', issuingCountry: 'China', creator: '遂億換揭貸倉有限公司', issueDate: '1903-01-01', currency: 'CNY', language: 'Chinese', numberPages: 1, period: '20th Century', notes: '遂億換揭貸倉有限公司股份憑票 (Sui Yi Mortgage Warehouse Co.), Guangxu Year 29 (1903)' });

setDoc(323, '筒碧鐵路公司 (Tong Bi Railway Company), Share Certificate',
  'This Chinese railway share certificate (股票) is from the 筒碧鐵路公司 (Tong Bi Railway Company), issued in Republic of China Year 20 (1931), and features a printed map showing the company\'s railway route. The inclusion of a route map served both as corporate branding and as documentary evidence of the planned physical infrastructure. The Tong Bi Railway was a provincial or private railway project of the Nanjing Decade (1927–1937), reflecting the Nationalist government\'s efforts to expand China\'s railway network beyond the coastal treaty port regions.',
  { type: 'Stock Certificate', subjectCountry: 'China', issuingCountry: 'China', creator: '筒碧鐵路公司', issueDate: '1931-01-01', currency: 'CNY', language: 'Chinese', numberPages: 1, period: '20th Century', notes: '筒碧鐵路公司股票 (Tong Bi Railway Company), Republic of China Year 20 (1931), with map' });

setDoc(324, '日本勤業銀行 (Japan Industrial Bank) Discount Bond, 3,000 Yen',
  'This Japanese bond certificate is a 第十面割引勤業債券 (10th Issue Discount Industrial Bond) of the 日本勤業銀行 (Nippon Kangyo Ginko / Japan Industrial Bank), denomination 3,000 yen, with a sale period of July 15 to July 18. The Japan Industrial Bank, established in 1897, was a government-chartered special bank that issued debentures (勤業債券) to fund agricultural improvement, land development, and rural credit throughout the empire. Discount bonds were sold below face value and redeemed at par, with no periodic coupon payments.',
  { type: 'Bond', subjectCountry: 'Japan', issuingCountry: 'Japan', creator: '日本勤業銀行 (Japan Industrial Bank)', issueDate: '1900-07-15', currency: 'JPY', language: 'Japanese', numberPages: 1, period: '20th Century', notes: '日本勤業銀行 第十面割引勤業債券 (Japan Industrial Bank 10th Discount Bond), 3,000 yen' });

setDoc(325, 'Republic of China Third Lottery Bond (第叁次有獎公債), 1927',
  'This Chinese government bond is the Third Prize Lottery Bond (第叁次有獎公債) issued by the National Government Ministry of Finance (國民政府財政部), Republic of China Year 16 (August 1, 1927), denomination 5 yuan, No. 0675928. The bond features a portrait of Sun Yat-sen, founding father of the Republic. Lottery bonds—public debt instruments with periodic prize drawings instead of ordinary interest—were popular Nationalist fundraising devices; this third issue was among the earliest such bonds issued by the newly established Nanjing government after its 1927 victory in the Northern Expedition.',
  { type: 'Bond', subjectCountry: 'China', issuingCountry: 'China', creator: '國民政府財政部 (Republic of China Ministry of Finance)', issueDate: '1927-08-01', currency: 'CNY', language: 'Chinese', numberPages: 1, period: '20th Century', notes: 'Republic of China Third Lottery Bond (第叁次有獎公債), 5 yuan No. 0675928, August 1, 1927, with Sun Yat-sen portrait' });

setDoc(326, '裕真地產有限股份公司 (Keh-Yul Land Co.), Share Certificate',
  'This Chinese share certificate (股份憑票) is from 裕真地產有限股份公司 (Yuzheng Didi Youxian Gufen Gongsi / Keh-Yul Land Co.), No. 012430, representing 500 shares, signed by board members. The company is a real estate enterprise organized under the limited joint-stock form in Republican-era China. Urban real estate companies of this type were active in Chinese treaty ports and major cities, developing commercial and residential property amid the rapid urbanization of the 1920s–1930s.',
  { type: 'Stock Certificate', subjectCountry: 'China', issuingCountry: 'China', creator: '裕真地產有限股份公司', issueDate: '1930-01-01', currency: 'CNY', language: 'Chinese', numberPages: 1, period: '20th Century', notes: '裕真地產有限股份公司 (Keh-Yul Land Co.), No. 012430, 500 shares' });

setDoc(327, 'Cinco Por Ciento Español (Spanish 5% Perpetual Rente), Primera Serie',
  'This Spanish government bond certificate represents a share of the Cinco Por Ciento Español (Spanish 5% Perpetual Rente), Primera Serie, issued at Oñate on February 6, 1860, with interest payable in Madrid, London, Paris, Brussels, and Turin. The certificate has attached coupon strips for the periodic interest payments. The Spanish 5% perpetual rente was a consolidated national debt instrument traded on multiple European exchanges to attract international investors; it was serviced at financial centers across Europe, reflecting Spain\'s integration into international capital markets in the mid-nineteenth century.',
  { type: 'Bond', subjectCountry: 'Spain', issuingCountry: 'Spain', creator: 'Government of Spain', issueDate: '1860-02-06', currency: 'ESP', language: 'Spanish', numberPages: 1, period: '19th Century', notes: 'Cinco Por Ciento Español (5% Spanish Perpetual Rente), Primera Serie, Oñate, February 6, 1860, payable in Madrid/London/Paris/Brussels/Turin, with coupon strips' });

setDoc(328, 'Citrus Belt Land Company, Stock Certificate',
  'This stock certificate represents 50 shares in the Citrus Belt Land Company of California, Certificate No. 426, issued to S.J. White on January 26, 1911. The Citrus Belt Land Company operated in Southern California\'s "Citrus Belt"—the citrus-growing regions of San Bernardino and Riverside counties, developed through irrigation projects and the Southern Pacific and Santa Fe railroads in the late nineteenth century. Land companies of this type subdivided and marketed agricultural land to settlers as part of the boosterism that transformed Southern California from semi-arid ranchland into densely settled citrus groves and suburbs.',
  { type: 'Stock Certificate', subjectCountry: 'United States', issuingCountry: 'United States', creator: 'Citrus Belt Land Company', issueDate: '1911-01-26', currency: 'USD', language: 'English', numberPages: 1, period: '20th Century', notes: 'Citrus Belt Land Company, California, 50 shares, No. 426, S.J. White, January 26, 1911' });

setDoc(329, 'Clădirea Românească (Romanian Building Society), Share Certificate',
  'This ornately printed share certificate represents 50 nominative shares (acțiuni nominative) of 500 lei each—totaling lei 25,000—in Clădirea Românească S.A. (Romanian Building Society), Emission VI-a, Nos. 1095201–1095250, Capital Social 600,000,000 lei, issued in Bucharest in January 1946. Clădirea Românească was a Romanian construction and real estate company. Its January 1946 issuance falls in the turbulent early postwar period, as Romania transitioned from wartime German occupation toward Soviet-aligned communist rule; such private companies were nationalized under the communist government that consolidated power in 1947–1948.',
  { type: 'Stock Certificate', subjectCountry: 'Romania', issuingCountry: 'Romania', creator: 'Clădirea Românească S.A.', issueDate: '1946-01-01', currency: 'ROL', language: 'Romanian', numberPages: 1, period: '20th Century', notes: 'Clădirea Românească S.A., 50 nominative shares of 500 lei each (Nos. 1095201-1095250), Emission VI-a, Capital 600,000,000 lei, Bucharest, January 1946' });

setDoc(330, 'Colombian India-Rubber Exploration Company Limited, Share Warrant',
  'This share warrant certifies the bearer\'s ownership of 25 shares (Nos. 188,476–188,500) in the Colombian India-Rubber Exploration Company Limited, issued on March 4, 1907. The company was a British enterprise formed to exploit natural rubber concessions in Colombia during the Edwardian rubber boom, when wild rubber from South America commanded premium prices before East Asian plantation rubber came to dominate the market. The company participated in the global scramble for tropical rubber resources that also drove the atrocities of the Congo Free State and the Peruvian Amazon.',
  { type: 'Share Warrant', subjectCountry: 'Colombia', issuingCountry: 'United Kingdom', creator: 'Colombian India-Rubber Exploration Company Limited', issueDate: '1907-03-04', currency: 'GBP', language: 'English', numberPages: 1, period: '20th Century', notes: 'Colombian India-Rubber Exploration Company Limited, Share Warrant 25 shares (188,476-188,500), March 4, 1907' });

setDoc(331, 'Compagnie d\'Inguaran (Mexique), Part de Bénéfices',
  'This profit-sharing certificate, No. 07,399, is a Part de Bénéfices (profit participation share) of the Compagnie d\'Inguaran, a French company with interests in Mexico, issued in Paris on January 15, 1898. The Compagnie d\'Inguaran held concessions in the Inguaran district of Michoacán—historically important for copper and silver mining—during the Porfiriato era, when the Mexican government under Porfirio Díaz actively courted foreign capital. Parts de bénéfices were a French corporate form entitling holders to a share of profits without a fixed par value or voting rights.',
  { type: 'Profit Share', subjectCountry: 'Mexico', issuingCountry: 'France', creator: 'Compagnie d\'Inguaran', issueDate: '1898-01-15', currency: 'FRF', language: 'French', numberPages: 1, period: '19th Century', notes: 'Compagnie d\'Inguaran (Mexique), Part de Bénéfices No. 07,399, Paris, January 15, 1898' });

setDoc(332, 'Compagnie Occidentale de Madagascar, 100 Franc Share',
  'This engraved share certificate represents one action of 100 Francs in the Compagnie Occidentale de Madagascar (Western Madagascar Company), No. 16,102, issued in Paris in the early twentieth century. The Compagnie Occidentale de Madagascar was a French colonial company operating in western Madagascar, France\'s large island possession in the Indian Ocean (colonized 1896). Such companies typically pursued agricultural production, stock ranching, or mineral extraction; western Madagascar\'s grasslands were suited to cattle ranching and the export of hides and other livestock products.',
  { type: 'Stock Certificate', subjectCountry: 'Madagascar', issuingCountry: 'France', creator: 'Compagnie Occidentale de Madagascar', issueDate: '1910-01-01', currency: 'FRF', language: 'French', numberPages: 1, period: '20th Century', notes: 'Compagnie Occidentale de Madagascar, Action de Cent Francs No. 16,102, Paris, ca. 1910s' });

setDoc(333, 'Compañía Salitrera de Tarapacá y Antofagasta, 10 Shares',
  'This stock certificate represents 10 shares at 50 pesos each (Certificate No. C30095) in the Compañía Salitrera de Tarapacá y Antofagasta, issued in Santiago on April 30, 1937, with a large dividend coupon sheet attached. The company operated in the nitrate (salitre) mining districts of Tarapacá and Antofagasta in northern Chile, territories wrested from Peru and Bolivia in the War of the Pacific (1879–1884). Chilean nitrate, used as fertilizer and in explosives, had dominated world markets before the Haber-Bosch synthetic nitrogen process (1913) began displacing natural nitrates, shrinking the industry dramatically by the 1930s.',
  { type: 'Stock Certificate', subjectCountry: 'Chile', issuingCountry: 'Chile', creator: 'Compañía Salitrera de Tarapacá y Antofagasta', issueDate: '1937-04-30', currency: 'CLP', language: 'Spanish', numberPages: 1, period: '20th Century', notes: 'Compañía Salitrera de Tarapacá y Antofagasta, 10 shares at 50 pesos, No. C30095, Santiago, April 30, 1937, with coupon sheet' });

setDoc(334, 'Companhia de Mossamedes, Share Warrant to Bearer for Five Shares',
  'This share warrant certifies the bearer\'s ownership of five fully paid shares (Nos. 2,272,126–2,272,130), each of Escudos 4.50, in the Companhia de Mossamedes (Moçâmedes Company), a Portuguese Société Anonyme à Responsabilité Limitée with capital of 13,995,000 Escudos divided into 3,110,000 shares. The company operated in Moçâmedes (now Namibe), a port city in southern Angola, then Portuguese West Africa. The warrant is trilingual (Portuguese, French, and English), reflecting the company\'s use of international capital markets to fund its colonial operations.',
  { type: 'Share Warrant', subjectCountry: 'Angola', issuingCountry: 'Portugal', creator: 'Companhia de Mossamedes', issueDate: '1925-01-01', currency: 'PTE', language: 'Portuguese', numberPages: 1, period: '20th Century', notes: 'Companhia de Mossamedes, Share Warrant for 5 shares (Nos. 2272126-2272130), Esc. 4.50 each, Capital 13,995,000 Esc.' });

setDoc(335, 'Compañía Anónima Minas de Guanandi, 500 Pesos Share',
  'This stock certificate represents one share of 500 Pesos ¾ Oro in the Compañía Anónima Minas de Guanandi, No. 2980, Capital 2,000,000 Pesos Oro (4,000 shares). Issued in Montevideo, 1887. The company was authorized by Decree of October 19, 1887 of the Government of Uruguay to exploit 50 mineral concessions in the municipality of Pocone, province of Matto Grosso (Brazil), granted by Brazilian Imperial Decree No. 9239 of June 28, 1884. This certificate illustrates the cross-border corporate investment of the late nineteenth century, with a Uruguayan company holding concessions to mine precious metals in Brazilian territory.',
  { type: 'Stock Certificate', subjectCountry: 'Brazil', issuingCountry: 'Uruguay', creator: 'Compañía Anónima Minas de Guanandi', issueDate: '1887-01-01', currency: 'UYU', language: 'Spanish', numberPages: 1, period: '19th Century', notes: 'Compañía Anónima Minas de Guanandi, 1 share 500 Pesos ¾ Oro, No. 2980, Capital 2,000,000 Pesos, Montevideo 1887; concessions in Mato Grosso, Brazil' });

const newWs = xlsx.utils.aoa_to_sheet(data);
newWs['!cols'] = ws['!cols'];
wb.Sheets['Documents'] = newWs;
xlsx.writeFile(wb, filePath);
console.log('Done. Updated rows 286-335 (24 Geo II Acts continuation + individual documents).');
