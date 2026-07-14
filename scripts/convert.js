const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const rootDir = path.join(__dirname, '..');
const dataDir = path.join(rootDir, 'data');

if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
}

function findLatestExcelFile() {
  const candidateDirs = [rootDir, dataDir];
  let found = [];
  for (const dir of candidateDirs) {
    if (!fs.existsSync(dir)) continue;
    const files = fs.readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith('.xlsx'))
    .map((f) => path.join(dir, f));
    found = found.concat(files);
  }
  if (found.length === 0) return null;
  found.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
  return found[0];
}

const inputFile = findLatestExcelFile();
if (!inputFile) {
  console.error('Geen .xlsx bestand gevonden in de repository root of in /data.');
  process.exit(1);
}

console.log('Inlezen van: ' + inputFile);
const workbook = XLSX.readFile(inputFile);

const sheetName = workbook.SheetNames.find(
  (name) => name.trim().toLowerCase() === 'alle woningen'
  ) || workbook.SheetNames[0];

const sheet = workbook.Sheets[sheetName];
const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });

const EUR = '\u20AC';
const M2 = '\u00B2';

const columnMap = {
  'Funda ID': 'id',
  'Funda URL': 'url',
  'Projectnaam': 'project',
  'Adres': 'adres',
  'Bouwnummer': 'bouwnr',
  'Postcode': 'postcode',
  'Plaats': 'plaats',
  'Gemeente': 'gemeente',
  'Provincie': 'provincie',
  'Regio': 'regio',
  'Status': 'status',
  'Datum plaatsing': 'datumPlaatsing',
  'Looptijd (dagen)': 'looptijd',
  ['Vraagprijs (' + EUR + ')']: 'vp',
  'Kostenlabel': 'betaalbaarheid',
  ['GBO (m' + M2 + ')']: 'gbo',
  ['VON-prijs per m' + M2 + ' (' + EUR + ')']: 'ppm',
  ['Perceeloppervlakte (m' + M2 + ')']: 'perceel',
  'Energielabel': 'energie',
  'Aantal kamers': 'kamers',
  'Aantal slaapkamers': 'slaapkamers',
  'Woningcategorie': 'woningcategorie',
  'Woningtype': 'type',
  'Appartement categorie': 'appartementCategorie',
  'Parkeerplaats': 'parkeerplaats',
  'Type parkeerplaats': 'typeParkeerplaats',
  'Bouwjaar': 'bouwjaar',
  'Prijs op aanvraag': 'prijsOpAanvraag',
  'Outlier flag': 'outlier',
  'GemeenteOK': 'gemeenteOk',
};

const records = rows.map((row) => {
  const rec = {};
  for (const excelCol of Object.keys(columnMap)) {
    const key = columnMap[excelCol];
    if (row[excelCol] !== undefined) {
      rec[key] = row[excelCol];
    }
  }
  if (
    rec.woningcategorie &&
    String(rec.woningcategorie).toLowerCase().indexOf('appartement') !== -1 &&
    rec.appartementCategorie
    ) {
    rec.type = rec.appartementCategorie;
  }
  return rec;
});

const output = {
  generatedAt: new Date().toISOString(),
  sourceFile: path.basename(inputFile),
  sheet: sheetName,
  count: records.length,
  records: records,
};

fs.writeFileSync(
  path.join(dataDir, 'data.json'),
  JSON.stringify(output, null, 2)
  );

console.log(records.length + ' woningen geschreven naar data/data.json');
