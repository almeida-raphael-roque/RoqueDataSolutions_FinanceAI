const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const regex = /function getPlanningData\(year\) \{\s*const sheet = getTable\('Transactions'\);\s*const data = sheet\.getDataRange\(\)\.getValues\(\);/;
gs = gs.replace(regex, `function getPlanningData(year) {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();`);

gs = gs.replace(/const d = parseFlexDate\(row\[5\]\);/g, 'const d = parseFlexDate(displayData[i][5]);');

fs.writeFileSync('gas-project/Transactions.gs', gs);
console.log('Patched getPlanningData');
