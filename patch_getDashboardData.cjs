const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const regex = /function getDashboardData\(year = 'all', month = 'all'\) \{\s*const sheet = getTable\('Transactions'\);\s*const data = sheet\.getDataRange\(\)\.getValues\(\);/;
gs = gs.replace(regex, `function getDashboardData(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();`);

// replace row[5] inside loop:
// const dataVal = row[5];
// to:
// const dataVal = displayData[i][5];

gs = gs.replace(/const dataVal = row\[5\];/g, 'const dataVal = displayData[i][5];');

fs.writeFileSync('gas-project/Transactions.gs', gs);
console.log('Patched getDashboardData');
