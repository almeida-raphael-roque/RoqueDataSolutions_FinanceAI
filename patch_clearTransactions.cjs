const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const regex = /function clearTransactions\(yearStr, monthStr\) \{\s*backupForUndo\(\);\s*const sheet = getTable\('Transactions'\);\s*const data = sheet\.getDataRange\(\)\.getValues\(\);/;
gs = gs.replace(regex, `function clearTransactions(yearStr, monthStr) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();`);

gs = gs.replace(/if \(!isMatch\(row\[5\], yearStr, monthStr\)\)/g, 'if (!isMatch(displayData[i][5], yearStr, monthStr))');

fs.writeFileSync('gas-project/Transactions.gs', gs);
console.log('Patched clearTransactions');
