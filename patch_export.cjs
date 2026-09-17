const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const regex = /function exportToCsv\(\) \{\s*const sheet = getTable\('Transactions'\);\s*const data = sheet\.getDataRange\(\)\.getValues\(\);/;
gs = gs.replace(regex, `function exportToCsv() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getDisplayValues();`);

// It currently checks `if (cell instanceof Date)`. Since we changed to getDisplayValues(), 
// it will never be a Date! It will just be the display string!
// We can just leave `if (cell instanceof Date)` as a no-op fallback, it won't hurt, it will just go to `else { str = String(cell); }`.

fs.writeFileSync('gas-project/Transactions.gs', gs);
console.log('Patched exportToCsv');
