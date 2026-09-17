const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

// We will inject `const displayData = sheet.getDataRange().getDisplayValues();`
// And change `let dateStr = row[5];` to `let dateStr = displayData[i][5];`
// And remove `if (dateStr instanceof Date)` block since displayData is always a string!

function patchFunction(funcName, dateColIndex) {
  const regex = new RegExp(`(function ${funcName}\\([^{]*\\)\\s*\\{[\\s\\S]*?const data = sheet\\.getDataRange\\(\\)\\.getValues\\(\\);)`);
  gs = gs.replace(regex, `$1\n  const displayData = sheet.getDataRange().getDisplayValues();`);
  
  // Now replace the extraction
  // Usually it is `const row = data[i];` and then `row[5]` or `dataVal = row[5]`
  // We want to replace `row[${dateColIndex}]` with `displayData[i][${dateColIndex}]` where it assigns to date
  
  // For getTransactions:
  if (funcName === 'getTransactions') {
    gs = gs.replace(/if \\(isMatch\\(row\\[5\\], year, month\\)\\) \\{\\s*\\/\\/ Formatação básica para retornar como string consistente\\s*let dateStr = row\\[5\\];\\s*if \\(dateStr instanceof Date\\) \\{\\s*dateStr = Utilities\\.formatDate\\(dateStr, tz, "yyyy-MM-dd"\\);\\s*\\}/,
      `let rawDateStr = displayData[i][5];\n    if (isMatch(rawDateStr, year, month)) {\n      let dateStr = rawDateStr;`);
  }
}

// Let's just do it manually with string replacements since regex is tricky for this block.
