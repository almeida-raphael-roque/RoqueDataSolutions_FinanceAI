const fs = require('fs');
let code = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

code = code.replace(
  'function getMacroCategory(categoria, descricao, catMacroMap, typeof catMacroMap !== "undefined" ? catMacroMap : null) {',
  'function getMacroCategory(categoria, descricao, catMacroMap) {'
);

fs.writeFileSync('gas-project/Transactions.gs', code);
