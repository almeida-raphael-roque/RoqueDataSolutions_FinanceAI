const fs = require('fs');
let code = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

if (code.includes('catMacroMap')) {
    console.log("Success! catMacroMap found.");
} else {
    console.log("Failed to patch Transactions.gs");
}
