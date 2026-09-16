const fs = require('fs');
let code = fs.readFileSync('gas-project/Scripts.html', 'utf8');

const oldFilter = `$('#filter-month, #filter-year').change(function() {
      // Depending on active page, reload data
      if (!$('#page-dashboard').hasClass('hidden-page')) {
        loadDashboard();
      }
      if (!$('#page-transactions').hasClass('hidden-page')) {
        loadTransactions();
      }
      if (!$('#page-spending-control').hasClass('hidden-page') && typeof loadSpendingControl === 'function') {
        loadSpendingControl();
      }
    });`;

const newFilter = `$('#filter-month, #filter-year').change(function() {
      invalidateCache();
    });`;

code = code.replace(oldFilter, newFilter);
fs.writeFileSync('gas-project/Scripts.html', code);
console.log("Filter patched");
