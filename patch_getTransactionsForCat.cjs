const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const oldFunc = `function getTransactionsForCategorization() {
  const sheet = getTable('Transactions');
  let tz = Session.getScriptTimeZone();
  try { tz = sheet.getParent().getSpreadsheetTimeZone(); } catch(e) {}
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let dateStr = row[5];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, tz, "yyyy-MM-dd");
    }
    rows.push({
      rowIndex: i + 1,
      origem: row[0] || 'Importado',
      descricao: String(row[1] || '').trim(),
      categoria: String(row[2] || '').trim(),
      tipo: row[3] || 'Saída',
      valor: parseValor(row[4]),
      data: dateStr || ''
    });
  }
  return rows;
}`;

const newFunc = `function getTransactionsForCategorization() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let displayDate = displayData[i][5];
    let dateStr = displayDate;
    if (dateStr.includes('/')) {
      let parts = dateStr.split('/');
      if (parts.length === 3) {
        dateStr = \`\${parts[2]}-\${parts[1]}-\${parts[0]}\`;
      }
    }

    rows.push({
      rowIndex: i + 1,
      origem: String(row[0] || 'Importado'),
      descricao: String(row[1] || '').trim(),
      categoria: String(row[2] || '').trim(),
      tipo: String(row[3] || 'Saída'),
      valor: parseValor(row[4]),
      data: dateStr || ''
    });
  }
  return rows;
}`;

if (gs.includes(oldFunc)) {
  gs = gs.replace(oldFunc, newFunc);
  fs.writeFileSync('gas-project/Transactions.gs', gs);
  console.log('Patched getTransactionsForCategorization');
} else {
  console.log('Could not find oldFunc exactly');
}
