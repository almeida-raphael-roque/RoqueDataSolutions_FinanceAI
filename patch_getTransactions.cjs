const fs = require('fs');
let gs = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

const oldFunc = `function getTransactions(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  let tz = Session.getScriptTimeZone();
  try { tz = sheet.getParent().getSpreadsheetTimeZone(); } catch(e) {}
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catNameMap = {};
  cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);
  
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    if (isMatch(row[5], year, month)) {
      // Formatação básica para retornar como string consistente
      let dateStr = row[5];
      if (dateStr instanceof Date) {
        dateStr = Utilities.formatDate(dateStr, tz, "yyyy-MM-dd");
      }
      let tipo = String(row[3] || '').trim();
      const tipoL = tipo.toLowerCase();
      if (tipoL === 'saída' || tipoL === 'saida' || tipoL === 'despesa') tipo = 'Saída';
      else if (tipoL === 'entrada' || tipoL === 'receita') tipo = 'Entrada';
      let categoria = String(row[2] || '').trim() || 'Outros';
      if (catNameMap[categoria.toLowerCase()]) {
        categoria = catNameMap[categoria.toLowerCase()];
      }
      rows.push({
        origem: String(row[0] || '').trim(),
        descricao: String(row[1] || '').trim(),
        categoria: categoria,
        tipo: tipo,
        valor: parseValor(row[4]),
        data: dateStr
      });
    }
  }
  return rows;
}`;

const newFunc = `function getTransactions(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return []; 
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catNameMap = {};
  cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);
  
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    const displayDate = displayData[i][5];
    if (isMatch(displayDate, year, month)) {
      let dateStr = displayDate;
      // If it is in dd/mm/yyyy format from display, convert to yyyy-mm-dd so frontend parses it consistently
      if (dateStr.includes('/')) {
        let parts = dateStr.split('/');
        if (parts.length === 3) {
          dateStr = \`\${parts[2]}-\${parts[1]}-\${parts[0]}\`;
        }
      }

      let tipo = String(row[3] || '').trim();
      const tipoL = tipo.toLowerCase();
      if (tipoL === 'saída' || tipoL === 'saida' || tipoL === 'despesa') tipo = 'Saída';
      else if (tipoL === 'entrada' || tipoL === 'receita') tipo = 'Entrada';
      let categoria = String(row[2] || '').trim() || 'Outros';
      if (catNameMap[categoria.toLowerCase()]) {
        categoria = catNameMap[categoria.toLowerCase()];
      }
      rows.push({
        origem: String(row[0] || '').trim(),
        descricao: String(row[1] || '').trim(),
        categoria: categoria,
        tipo: tipo,
        valor: parseValor(row[4]),
        data: dateStr
      });
    }
  }
  return rows;
}`;

if (gs.includes(oldFunc)) {
  gs = gs.replace(oldFunc, newFunc);
  fs.writeFileSync('gas-project/Transactions.gs', gs);
  console.log('Patched getTransactions');
} else {
  console.log('Could not find oldFunc exactly');
}
