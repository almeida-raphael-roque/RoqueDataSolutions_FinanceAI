/**
 * Funções relacionadas à importação de transações, regras do usuário e categorias
 */

function ensureUserRulesSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName("UserRules");
  if (!sheet) {
    sheet = ss.insertSheet("UserRules");
    sheet.appendRow(["Padrão (Normalizado)", "Categoria", "Estabelecimento"]);
    sheet.getRange("A1:C1").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getUserRules() {
  const sheet = ensureUserRulesSheet();
  const data = sheet.getDataRange().getValues();
  const rules = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    rules.push({
      id: i,
      pattern: String(data[i][0]).toUpperCase(),
      categoria: data[i][1],
      clean: data[i][2] || String(data[i][0]),
      estabelecimento: data[i][2] || ''
    });
  }
  return rules;
}

function saveUserRulesBatch(rules) {
  if (!rules || rules.length === 0) return true;
  const sheet = ensureUserRulesSheet();
  const existing = getUserRules().map(r => r.pattern.toUpperCase());
  
  const rowsToAppend = [];
  rules.forEach(r => {
    const pat = String(r.pattern || '').trim().toUpperCase();
    if (pat && !existing.includes(pat)) {
      rowsToAppend.push([pat, r.categoria, r.clean || r.estabelecimento || pat]);
      existing.push(pat);
    }
  });
  
  if (rowsToAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, 3).setValues(rowsToAppend);
  }
  return true;
}

function saveImportedTransactions(transactions) {
  if (typeof backupForUndo === 'function') backupForUndo();
  const sheet = getTable('Transactions');
  const rows = transactions.map(t => [t.origem || 'Importado', t.descricao, t.categoria, t.tipo, t.valor, t.data]);
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
  return true;
}

function saveAllUserRules(rules) {
  if (!rules) return { success: true, updatedCount: 0, totalRules: 0 };
  const sheet = ensureUserRulesSheet();
  sheet.clear();
  sheet.appendRow(["Padrão (Normalizado)", "Categoria", "Estabelecimento"]);
  sheet.getRange("A1:C1").setFontWeight("bold");
  sheet.setFrozenRows(1);
  
  if (rules.length > 0) {
    const rows = rules.map(r => [
      String(r.pattern || '').trim().toUpperCase(),
      r.categoria || '',
      r.clean || r.estabelecimento || r.pattern || ''
    ]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  SpreadsheetApp.flush();

  // Aplica imediatamente as novas regras nas transações da planilha
  let txUpdatedCount = 0;
  try {
    if (typeof applyUserRulesToTransactionsBackend === 'function') {
      const res = applyUserRulesToTransactionsBackend(rules);
      if (res && res.updatedCount) txUpdatedCount = res.updatedCount;
    }
  } catch(e) {
    Logger.log("Aviso ao aplicar regras às transações na planilha: " + e);
  }

  return { success: true, updatedCount: txUpdatedCount, totalRules: (rules || []).length };
}


