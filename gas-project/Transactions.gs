function saveTransaction(origem, descricao, categoria, tipo, valor, data) {
  backupForUndo();
  const sheet = getTable('Transactions');
  sheet.appendRow([origem, descricao, categoria, tipo, valor, data]);
  SpreadsheetApp.flush();
  
  return true;
}

function saveBatchTransactions(items) {
  if (!items || !items.length) return true;
  backupForUndo();
  const sheet = getTable('Transactions');
  const rows = items.map(function(t) {
    return [t.origem, t.descricao, t.categoria, t.tipo, t.valor, t.data];
  });
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length).setValues(rows);
  SpreadsheetApp.flush();
  
  return true;
}


function clearTransactions(yearStr, monthStr) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return true; // Somente cabeçalho
  
  const header = data[0];
  const toKeep = [header];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!isMatch(displayData[i][5], yearStr, monthStr)) {
      toKeep.push(row);
    }
  }
  
  sheet.clear();
  if (toKeep.length > 0) {
    sheet.getRange(1, 1, toKeep.length, toKeep[0].length).setValues(toKeep);
  }
  
  return true;
}

function backupForUndo() {
  if (typeof SpreadsheetApp === 'undefined') return;
  const ss = getSpreadsheet();
  const source = getTable('Transactions');
  let backup = ss.getSheetByName('_Undo_Backup');
  if (!backup) {
    backup = ss.insertSheet('_Undo_Backup');
    backup.hideSheet();
  }
  backup.clear();
  const data = source.getDataRange().getValues();
  if (data.length > 0) {
    backup.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
}

function undoLastAction() {
  if (typeof SpreadsheetApp === 'undefined') return;
  const ss = getSpreadsheet();
  const backup = ss.getSheetByName('_Undo_Backup');
  if (!backup) throw new Error("Nenhuma ação anterior para desfazer.");
  
  const target = getTable('Transactions');
  const backupData = backup.getDataRange().getValues();
  const currentData = target.getDataRange().getValues();
  
  if (backupData.length === 0) throw new Error("Nenhuma ação anterior para desfazer.");
  
  target.clear();
  target.getRange(1, 1, backupData.length, backupData[0].length).setValues(backupData);
  
  backup.clear();
  if (currentData.length > 0) {
    backup.getRange(1, 1, currentData.length, currentData[0].length).setValues(currentData);
  }
  
  return true;
}

function parseFlexDate(dateVal) {
  if (!dateVal) return new Date(NaN);
  if (dateVal instanceof Date) return dateVal;
  
  const strVal = dateVal.toString().trim();
  
  let parts = strVal.split('T')[0].split('-');
  if (parts.length === 3) {
    return new Date(parts[0], parseInt(parts[1])-1, parts[2]);
  }
  
  let ptParts = strVal.split('/');
  if (ptParts.length === 3) {
    return new Date(ptParts[2], parseInt(ptParts[1])-1, ptParts[0]);
  }
  
  return new Date(strVal);
}

function parseValor(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  let str = String(val).replace(/[R\$\s]/gi, '').trim();
  if (str.includes(',') && str.includes('.')) {
    str = str.replace(/\./g, '').replace(',', '.');
  } else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  const parsed = parseFloat(str);
  return isNaN(parsed) ? 0 : parsed;
}

function isMatch(dateVal, yearStr, monthStr) {
  if (yearStr === 'all' && monthStr === 'all') return true;
  if (!dateVal) return false;
  
  const d = parseFlexDate(dateVal);
  
  if (isNaN(d.getTime())) return false; // Ignore invalids if filtering
  
  const y = d.getFullYear().toString();
  const m = (d.getMonth() + 1).toString();
  
  const matchY = (String(yearStr) === 'all' || String(yearStr) === y);
  const matchM = (String(monthStr) === 'all' || String(monthStr) === m);
  
  return matchY && matchM;
}

function getTransactions(year = 'all', month = 'all') {
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
          dateStr = `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
      }

      let tipo = String(row[3] || '').trim();
      const tipoL = tipo.toLowerCase();
      if (tipoL === 'saída' || tipoL === 'saida' || tipoL === 'despesa') tipo = 'Saída';
      else if (tipoL === 'entrada' || tipoL === 'receita') tipo = 'Entrada';
      let categoria = String(row[2] || '').trim();
      if (categoria && catNameMap[categoria.toLowerCase()]) {
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
}

function getMacroCategory(categoria, descricao, catMacroMap) {
  const catLower = String(categoria || '').toLowerCase().trim();
  const descLower = String(descricao || '').toLowerCase().trim();
  
  // 1. Check Global Categories (Single Source of Truth)
  if (catMacroMap && catMacroMap[catLower]) {
      return catMacroMap[catLower];
  }
  
  // 2. Fallback heuristics for unmapped categories
  if (catLower.includes('moradia') || descLower.includes('aluguel') || descLower.includes('internet') || descLower.includes('condomínio') || catLower.includes('educação') || catLower.includes('telefone') || catLower.includes('celular') || descLower.includes('vivo') || descLower.includes('claro') || descLower.includes('tim')) return 'Fixas';
  if (descLower.includes('energia') || descLower.includes('água') || descLower.includes('gás') || descLower.includes('luz') || catLower.includes('saúde') || descLower.includes('mercado') || descLower.includes('supermercado') || descLower.includes('farmácia') || catLower.includes('transporte') || catLower.includes('alimentação')) return 'Variáveis Essenciais';
  if (catLower.includes('lazer') || descLower.includes('restaurante') || descLower.includes('delivery') || descLower.includes('compras') || descLower.includes('passeio') || descLower.includes('cinema') || descLower.includes('netflix') || descLower.includes('uber')) {
    if (descLower.includes('mercado') || descLower.includes('supermercado')) return 'Variáveis Essenciais';
    return 'Discricionárias';
  }
  return 'Discricionárias';
}

const MACRO_COLORS = {
  'Fixas': '#991b1b', // Vermelho escuro
  'Variáveis Essenciais': '#ef4444', // Vermelho intermediário
  'Discricionárias': '#f87171' // Vermelho claro
};

function getDashboardData(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catNameMap = {};
  const catMacroMap = {};
  cats.forEach(c => {
    catNameMap[String(c.name).toLowerCase()] = c.name;
    if (c.profile) {
      if (c.profile.includes('Fixa')) catMacroMap[String(c.name).toLowerCase()] = 'Fixas';
      else if (c.profile.includes('Variável')) catMacroMap[String(c.name).toLowerCase()] = 'Variáveis Essenciais';
      else catMacroMap[String(c.name).toLowerCase()] = 'Discricionárias';
    }
  });
  
  let income = 0;
  let expense = 0;
  const grouped = {};
  const categoryGrouped = {};
  const categoryBreakdown = {};
  const incomeCategoryGrouped = {};
  const macroGrouped = {
    'Fixas': 0,
    'Variáveis Essenciais': 0,
    'Discricionárias': 0
  };
  const rawExpenses = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const origem = String(row[0] || '').trim();
    const desc = String(row[1] || 'Desconhecido').trim();
    let categoria = String(row[2] || 'Outros').trim();
    if (catNameMap[categoria.toLowerCase()]) {
      categoria = catNameMap[categoria.toLowerCase()];
    }
    let tipo = String(row[3] || '').trim();
    const tipoLower = tipo.toLowerCase();
    if (tipoLower === 'saída' || tipoLower === 'saida' || tipoLower === 'despesa') tipo = 'Saída';
    else if (tipoLower === 'entrada' || tipoLower === 'receita') tipo = 'Entrada';
    const valor = parseValor(row[4]);
    const dataVal = displayData[i][5];
    
    // Parse date once
    const d = parseFlexDate(dataVal);
    
    // Save raw expenses for temporal drill-down (all-time)
    if (tipo === 'Saída' && !isNaN(d.getTime())) {
      rawExpenses.push({
        date: d.getTime(), // ms timestamp
        desc: desc,
        value: valor
      });
    }
    
    // Soma para os cards principais respeitando MÊS e ANO
    if (isMatch(dataVal, year, month)) {
      if (tipo === 'Entrada') {
        income += valor;
        if (!incomeCategoryGrouped[categoria]) incomeCategoryGrouped[categoria] = 0;
        incomeCategoryGrouped[categoria] += valor;
      } else if (tipo === 'Saída') {
        expense += valor;
        const macro = getMacroCategory(categoria, desc, typeof catMacroMap !== "undefined" ? catMacroMap : null);
        macroGrouped[macro] += valor;

        if (!categoryGrouped[categoria]) categoryGrouped[categoria] = { value: 0, macros: { 'Fixas': 0, 'Variáveis Essenciais': 0, 'Discricionárias': 0 } };
        categoryGrouped[categoria].value += valor;
        categoryGrouped[categoria].macros[macro] += valor;
        
        if (!categoryBreakdown[categoria]) categoryBreakdown[categoria] = {};
        if (!categoryBreakdown[categoria][desc]) categoryBreakdown[categoria][desc] = { value: 0, macro: macro };
        categoryBreakdown[categoria][desc].value += valor;
      }
    }
    
    // Lógica de agrupamento para o Gráfico Temporal (respeita apenas o ANO para mostrar a evolução temporal correta)
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const mNum = (d.getMonth() + 1).toString().padStart(2, '0');
      
      // Se tiver filtro de ano, aplica ao gráfico também (mas o mês ignora para poder montar a linha do tempo)
      const matchYearForChart = (String(year) === 'all' || String(year) === y.toString());
      
      if (matchYearForChart) {
        const key = `${y}-${mNum}`;
        const label = `${mNum}/${y}`;
        
        if (!grouped[key]) {
          grouped[key] = { label: label, sortKey: key, income: 0, expense: 0 };
        }
        
        if (tipo === 'Entrada') grouped[key].income += valor;
        else if (tipo === 'Saída') grouped[key].expense += valor;
      }
    }
  }
  
  const chartData = Object.values(grouped).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
  
  const categoryChartData = Object.keys(categoryGrouped)
    .map(k => {
      const catData = categoryGrouped[k];
      let dominantMacro = Object.keys(catData.macros).reduce((a, b) => catData.macros[a] > catData.macros[b] ? a : b);
      
      const breakdownObj = categoryBreakdown[k];
      const breakdownArray = Object.keys(breakdownObj)
        .map(desc => ({ label: desc, value: breakdownObj[desc].value, color: MACRO_COLORS[breakdownObj[desc].macro] }))
        .sort((a, b) => b.value - a.value);
        
      return { label: k, value: catData.value, color: MACRO_COLORS[dominantMacro], macro: dominantMacro, breakdown: breakdownArray };
    })
    .sort((a, b) => b.value - a.value);
    
  const incomeCategoryChartData = Object.keys(incomeCategoryGrouped)
    .map(k => ({ label: k, value: incomeCategoryGrouped[k] }))
    .sort((a, b) => b.value - a.value);
    
  const macroChartData = Object.keys(macroGrouped)
    .map(k => ({ label: k, value: macroGrouped[k], color: MACRO_COLORS[k] }))
    .filter(d => d.value > 0)
    .sort((a, b) => b.value - a.value);
  
  return {
    balance: income - expense,
    income: income,
    expense: expense,
    chartData: chartData,
    categoryChartData: categoryChartData,
    incomeCategoryChartData: incomeCategoryChartData,
    macroChartData: macroChartData,
    rawExpenses: rawExpenses
  };
}


function getMonthShortName(m) {
  const names = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
  return names[m-1];
}

function getPlanningData(year) {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  const displayData = sheet.getDataRange().getDisplayValues();
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catMacroMap = {};
  cats.forEach(c => {
    if (c.name && c.profile) {
      const cLower = String(c.name).toLowerCase().trim();
      const prof = String(c.profile).trim();
      if (prof.includes('Receita')) {
        catMacroMap[cLower] = 'Receitas';
      } else if (prof.includes('Fixa')) {
        catMacroMap[cLower] = 'Fixas';
      } else if (prof.includes('Variável') || prof.includes('Essencia')) {
        catMacroMap[cLower] = 'Variáveis Essenciais';
      } else {
        catMacroMap[cLower] = 'Discricionárias';
      }
    }
  });
  
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const currentYear = now.getFullYear();
  const currentMonthNum = now.getMonth() + 1;
  const targetYear = parseInt(year, 10);
  
  const monthsMeta = [];
  for (let i = 1; i <= 12; i++) {
    let status = 'planned';
    if (targetYear < currentYear || (targetYear === currentYear && i < currentMonthNum)) {
      status = 'realized';
    } else if (targetYear === currentYear && i === currentMonthNum) {
      status = 'current';
    }
    monthsMeta.push({ month: i, status: status, name: getMonthShortName(i) });
  }

  const incomeCategories = {};
  const expenseMacros = {
    'Fixas': {},
    'Variáveis Essenciais': {},
    'Discricionárias': {}
  };
  
  // Pre-seed categories from global configuration so customized categories appear
  cats.forEach(c => {
    if (!c.name) return;
    const cName = String(c.name).trim();
    const cLower = cName.toLowerCase();
    const macro = catMacroMap[cLower];
    if (macro === 'Receitas') {
      if (!incomeCategories[cName]) {
        incomeCategories[cName] = Array.from({length: 12}, () => ({ realized: 0, planned: 0, total: 0 }));
      }
    } else if (macro && expenseMacros[macro]) {
      if (!expenseMacros[macro][cName]) {
        expenseMacros[macro][cName] = Array.from({length: 12}, () => ({ realized: 0, planned: 0, total: 0 }));
      }
    }
  });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const d = parseFlexDate(displayData[i][5]);
    if (isNaN(d.getTime())) continue;
    if (d.getFullYear() !== targetYear) continue;

    const mIdx = d.getMonth();
    const valor = parseValor(row[4]);
    let categoria = String(row[2] || 'Outros').trim();
    let tipo = String(row[3] || '').trim();
    if (tipo.toLowerCase() === 'saída' || tipo.toLowerCase() === 'despesa') tipo = 'Saída';
    else tipo = 'Entrada';
    
    const desc = String(row[1] || '').trim();
    
    let txStatus = 'planned';
    if (d.getTime() <= today.getTime()) {
        txStatus = 'realized';
    }

    const catLower = categoria.toLowerCase();
    const explicitMacro = catMacroMap[catLower];

    let isIncome = false;
    if (explicitMacro === 'Receitas') {
      isIncome = true;
    } else if (explicitMacro) {
      isIncome = false;
    } else {
      isIncome = (tipo === 'Entrada');
    }

    if (isIncome) {
      if (!incomeCategories[categoria]) {
        incomeCategories[categoria] = Array.from({length: 12}, () => ({ realized: 0, planned: 0, total: 0 }));
      }
      incomeCategories[categoria][mIdx][txStatus] += valor;
      incomeCategories[categoria][mIdx].total += valor;
    } else {
      let macro = explicitMacro || getMacroCategory(categoria, desc, catMacroMap);
      if (!expenseMacros[macro]) macro = 'Discricionárias';
      if (!expenseMacros[macro][categoria]) {
        expenseMacros[macro][categoria] = Array.from({length: 12}, () => ({ realized: 0, planned: 0, total: 0 }));
      }
      expenseMacros[macro][categoria][mIdx][txStatus] += valor;
      expenseMacros[macro][categoria][mIdx].total += valor;
    }
  }

  const monthlyTotals = Array.from({length: 12}, () => ({ income: 0, expense: 0, balance: 0 }));
  let summary = {
    incomeRealized: 0,
    incomePlanned: 0,
    expenseRealized: 0,
    expensePlanned: 0,
    currentBalance: 0,
    projectedBalance: 0
  };

  Object.values(incomeCategories).forEach(catMonths => {
    catMonths.forEach((m, idx) => {
      monthlyTotals[idx].income += m.total;
      summary.incomeRealized += m.realized;
      summary.incomePlanned += m.planned;
    });
  });

  Object.values(expenseMacros).forEach(macroGroup => {
    Object.values(macroGroup).forEach(catMonths => {
      catMonths.forEach((m, idx) => {
        monthlyTotals[idx].expense += m.total;
        summary.expenseRealized += m.realized;
        summary.expensePlanned += m.planned;
      });
    });
  });

  monthlyTotals.forEach((m, idx) => {
    m.balance = m.income - m.expense;
  });

  summary.currentBalance = summary.incomeRealized - summary.expenseRealized;
  summary.projectedBalance = (summary.incomeRealized + summary.incomePlanned) - (summary.expenseRealized + summary.expensePlanned);

  // Merge server-side planning manual overrides if any
  try {
    const overrides = getPlanningOverridesBackend();
    const yStr = String(targetYear);
    if (overrides && overrides[yStr]) {
      const yearOv = overrides[yStr];
      Object.keys(yearOv).forEach(cat => {
        for (let m = 1; m <= 12; m++) {
          if (yearOv[cat] && yearOv[cat][m] !== undefined) {
            const ov = yearOv[cat][m];
            if (incomeCategories[cat]) {
              incomeCategories[cat][m - 1].total = ov.val;
              incomeCategories[cat][m - 1].isManualOverride = true;
            } else {
              for (const macro of ['Fixas', 'Variáveis Essenciais', 'Discricionárias']) {
                if (expenseMacros[macro] && expenseMacros[macro][cat]) {
                  expenseMacros[macro][cat][m - 1].total = ov.val;
                  expenseMacros[macro][cat][m - 1].isManualOverride = true;
                  break;
                }
              }
            }
          }
        }
      });
    }
  } catch (e) {
    console.error('Error applying server overrides', e);
  }

  return {
    year: targetYear,
    currentMonthNum: currentMonthNum,
    monthsMeta: monthsMeta,
    incomeCategories: incomeCategories,
    expenseMacros: expenseMacros,
    monthlyTotals: monthlyTotals,
    summary: summary
  };
}

function savePlanningOverrideBackend(year, cat, month, tipo, val) {
  try {
    const userProp = PropertiesService.getUserProperties();
    let overrides = {};
    const raw = userProp.getProperty('PLANNING_MANUAL_OVERRIDES');
    if (raw) {
      try { overrides = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    if (!overrides[yStr]) overrides[yStr] = {};
    if (!overrides[yStr][cat]) overrides[yStr][cat] = {};
    overrides[yStr][cat][month] = {
      val: val,
      tipo: tipo,
      updatedAt: new Date().getTime()
    };
    userProp.setProperty('PLANNING_MANUAL_OVERRIDES', JSON.stringify(overrides));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getPlanningOverridesBackend() {
  try {
    const userProp = PropertiesService.getUserProperties();
    const raw = userProp.getProperty('PLANNING_MANUAL_OVERRIDES');
    return raw ? JSON.parse(raw) : {};
  } catch(e) {
    return {};
  }
}

function removePlanningOverrideBackend(year, cat, month) {
  try {
    const userProp = PropertiesService.getUserProperties();
    let overrides = {};
    const raw = userProp.getProperty('PLANNING_MANUAL_OVERRIDES');
    if (raw) {
      try { overrides = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    if (overrides[yStr] && overrides[yStr][cat] && overrides[yStr][cat][month] !== undefined) {
      delete overrides[yStr][cat][month];
      userProp.setProperty('PLANNING_MANUAL_OVERRIDES', JSON.stringify(overrides));
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getTransactionsForCategorization() {
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
        dateStr = `${parts[2]}-${parts[1]}-${parts[0]}`;
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
}

function saveCategorizedTransactionsBatch(updates, rulesToSave, deletedRowIndices) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return true;
  
  // updates is array of { rowIndex, categoria, descricao }
  const updateMap = {};
  if (updates && updates.length > 0) {
    updates.forEach(u => {
      const cat = String(u.categoria || '').trim();
      if (u.rowIndex && cat && cat !== 'Revisar') {
        updateMap[u.rowIndex] = cat;
      }
    });
  }

  const deleteSet = new Set((deletedRowIndices || []).map(Number));
  const newData = [data[0]]; // keep header
  let modified = false;

  for (let i = 1; i < data.length; i++) {
    const rNum = i + 1;
    if (deleteSet.has(rNum)) {
      modified = true;
      continue; // Row is deleted
    }

    const row = data[i];
    if (updateMap[rNum] !== undefined) {
      row[2] = updateMap[rNum]; // Column index 2 is categoria
      modified = true;
    }
    newData.push(row);
  }
  
  if (modified) {
    sheet.clear();
    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
  }
  
  // If there are new rules to save
  if (rulesToSave && rulesToSave.length > 0) {
    saveUserRulesBatch(rulesToSave);
  }
  
  
  return true;
}

function exportToExcel() {
  const ss = getSpreadsheet();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  const params = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, params);
  const blob = response.getBlob();
  
  return {
    fileName: ss.getName() + ".xlsx",
    data: Utilities.base64Encode(blob.getBytes()),
    mimeType: blob.getContentType()
  };
}

function exportToCsv() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getDisplayValues();
  
  const csvLines = data.map(row => {
    return row.map(cell => {
      let str = "";
      if (cell instanceof Date) {
        const d = cell;
        const y = d.getFullYear();
        const m = (d.getMonth() + 1).toString().padStart(2, '0');
        const day = d.getDate().toString().padStart(2, '0');
        str = `${day}/${m}/${y}`;
      } else {
        str = String(cell);
      }
      
      // Sanitização com aspas duplas caso o campo contenha o separador ou aspas
      if (str.includes(';') || str.includes('"') || str.includes('\n')) {
        str = '"' + str.replace(/"/g, '""') + '"';
      }
      return str;
    }).join(';');
  });
  
  // Adiciona explicitamente o BOM UTF-8 no início do arquivo (\uFEFF)
  const csvString = '\uFEFF' + csvLines.join('\n');
  return {
    fileName: "Transacoes.csv",
    data: csvString,
    mimeType: "text/csv;charset=UTF-8"
  };
}

