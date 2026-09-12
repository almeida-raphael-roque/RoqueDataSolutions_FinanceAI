/**
 * Funções de Transações e Dashboard
 */
function saveTransaction(origem, descricao, categoria, tipo, valor, data) {
  backupForUndo();
  const sheet = getTable('Transactions');
  sheet.appendRow([origem, descricao, categoria, tipo, valor, data]);
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
  return true;
}

/**
 * Exclui transações com base no ano, intervalo de meses e categoria / macrocategoria
 */
function deleteTransactionsRange(year, startMonth, endMonth, category, tipo) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return true;

  const header = data[0];
  const toKeep = [header];
  const yTarget = parseInt(year, 10);
  const mStart = Math.min(parseInt(startMonth, 10), parseInt(endMonth, 10));
  const mEnd = Math.max(parseInt(startMonth, 10), parseInt(endMonth, 10));

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const d = parseFlexDate(row[5]);
    if (!isNaN(d.getTime())) {
      const rowYear = d.getFullYear();
      const rowMonth = d.getMonth() + 1;
      if (rowYear === yTarget && rowMonth >= mStart && rowMonth <= mEnd) {
        let match = false;
        const rowCat = String(row[2] || '').trim();
        const rowTipo = String(row[3] || '').trim();
        const rowDesc = String(row[1] || '').trim();

        if (category === '__ALL_INCOME__') {
          match = (rowTipo === 'Entrada');
        } else if (category === '__ALL_EXPENSE__') {
          match = (rowTipo === 'Saída');
        } else if (category.startsWith('__MACRO_')) {
          const macroKey = category.replace('__MACRO_', '');
          const rowMacro = getMacroCategory(rowCat, rowDesc);
          match = (rowMacro === macroKey && (!tipo || rowTipo === tipo));
        } else {
          match = (rowCat === category && (!tipo || rowTipo === tipo));
        }

        if (match) {
          // Excluir (não adiciona ao toKeep)
          continue;
        }
      }
    }
    toKeep.push(row);
  }

  sheet.clear();
  if (toKeep.length > 0) {
    sheet.getRange(1, 1, toKeep.length, toKeep[0].length).setValues(toKeep);
  }
  return true;
}

function clearTransactions(yearStr, monthStr) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return true; // Somente cabeçalho
  
  const header = data[0];
  const toKeep = [header];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!isMatch(row[5], yearStr, monthStr)) {
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
  if (data.length <= 1) return []; 
  
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    if (isMatch(row[5], year, month)) {
      // Formatação básica para retornar como string consistente
      let dateStr = row[5];
      if (dateStr instanceof Date) {
        dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      rows.push({
        origem: row[0],
        descricao: row[1],
        categoria: row[2],
        tipo: row[3],
        valor: parseValor(row[4]),
        data: dateStr
      });
    }
  }
  return rows;
}

function getMacroCategory(categoria, descricao) {
  const cat = (categoria || '').toLowerCase();
  const desc = (descricao || '').toLowerCase();
  
  if (cat.includes('moradia') || desc.includes('aluguel') || desc.includes('internet') || desc.includes('condomínio') || cat.includes('educação') || cat.includes('telefone') || cat.includes('celular') || desc.includes('vivo') || desc.includes('claro') || desc.includes('tim')) return 'Fixas';
  
  if (desc.includes('energia') || desc.includes('água') || desc.includes('gás') || desc.includes('luz') || cat.includes('saúde') || desc.includes('mercado') || desc.includes('supermercado') || desc.includes('farmácia') || cat.includes('transporte')) return 'Variáveis Essenciais';
  
  if (cat.includes('lazer') || desc.includes('restaurante') || desc.includes('delivery') || desc.includes('compras') || desc.includes('passeio') || desc.includes('cinema') || desc.includes('netflix') || desc.includes('uber') || cat.includes('alimentação')) {
    if (desc.includes('mercado') || desc.includes('supermercado')) return 'Variáveis Essenciais';
    return 'Discricionárias';
  }
  
  return 'Fixas'; // Default behavior to align with client-side mapping for new unmapped categories
}

const MACRO_COLORS = {
  'Fixas': '#991b1b', // Vermelho escuro
  'Variáveis Essenciais': '#ef4444', // Vermelho intermediário
  'Discricionárias': '#f87171' // Vermelho claro
};

function getDashboardData(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  
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
    const origem = row[0];
    const desc = row[1] || 'Desconhecido';
    const categoria = row[2] || 'Outros';
    const tipo = row[3];
    const valor = parseValor(row[4]);
    const dataVal = row[5];
    
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
        const macro = getMacroCategory(categoria, desc);
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

function getTransactionsForCategorization() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    let dateStr = row[5];
    if (dateStr instanceof Date) {
      dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
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
}

function saveCategorizedTransactionsBatch(updates, rulesToSave, deletedRowIndices) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return true;
  
  // updates is array of { rowIndex, categoria, descricao }
  const updateMap = {};
  const descMap = {};
  if (updates && updates.length > 0) {
    updates.forEach(u => {
      if (u.rowIndex) {
        updateMap[u.rowIndex] = u.categoria;
      }
      if (u.descricao && u.categoria) {
        descMap[String(u.descricao).trim().toUpperCase()] = u.categoria;
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
    const rowDesc = String(row[1] || '').trim().toUpperCase();
    if (updateMap[rNum] !== undefined) {
      row[2] = updateMap[rNum]; // Column index 2 is categoria
      modified = true;
    } else if (descMap[rowDesc] !== undefined) {
      row[2] = descMap[rowDesc]; // Overwrite category matching this description across all rows
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
  const data = sheet.getDataRange().getValues();
  
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

/**
 * Retorna os dados agregados para o Planejamento Financeiro Anual
 * Progressão: REALIZADO -> MÊS ATUAL -> PREVISTO
 */
function getPlanningData(targetYear) {
  const now = new Date();
  const currentSystemYear = now.getFullYear();
  const currentSystemMonth = now.getMonth() + 1; // 1-12
  const currentSystemDay = now.getDate();

  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();

  const yearsSet = new Set();
  yearsSet.add(currentSystemYear);

  const parsedYear = targetYear ? parseInt(targetYear, 10) : currentSystemYear;

  const validRows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dataVal = row[5];
    const d = parseFlexDate(dataVal);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      yearsSet.add(y);
      if (y === parsedYear) {
        let dateStr = dataVal;
        if (dateStr instanceof Date) {
          dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        validRows.push({
          origem: row[0] || '',
          descricao: String(row[1] || '').trim(),
          categoria: String(row[2] || '').trim() || 'Outros',
          tipo: String(row[3] || '').trim(),
          valor: parseValor(row[4]),
          data: dateStr,
          d: d,
          month: d.getMonth() + 1,
          day: d.getDate()
        });
      }
    }
  }

  const availableYears = Array.from(yearsSet).sort((a, b) => b - a);

  const monthNames = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];
  const monthShort = ['JAN', 'FEV', 'MAR', 'ABR', 'MAI', 'JUN', 'JUL', 'AGO', 'SET', 'OUT', 'NOV', 'DEZ'];

  // Status de cada mês no ano selecionado
  const monthsMeta = [];
  for (let m = 1; m <= 12; m++) {
    let status = 'realized';
    if (parsedYear < currentSystemYear) {
      status = 'realized';
    } else if (parsedYear > currentSystemYear) {
      status = 'planned';
    } else {
      if (m < currentSystemMonth) status = 'realized';
      else if (m === currentSystemMonth) status = 'current';
      else status = 'planned';
    }
    monthsMeta.push({
      month: m,
      name: monthNames[m - 1],
      shortName: monthShort[m - 1],
      status: status
    });
  }

  // Inicializa estruturas de dados por mês
  const incomeCategoryMap = {}; // categoria -> { months: [1..12: {realized, planned, total}], totalYear, totalRealized, totalPlanned }
  const expenseMacroMap = {
    'Fixas': { label: 'Despesas Fixas', months: Array.from({length: 12}, () => ({realized: 0, planned: 0, total: 0})), totalYear: 0, subcategories: {} },
    'Variáveis Essenciais': { label: 'Variáveis Essenciais', months: Array.from({length: 12}, () => ({realized: 0, planned: 0, total: 0})), totalYear: 0, subcategories: {} },
    'Discricionárias': { label: 'Despesas Discricionárias', months: Array.from({length: 12}, () => ({realized: 0, planned: 0, total: 0})), totalYear: 0, subcategories: {} }
  };

  const monthlyTotals = Array.from({length: 12}, () => ({
    incomeRealized: 0,
    incomePlanned: 0,
    incomeTotal: 0,
    expenseRealized: 0,
    expensePlanned: 0,
    expenseTotal: 0,
    monthlyBalance: 0,
    accumulatedBalance: 0,
    status: 'realized'
  }));

  // Associa status a cada mês
  monthsMeta.forEach((mm, idx) => {
    monthlyTotals[idx].status = mm.status;
  });

  // Processa as transações do ano
  validRows.forEach(tx => {
    const mIdx = tx.month - 1;
    if (mIdx < 0 || mIdx > 11) return;

    const mStatus = monthsMeta[mIdx].status;
    let isTxRealized = true;

    if (mStatus === 'realized') {
      isTxRealized = true;
    } else if (mStatus === 'planned') {
      isTxRealized = false;
    } else {
      // Mês atual: antes ou até hoje é realizado, após hoje é previsto
      if (tx.day <= currentSystemDay) {
        isTxRealized = true;
      } else {
        isTxRealized = false;
      }
    }

    const txObj = {
      origem: tx.origem,
      descricao: tx.descricao,
      categoria: tx.categoria,
      tipo: tx.tipo,
      valor: tx.valor,
      data: tx.data,
      isRealized: isTxRealized,
      statusLabel: isTxRealized ? 'Realizado' : 'Previsto'
    };

    if (tx.tipo === 'Entrada') {
      const cat = tx.categoria;
      if (!incomeCategoryMap[cat]) {
        incomeCategoryMap[cat] = {
          label: cat,
          months: Array.from({length: 12}, () => ({realized: 0, planned: 0, total: 0, txs: []})),
          totalYear: 0,
          totalRealized: 0,
          totalPlanned: 0
        };
      }

      if (isTxRealized) {
        incomeCategoryMap[cat].months[mIdx].realized += tx.valor;
        incomeCategoryMap[cat].totalRealized += tx.valor;
        monthlyTotals[mIdx].incomeRealized += tx.valor;
      } else {
        incomeCategoryMap[cat].months[mIdx].planned += tx.valor;
        incomeCategoryMap[cat].totalPlanned += tx.valor;
        monthlyTotals[mIdx].incomePlanned += tx.valor;
      }

      incomeCategoryMap[cat].months[mIdx].total += tx.valor;
      incomeCategoryMap[cat].months[mIdx].txs.push(txObj);
      incomeCategoryMap[cat].totalYear += tx.valor;
      monthlyTotals[mIdx].incomeTotal += tx.valor;

    } else if (tx.tipo === 'Saída') {
      let macro = getMacroCategory(tx.categoria, tx.descricao);
      if (!expenseMacroMap[macro]) {
        macro = 'Discricionárias';
      }

      const macroObj = expenseMacroMap[macro];
      if (isTxRealized) {
        macroObj.months[mIdx].realized += tx.valor;
        monthlyTotals[mIdx].expenseRealized += tx.valor;
      } else {
        macroObj.months[mIdx].planned += tx.valor;
        monthlyTotals[mIdx].expensePlanned += tx.valor;
      }
      macroObj.months[mIdx].total += tx.valor;
      macroObj.totalYear += tx.valor;
      monthlyTotals[mIdx].expenseTotal += tx.valor;

      const cat = tx.categoria;
      if (!macroObj.subcategories[cat]) {
        macroObj.subcategories[cat] = {
          label: cat,
          macro: macro,
          months: Array.from({length: 12}, () => ({realized: 0, planned: 0, total: 0, txs: []})),
          totalYear: 0,
          totalRealized: 0,
          totalPlanned: 0
        };
      }

      const subCat = macroObj.subcategories[cat];
      if (isTxRealized) {
        subCat.months[mIdx].realized += tx.valor;
        subCat.totalRealized += tx.valor;
      } else {
        subCat.months[mIdx].planned += tx.valor;
        subCat.totalPlanned += tx.valor;
      }
      subCat.months[mIdx].total += tx.valor;
      subCat.months[mIdx].txs.push(txObj);
      subCat.totalYear += tx.valor;
    }
  });

  // Calcula Saldos Mensais e Saldo Acumulado
  let runningAccumulated = 0;
  let currentAccumulatedBalance = 0;
  let totalYearIncome = 0;
  let totalYearIncomeRealized = 0;
  let totalYearIncomePlanned = 0;
  let totalYearExpense = 0;
  let totalYearExpenseRealized = 0;
  let totalYearExpensePlanned = 0;

  for (let m = 0; m < 12; m++) {
    const item = monthlyTotals[m];
    totalYearIncome += item.incomeTotal;
    totalYearIncomeRealized += item.incomeRealized;
    totalYearIncomePlanned += item.incomePlanned;

    totalYearExpense += item.expenseTotal;
    totalYearExpenseRealized += item.expenseRealized;
    totalYearExpensePlanned += item.expensePlanned;

    // Fórmula do saldo conforme especificação:
    // Meses realizados: Realizado - Realizado
    // Mês atual: Total (Realizado + Previsto) Receitas - Total Despesas
    // Meses futuros: Previsto Receitas - Previsto Despesas (ou Total se lançado como previsto)
    if (item.status === 'realized') {
      item.monthlyBalance = item.incomeRealized - item.expenseRealized;
    } else if (item.status === 'current') {
      item.monthlyBalance = item.incomeTotal - item.expenseTotal;
    } else {
      item.monthlyBalance = (item.incomePlanned || item.incomeTotal) - (item.expensePlanned || item.expenseTotal);
    }

    runningAccumulated += item.monthlyBalance;
    item.accumulatedBalance = runningAccumulated;

    // Posição realizada até a data presente
    if (item.status === 'realized' || (item.status === 'current' && parsedYear === currentSystemYear)) {
      currentAccumulatedBalance = runningAccumulated;
    }
  }

  const projectedYearEndBalance = runningAccumulated;

  // Transforma mapas em arrays ordenados para visualização
  const incomeCategoriesList = Object.values(incomeCategoryMap).sort((a, b) => b.totalYear - a.totalYear);

  const expenseMacrosList = Object.keys(expenseMacroMap).map(macroKey => {
    const mData = expenseMacroMap[macroKey];
    const subcats = Object.values(mData.subcategories).sort((a, b) => b.totalYear - a.totalYear);
    return {
      key: macroKey,
      label: mData.label,
      months: mData.months,
      totalYear: mData.totalYear,
      subcategories: subcats
    };
  });

  return {
    year: parsedYear,
    currentSystemYear: currentSystemYear,
    currentSystemMonth: currentSystemMonth,
    currentSystemDay: currentSystemDay,
    availableYears: availableYears,
    monthsMeta: monthsMeta,
    monthlyTotals: monthlyTotals,
    incomeCategories: incomeCategoriesList,
    expenseMacros: expenseMacrosList,
    summary: {
      totalYearIncome: totalYearIncome,
      totalYearIncomeRealized: totalYearIncomeRealized,
      totalYearIncomePlanned: totalYearIncomePlanned,
      totalYearExpense: totalYearExpense,
      totalYearExpenseRealized: totalYearExpenseRealized,
      totalYearExpensePlanned: totalYearExpensePlanned,
      currentAccumulatedBalance: currentAccumulatedBalance,
      projectedYearEndBalance: projectedYearEndBalance
    }
  };
}
