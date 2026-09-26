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

function parseFlexDate(dateVal, fallbackYear) {
  if (!dateVal && dateVal !== 0) return new Date(NaN);
  
  const targetYr = (fallbackYear !== undefined && fallbackYear !== 'all' && parseInt(fallbackYear, 10))
    ? parseInt(fallbackYear, 10)
    : new Date().getFullYear();

  // Excel / Google Sheets numeric date serial (e.g. 45000 - 60000)
  if (typeof dateVal === 'number' && dateVal > 25000 && dateVal < 80000) {
    const dNum = new Date(Math.round((dateVal - 25569) * 86400000));
    if (!isNaN(dNum.getTime())) {
      let y = dNum.getFullYear();
      if ((y === 2001 || y === 1926) && targetYr !== y) dNum.setFullYear(targetYr);
      return dNum;
    }
  }

  if (dateVal instanceof Date) {
    if (isNaN(dateVal.getTime())) return new Date(NaN);
    let y = dateVal.getFullYear();
    if (y < 100) {
      dateVal.setFullYear(y + 2000);
    }
    // If year was parsed as 2001 or 1926 due to legacy parsing of DD/MM or 2-digit year
    if ((y === 2001 || y === 1926) && targetYr !== y) {
      dateVal.setFullYear(targetYr);
    }
    return dateVal;
  }
  
  let cleanStr = String(dateVal).trim();
  if (!cleanStr) return new Date(NaN);

  // Numeric serial stored as string
  if (/^\d{5}$/.test(cleanStr)) {
    const num = parseInt(cleanStr, 10);
    if (num > 25000 && num < 80000) {
      const dNum = new Date(Math.round((num - 25569) * 86400000));
      if (!isNaN(dNum.getTime())) {
        let y = dNum.getFullYear();
        if ((y === 2001 || y === 1926) && targetYr !== y) dNum.setFullYear(targetYr);
        return dNum;
      }
    }
  }

  // Normalize dots to slashes if formatted as DD.MM.YYYY or DD.MM
  if (/^\d{1,2}\.\d{1,2}(\.\d{2,4})?$/.test(cleanStr)) {
    cleanStr = cleanStr.replace(/\./g, '/');
  }

  // Format DD/MM/YYYY or DD/MM/YY or DD/MM
  if (cleanStr.includes('/')) {
    let ptParts = cleanStr.split('/');
    if (ptParts.length === 3) {
      let d = parseInt(ptParts[0], 10);
      let m = parseInt(ptParts[1], 10) - 1;
      let y = parseInt(ptParts[2], 10);
      if (y < 100) y += 2000;
      return new Date(y, m, d);
    } else if (ptParts.length === 2) {
      // E.g. "04/09" -> 4th of September
      let d = parseInt(ptParts[0], 10);
      let m = parseInt(ptParts[1], 10) - 1;
      return new Date(targetYr, m, d);
    }
  }
  
  // Format YYYY-MM-DD or DD-MM-YYYY or DD-MM
  if (cleanStr.includes('-')) {
    let parts = cleanStr.split('T')[0].split('-');
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        // YYYY-MM-DD
        return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      } else {
        // DD-MM-YYYY
        let d = parseInt(parts[0], 10);
        let m = parseInt(parts[1], 10) - 1;
        let y = parseInt(parts[2], 10);
        if (y < 100) y += 2000;
        return new Date(y, m, d);
      }
    } else if (parts.length === 2) {
      // E.g. "04-09" -> 4th of September
      let d = parseInt(parts[0], 10);
      let m = parseInt(parts[1], 10) - 1;
      return new Date(targetYr, m, d);
    }
  }
  
  const standard = new Date(cleanStr);
  if (!isNaN(standard.getTime())) {
    let y = standard.getFullYear();
    if (y < 100) standard.setFullYear(y + 2000);
    if ((y === 2001 || y === 1926) && targetYr !== y) standard.setFullYear(targetYr);
    return standard;
  }
  
  return new Date(NaN);
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
  if (!dateVal && dateVal !== 0) return false;
  
  const targetY = (yearStr !== 'all' && parseInt(yearStr, 10)) ? parseInt(yearStr, 10) : new Date().getFullYear();
  const d = parseFlexDate(dateVal, targetY);
  
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
  cats.forEach(c => {
    if (c.name) {
      catNameMap[String(c.name).toLowerCase()] = c.name;
      catNameMap[String(c.name).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim()] = c.name;
    }
  });
  
  const targetYear = (year !== 'all' && parseInt(year, 10)) ? parseInt(year, 10) : new Date().getFullYear();
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    const displayDate = (displayData && displayData[i]) ? displayData[i][5] : '';
    const rawDate = row[5];
    
    // Check match with both display and raw
    const matched = isMatch(displayDate, year, month) || isMatch(rawDate, year, month);
    if (matched) {
      let d = parseFlexDate(rawDate, targetYear);
      if (isNaN(d.getTime()) || (d.getFullYear() !== targetYear && displayDate)) {
        const dDisp = parseFlexDate(displayDate, targetYear);
        if (!isNaN(dDisp.getTime())) d = dDisp;
      }
      
      let dateStr = '';
      if (!isNaN(d.getTime())) {
        const yStr = d.getFullYear();
        const mStr = String(d.getMonth() + 1).padStart(2, '0');
        const dStr = String(d.getDate()).padStart(2, '0');
        dateStr = `${yStr}-${mStr}-${dStr}`;
      } else {
        dateStr = String(displayDate || rawDate || '');
      }

      let tipo = String(row[3] || '').trim();
      const tipoL = tipo.toLowerCase();
      if (tipoL === 'saída' || tipoL === 'saida' || tipoL === 'despesa') tipo = 'Saída';
      else if (tipoL === 'entrada' || tipoL === 'receita') tipo = 'Entrada';
      let categoria = String(row[2] || '').trim();
      const catNorm = categoria.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().trim();
      if (categoria && catNameMap[catNorm]) {
        categoria = catNameMap[catNorm];
      } else if (categoria && catNameMap[categoria.toLowerCase()]) {
        categoria = catNameMap[categoria.toLowerCase()];
      }
      rows.push({
        rowIndex: i + 1,
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
  
  function normStr(s) {
    if (!s) return '';
    return String(s)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();
  }

  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catMacroMap = {};
  const catNameMap = {};
  cats.forEach(c => {
    if (c.name) {
      const cName = String(c.name).trim();
      const cLower = cName.toLowerCase();
      const cNorm = normStr(cName);
      catNameMap[cLower] = cName;
      catNameMap[cNorm] = cName;

      const prof = String(c.profile || '').trim();
      let macro = 'Discricionárias';
      if (prof.includes('Receita')) {
        macro = 'Receitas';
      } else if (prof.includes('Fixa')) {
        macro = 'Fixas';
      } else if (prof.includes('Variável') || prof.includes('Essencia')) {
        macro = 'Variáveis Essenciais';
      }
      catMacroMap[cLower] = macro;
      catMacroMap[cNorm] = macro;
    }
  });

  // Ensure 'SALARIO' is mapped to Receitas
  catNameMap['salario'] = 'SALARIO';
  catNameMap['salário'] = 'SALARIO';
  catNameMap['SALARIO'] = 'SALARIO';
  catMacroMap['salario'] = 'Receitas';
  catMacroMap['salário'] = 'Receitas';
  catMacroMap['SALARIO'] = 'Receitas';

  let userRules = [];
  try { userRules = getUserRules(); } catch(e) {}

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const currentYear = now.getFullYear();
  const currentMonthNum = now.getMonth() + 1;
  const targetYear = parseInt(year, 10) || currentYear;
  
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
  
  // Get excluded categories specifically for targetYear
  const excludedCatsList = getPlanningExcludedCategoriesBackend(targetYear);
  const excludedCatsSet = new Set();
  if (Array.isArray(excludedCatsList)) {
    excludedCatsList.forEach(c => {
      excludedCatsSet.add(String(c).trim().toLowerCase());
      excludedCatsSet.add(normStr(c));
    });
  }

  // Pre-seed categories from global configuration so customized categories appear (excluding ones excluded for this year)
  cats.forEach(c => {
    if (!c.name) return;
    const cName = String(c.name).trim();
    const cLower = cName.toLowerCase();
    const cNorm = normStr(cName);
    if (excludedCatsSet.has(cLower) || excludedCatsSet.has(cNorm)) return;

    const macro = catMacroMap[cLower] || catMacroMap[cNorm];
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

  // Ensure SALARIO category exists in incomeCategories unless explicitly excluded for this year
  if (!excludedCatsSet.has('salario') && !excludedCatsSet.has('salário') && !incomeCategories['SALARIO']) {
    incomeCategories['SALARIO'] = Array.from({length: 12}, () => ({ realized: 0, planned: 0, total: 0 }));
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dispVal = (displayData && displayData[i]) ? displayData[i][5] : '';
    const rawVal = row[5];
    
    // Priority: use visual display string first to avoid any UTC timezone boundary shifting
    let d = null;
    if (dispVal) {
      const dDisp = parseFlexDate(dispVal, targetYear);
      if (!isNaN(dDisp.getTime())) {
        d = dDisp;
      }
    }
    if (!d || isNaN(d.getTime())) {
      d = parseFlexDate(rawVal, targetYear);
    }
    if (!d || isNaN(d.getTime())) continue;
    if (d.getFullYear() !== targetYear) continue;

    const mIdx = d.getMonth();
    const valor = Math.abs(parseValor(row[4]));
    if (valor === 0) continue;

    let categoria = String(row[2] || '').trim();
    let tipo = String(row[3] || '').trim();
    const desc = String(row[1] || '').trim();
    const normDesc = normStr(desc);

    // Rule-based classification if category is empty, 'Outros', or 'Revisar'
    if (!categoria || categoria.toLowerCase() === 'outros' || categoria.toLowerCase() === 'revisar') {
      if (userRules && userRules.length > 0) {
        for (let r = 0; r < userRules.length; r++) {
          const pat = normStr(userRules[r].pattern);
          if (pat && normDesc.includes(pat)) {
            categoria = userRules[r].categoria;
            break;
          }
        }
      }
    }

    // Heuristics for salary if description matches
    if (normDesc.includes('salario') || normDesc.includes('remuneracao') || normDesc.includes('provento') || normDesc.includes('folha de pag')) {
      categoria = 'SALARIO';
      tipo = 'Entrada';
    }

    // Canonical category name
    const catNorm = normStr(categoria);
    if (categoria.toLowerCase() === 'salario' || categoria.toLowerCase() === 'salário') {
      categoria = 'SALARIO';
    } else if (categoria && catNameMap[catNorm]) {
      categoria = catNameMap[catNorm];
    } else if (categoria && catNameMap[categoria.toLowerCase()]) {
      categoria = catNameMap[categoria.toLowerCase()];
    } else if (!categoria) {
      categoria = (tipo.toLowerCase() === 'entrada' || tipo.toLowerCase() === 'receita') ? 'Outras Receitas' : 'Outros';
    }

    // Check if category is excluded specifically for targetYear
    const catClean = categoria.trim().toLowerCase();
    const catNormClean = normStr(categoria);
    if (excludedCatsSet.has(catClean) || excludedCatsSet.has(catNormClean)) {
      continue;
    }

    const explicitMacro = catMacroMap[catNorm] || catMacroMap[categoria.toLowerCase()];

    let isIncome = false;
    if (categoria === 'SALARIO' || explicitMacro === 'Receitas') {
      isIncome = true;
    } else if (tipo.toLowerCase() === 'entrada' || tipo.toLowerCase() === 'receita') {
      isIncome = true;
    } else if (catNorm.includes('salario') || catNorm.includes('rendimento') || catNorm.includes('receita')) {
      isIncome = true;
    } else if (explicitMacro) {
      isIncome = false;
    } else {
      isIncome = (tipo.toLowerCase() === 'entrada' || tipo.toLowerCase() === 'receita');
    }

    let txStatus = 'planned';
    if (d.getTime() <= today.getTime() || (d.getFullYear() === currentYear && (mIdx + 1) <= currentMonthNum)) {
      txStatus = 'realized';
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

  // Merge server-side planning manual overrides if any
  try {
    const overrides = getPlanningOverridesBackend();
    const yStr = String(targetYear);
    if (overrides && overrides[yStr]) {
      const yearOv = overrides[yStr];
      let ovBackendChanged = false;
      Object.keys(yearOv).forEach(cat => {
        const catClean = cat.trim().toLowerCase();
        const catNormClean = normStr(cat);
        if (excludedCatsSet.has(catClean) || excludedCatsSet.has(catNormClean)) {
          return;
        }
        for (let m = 1; m <= 12; m++) {
          if (yearOv[cat] && yearOv[cat][m] !== undefined) {
            const ov = yearOv[cat][m];
            // Apenas para o ano corrente (2026), se for resíduo zerado de Outubro em ENTRADAS ou JANETE
            if (targetYear === 2026 && (catClean === 'entradas' || catClean === 'janete') && m === 10 && ov.val === 0) {
              delete yearOv[cat][m];
              ovBackendChanged = true;
              continue;
            }
            if (incomeCategories[cat]) {
              const cell = incomeCategories[cat][m - 1];
              // Never let an override of 0 or deleted-forecast wipe out actual realized transactions!
              if (cell.realized === 0 || (ov.val > 0 && !ov.isDeletedForecast)) {
                cell.total = ov.val;
                cell.isManualOverride = true;
              }
            } else {
              for (const macro of ['Fixas', 'Variáveis Essenciais', 'Discricionárias']) {
                if (expenseMacros[macro] && expenseMacros[macro][cat]) {
                  const cell = expenseMacros[macro][cat][m - 1];
                  if (cell.realized === 0 || (ov.val > 0 && !ov.isDeletedForecast)) {
                    cell.total = ov.val;
                    cell.isManualOverride = true;
                  }
                  break;
                }
              }
            }
          }
        }
      });
      if (ovBackendChanged) {
        try {
          const userProp = PropertiesService.getUserProperties();
          userProp.setProperty('PLANNING_MANUAL_OVERRIDES', JSON.stringify(overrides));
        } catch(e) {}
      }
    }
  } catch (e) {
    console.error('Error applying server overrides', e);
  }

  // Ensure NO excluded category remains in the result for targetYear
  Object.keys(incomeCategories).forEach(cat => {
    if (excludedCatsSet.has(cat.trim().toLowerCase()) || excludedCatsSet.has(normStr(cat))) {
      delete incomeCategories[cat];
    }
  });
  ['Fixas', 'Variáveis Essenciais', 'Discricionárias'].forEach(macro => {
    if (expenseMacros[macro]) {
      Object.keys(expenseMacros[macro]).forEach(cat => {
        if (excludedCatsSet.has(cat.trim().toLowerCase()) || excludedCatsSet.has(normStr(cat))) {
          delete expenseMacros[macro][cat];
        }
      });
    }
  });

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

function savePlanningForecastsBackend(year, forecastsJson) {
  try {
    const userProp = PropertiesService.getUserProperties();
    let allForecasts = {};
    const raw = userProp.getProperty('PLANNING_AUTOMATED_FORECASTS');
    if (raw) {
      try { allForecasts = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    const parsedData = (typeof forecastsJson === 'string') ? JSON.parse(forecastsJson) : forecastsJson;
    allForecasts[yStr] = parsedData[yStr] || parsedData;
    userProp.setProperty('PLANNING_AUTOMATED_FORECASTS', JSON.stringify(allForecasts));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function getPlanningForecastsBackend() {
  try {
    const userProp = PropertiesService.getUserProperties();
    const raw = userProp.getProperty('PLANNING_AUTOMATED_FORECASTS');
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

function getPlanningExcludedCategoriesBackend(year) {
  try {
    const userProp = PropertiesService.getUserProperties();
    const raw = userProp.getProperty('PLANNING_EXCLUDED_CATEGORIES');
    if (!raw) return [];
    const obj = JSON.parse(raw);
    const yStr = String(year);
    return Array.isArray(obj[yStr]) ? obj[yStr] : [];
  } catch(e) {
    return [];
  }
}

function savePlanningExcludedCategoryBackend(year, cat) {
  try {
    if (!cat) return { success: false };
    const userProp = PropertiesService.getUserProperties();
    let obj = {};
    const raw = userProp.getProperty('PLANNING_EXCLUDED_CATEGORIES');
    if (raw) {
      try { obj = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    if (!Array.isArray(obj[yStr])) obj[yStr] = [];
    const catClean = String(cat).trim();
    if (!obj[yStr].some(c => c.toLowerCase() === catClean.toLowerCase())) {
      obj[yStr].push(catClean);
    }
    userProp.setProperty('PLANNING_EXCLUDED_CATEGORIES', JSON.stringify(obj));
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function removePlanningExcludedCategoryBackend(year, cat) {
  try {
    if (!cat) return { success: false };
    const userProp = PropertiesService.getUserProperties();
    let obj = {};
    const raw = userProp.getProperty('PLANNING_EXCLUDED_CATEGORIES');
    if (raw) {
      try { obj = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    if (Array.isArray(obj[yStr])) {
      const catClean = String(cat).trim().toLowerCase();
      obj[yStr] = obj[yStr].filter(c => String(c).trim().toLowerCase() !== catClean);
      userProp.setProperty('PLANNING_EXCLUDED_CATEGORIES', JSON.stringify(obj));
    }
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

function savePlanningExcludedCategoriesBackend(year, catsArray) {
  try {
    const userProp = PropertiesService.getUserProperties();
    let obj = {};
    const raw = userProp.getProperty('PLANNING_EXCLUDED_CATEGORIES');
    if (raw) {
      try { obj = JSON.parse(raw); } catch(e) {}
    }
    const yStr = String(year);
    const arr = Array.isArray(catsArray) ? catsArray : (typeof catsArray === 'string' ? JSON.parse(catsArray) : []);
    obj[yStr] = arr;
    userProp.setProperty('PLANNING_EXCLUDED_CATEGORIES', JSON.stringify(obj));
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

function applyUserRulesToTransactionsBackend(rulesList) {
  backupForUndo();
  const txSheet = getTable('Transactions');
  const txData = txSheet.getDataRange().getValues();
  if (txData.length <= 1) return { success: true, updatedCount: 0 };
  
  let rules = rulesList;
  if (!rules || !Array.isArray(rules) || rules.length === 0) {
    try {
      rules = getUserRules();
    } catch(e) {
      rules = [];
    }
  }
  if (!rules || rules.length === 0) return { success: true, updatedCount: 0 };
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catNameMap = {};
  cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);

  // Normalization helper for resilient pattern matching (accents, punctuation, multiple spaces)
  function normForMatch(s) {
    if (!s) return '';
    return String(s)
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toUpperCase()
      .replace(/[^A-Z0-9 ]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  let updatedCount = 0;
  for (let i = 1; i < txData.length; i++) {
    const rawDesc = String(txData[i][1] || '').trim().toUpperCase();
    const currentCat = String(txData[i][2] || '').trim();
    const cleanDesc = normForMatch(rawDesc);
    
    for (let r = 0; r < rules.length; r++) {
      const pattern = String(rules[r].pattern || '').trim().toUpperCase();
      const normPattern = normForMatch(pattern);
      if (pattern && (rawDesc.includes(pattern) || (normPattern && cleanDesc.includes(normPattern)))) {
        let targetCat = rules[r].categoria;
        if (targetCat && catNameMap[targetCat.toLowerCase()]) {
          targetCat = catNameMap[targetCat.toLowerCase()];
        }
        if (targetCat && targetCat !== currentCat) {
          txData[i][2] = targetCat;
          updatedCount++;
        }
        break;
      }
    }
  }

  if (updatedCount > 0) {
    txSheet.getRange(1, 1, txData.length, txData[0].length).setValues(txData);
    SpreadsheetApp.flush();
  }

  return { success: true, updatedCount: updatedCount };
}

function updateSingleTransactionCategory(rowIndex, newCategory, txDetails) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { success: false, message: 'Nenhuma transação encontrada.' };

  let targetRowIndex = -1;
  const numRows = data.length;

  // 1. Try direct rowIndex if valid and matches description
  if (rowIndex && rowIndex >= 2 && rowIndex <= numRows) {
    if (!txDetails || !txDetails.descricao) {
      targetRowIndex = rowIndex;
    } else {
      const targetDesc = String(txDetails.descricao || '').trim().toLowerCase();
      const currentDesc = String(data[rowIndex - 1][1] || '').trim().toLowerCase();
      if (currentDesc === targetDesc) {
        targetRowIndex = rowIndex;
      }
    }
  }

  // 2. Fallback search by description, value and date
  if (targetRowIndex === -1 && txDetails) {
    const targetDesc = String(txDetails.descricao || '').trim().toLowerCase();
    const targetVal = parseValor(txDetails.valor);
    const targetOrigem = String(txDetails.origem || '').trim().toLowerCase();

    for (let i = numRows - 1; i >= 1; i--) {
      const rowDesc = String(data[i][1] || '').trim().toLowerCase();
      const rowVal = parseValor(data[i][4]);
      const rowOrigem = String(data[i][0] || '').trim().toLowerCase();

      const descMatch = (rowDesc === targetDesc);
      const valMatch = (isNaN(targetVal) || Math.abs(rowVal - targetVal) < 0.01);
      const origemMatch = (!targetOrigem || rowOrigem === targetOrigem);

      if (descMatch && valMatch && origemMatch) {
        targetRowIndex = i + 1;
        break;
      }
    }

    if (targetRowIndex === -1) {
      for (let i = numRows - 1; i >= 1; i--) {
        const rowDesc = String(data[i][1] || '').trim().toLowerCase();
        if (rowDesc === targetDesc) {
          targetRowIndex = i + 1;
          break;
        }
      }
    }
  }

  if (targetRowIndex !== -1) {
    sheet.getRange(targetRowIndex, 3).setValue(newCategory || ''); // Column 3 is Categoria (C)
    SpreadsheetApp.flush();
    return { success: true, rowIndex: targetRowIndex, newCategory: newCategory };
  }

  return { success: false, message: 'Transação não encontrada na planilha.' };
}

function updateAllMatchingTransactionsCategory(descricao, newCategory) {
  backupForUndo();
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return { updatedCount: 0 };

  const targetDesc = String(descricao || '').trim().toLowerCase();
  let updatedCount = 0;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim().toLowerCase() === targetDesc) {
      data[i][2] = newCategory || '';
      updatedCount++;
    }
  }

  if (updatedCount > 0) {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    SpreadsheetApp.flush();
  }
  return { updatedCount: updatedCount, newCategory: newCategory };
}



