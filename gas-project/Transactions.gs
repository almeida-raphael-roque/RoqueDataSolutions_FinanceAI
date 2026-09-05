/**
 * Funções de Transações e Dashboard
 */
function saveTransaction(data, valor, tipo, descricao, categoria) {
  const sheet = getTable('Transactions');
  sheet.appendRow([data, valor, tipo, descricao, categoria]);
  return true;
}

function isMatch(dateVal, yearStr, monthStr) {
  if (yearStr === 'all' && monthStr === 'all') return true;
  if (!dateVal) return false;
  
  let d;
  if (dateVal instanceof Date) {
    d = dateVal;
  } else {
    // Basic parse (ex: YYYY-MM-DD)
    let parts = dateVal.toString().split('T')[0].split('-');
    if (parts.length === 3) {
      d = new Date(parts[0], parseInt(parts[1])-1, parts[2]);
    } else {
      d = new Date(dateVal);
    }
  }
  
  if (isNaN(d.getTime())) return true; // Se não for data válida, inclui (fallback)
  
  const y = d.getFullYear().toString();
  const m = (d.getMonth() + 1).toString();
  
  const matchY = (yearStr === 'all' || yearStr === y);
  const matchM = (monthStr === 'all' || monthStr === m);
  
  return matchY && matchM;
}

function getTransactions(year = 'all', month = 'all') {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    if (isMatch(row[0], year, month)) {
      // Formatação básica para retornar como string consistente
      let dateStr = row[0];
      if (dateStr instanceof Date) {
        dateStr = Utilities.formatDate(dateStr, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      rows.push({
        data: dateStr,
        valor: row[1],
        tipo: row[2],
        descricao: row[3],
        categoria: row[4]
      });
    }
  }
  return rows;
}

function getMacroCategory(categoria, descricao) {
  const cat = (categoria || '').toLowerCase();
  const desc = (descricao || '').toLowerCase();
  
  if (cat.includes('moradia') || desc.includes('aluguel') || desc.includes('internet') || desc.includes('condomínio') || cat.includes('educação')) return 'Fixas';
  
  if (desc.includes('energia') || desc.includes('água') || desc.includes('gás') || desc.includes('luz') || cat.includes('saúde') || desc.includes('mercado') || desc.includes('supermercado') || desc.includes('farmácia') || cat.includes('transporte')) return 'Variáveis Essenciais';
  
  if (cat.includes('lazer') || desc.includes('restaurante') || desc.includes('delivery') || desc.includes('compras') || desc.includes('passeio') || desc.includes('cinema') || desc.includes('netflix') || desc.includes('uber') || cat.includes('alimentação')) {
    if (desc.includes('mercado') || desc.includes('supermercado')) return 'Variáveis Essenciais';
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
    const valor = parseFloat(row[1]) || 0;
    const tipo = row[2];
    const categoria = row[4] || 'Outros';
    const desc = row[3] || 'Desconhecido';
    
    // Parse date once
    let d;
    const dateVal = row[0];
    if (dateVal instanceof Date) {
      d = dateVal;
    } else {
      let parts = dateVal.toString().split('T')[0].split('-');
      if (parts.length === 3) d = new Date(parts[0], parseInt(parts[1])-1, parts[2]);
      else d = new Date(dateVal);
    }
    
    // Save raw expenses for temporal drill-down (all-time)
    if (tipo === 'Saída' && !isNaN(d.getTime())) {
      rawExpenses.push({
        date: d.getTime(), // ms timestamp
        desc: desc,
        value: valor
      });
    }
    
    // Soma para os cards principais respeitando MÊS e ANO
    if (isMatch(row[0], year, month)) {
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
      const matchYearForChart = (year === 'all' || year === y.toString());
      
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
