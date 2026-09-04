/**
 * Funções de Transações e Dashboard
 */
function saveTransaction(data, valor, tipo, descricao, categoria) {
  const sheet = getTable('Transactions');
  sheet.appendRow([data, valor, tipo, descricao, categoria]);
  return true;
}

function getTransactions() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return []; 
  
  const rows = [];
  for (let i = data.length - 1; i >= 1; i--) { 
    const row = data[i];
    rows.push({
      data: row[0],
      valor: row[1],
      tipo: row[2],
      descricao: row[3],
      categoria: row[4]
    });
  }
  return rows;
}

function getDashboardData() {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  
  let income = 0;
  let expense = 0;
  
  for (let i = 1; i < data.length; i++) {
    const valor = parseFloat(data[i][1]) || 0;
    const tipo = data[i][2];
    if (tipo === 'Entrada') {
      income += valor;
    } else if (tipo === 'Saída') {
      expense += valor;
    }
  }
  
  return {
    balance: income - expense,
    income: income,
    expense: expense
  };
}
