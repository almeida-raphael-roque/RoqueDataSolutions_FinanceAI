/**
 * Conexão com Google Sheets e abstração de banco de dados
 */
function getSpreadsheet() {
  return SpreadsheetApp.openById("17Dk0WTAsgPprAmx6jIWXbRsKNrtTvu9LEfCWpG2MuC4");
}

function getTable(tableName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(tableName);
  
  if (!sheet) {
    sheet = createTable(tableName);
  }
  return sheet;
}

function createTable(tableName) {
  const ss = getSpreadsheet();
  const sheet = ss.insertSheet(tableName);
  let headers = [];
  
  if (tableName === 'Transactions') {
    headers = ['DATA', 'VALOR', 'TIPO', 'DESCRICAO', 'CATEGORIA'];
  }
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}
