const DEFAULT_SPREADSHEET_ID = "17Dk0WTAsgPprAmx6jIWXbRsKNrtTvu9LEfCWpG2MuC4";

function extractSpreadsheetId(input) {
  if (!input) return null;
  const str = String(input).trim();
  
  // Ex: https://docs.google.com/spreadsheets/d/17Dk0WTAsgPprAmx6jIWXbRsKNrtTvu9LEfCWpG2MuC4/edit#gid=0
  // ou docs.google.com/spreadsheets/d/17Dk0WTAsgPprAmx6jIWXbRsKNrtTvu9LEfCWpG2MuC4/...
  const match = str.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/i);
  if (match && match[1]) {
    return match[1];
  }
  
  // Se já for o ID puro (geralmente mais de 20 caracteres alfanuméricos com traços e sublinhados)
  if (/^[a-zA-Z0-9-_]{20,}$/.test(str)) {
    return str;
  }
  
  return null;
}

function getUserSpreadsheetId() {
  try {
    const userProp = PropertiesService.getUserProperties().getProperty('USER_SPREADSHEET_ID');
    if (userProp && userProp.trim().length > 0) {
      return userProp.trim();
    }
  } catch (e) {
    Logger.log("Aviso: Não foi possível obter UserProperties: " + e);
  }
  return DEFAULT_SPREADSHEET_ID;
}

function getSpreadsheet() {
  const sheetId = getUserSpreadsheetId();
  try {
    return SpreadsheetApp.openById(sheetId);
  } catch (e) {
    if (sheetId !== DEFAULT_SPREADSHEET_ID) {
      Logger.log("Falha ao abrir planilha vinculada do usuário, revertendo para padrão: " + e);
      return SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
    }
    throw e;
  }
}

function getLinkedSpreadsheetInfo() {
  const currentId = getUserSpreadsheetId();
  let title = "Planilha Principal";
  let url = "https://docs.google.com/spreadsheets/d/" + currentId + "/edit";
  const isCustom = (currentId !== DEFAULT_SPREADSHEET_ID);
  
  try {
    const ss = SpreadsheetApp.openById(currentId);
    title = ss.getName();
    url = ss.getUrl();
  } catch (e) {
    title = isCustom ? "Planilha personalizada (" + currentId.substring(0, 8) + "...)" : "Planilha Padrão";
  }

  return {
    id: currentId,
    url: url,
    title: title,
    isCustom: isCustom,
    defaultId: DEFAULT_SPREADSHEET_ID,
    defaultUrl: "https://docs.google.com/spreadsheets/d/" + DEFAULT_SPREADSHEET_ID + "/edit"
  };
}

function linkUserSpreadsheet(urlOrId) {
  const newId = extractSpreadsheetId(urlOrId);
  if (!newId) {
    return {
      success: false,
      message: "Link ou ID de planilha inválido. Cole a URL completa da planilha do Google Sheets."
    };
  }

  try {
    const ss = SpreadsheetApp.openById(newId);
    const title = ss.getName();

    // Garante que a tabela Transactions exista
    let txSheet = ss.getSheetByName('Transactions');
    if (!txSheet) {
      txSheet = ss.insertSheet('Transactions');
      const headers = ['ORIGEM', 'DESCRICAO', 'CATEGORIA', 'TIPO', 'VALOR', 'DATA'];
      txSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      txSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      txSheet.setFrozenRows(1);
    }

    // Garante que a tabela UserRules exista
    let rulesSheet = ss.getSheetByName('UserRules');
    if (!rulesSheet) {
      rulesSheet = ss.insertSheet('UserRules');
      const headers = ['Padrão (Normalizado)', 'Categoria', 'Estabelecimento'];
      rulesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      rulesSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      rulesSheet.setFrozenRows(1);
    }

    // Salva na UserProperties do usuário
    try {
      PropertiesService.getUserProperties().setProperty('USER_SPREADSHEET_ID', newId);
    } catch (err) {
      Logger.log("Aviso ao salvar UserProperty: " + err);
    }

    return {
      success: true,
      id: newId,
      url: ss.getUrl(),
      title: title,
      isCustom: (newId !== DEFAULT_SPREADSHEET_ID),
      message: 'Planilha "' + title + '" vinculada com sucesso!'
    };
  } catch (e) {
    return {
      success: false,
      message: "Não foi possível conectar à planilha. Verifique se o link está correto e se sua conta possui permissão de edição nela: " + e.message
    };
  }
}

function unlinkUserSpreadsheet() {
  try {
    PropertiesService.getUserProperties().deleteProperty('USER_SPREADSHEET_ID');
    return {
      success: true,
      id: DEFAULT_SPREADSHEET_ID,
      url: "https://docs.google.com/spreadsheets/d/" + DEFAULT_SPREADSHEET_ID + "/edit",
      title: "Planilha Padrão",
      isCustom: false,
      message: "Planilha restaurada para o padrão com sucesso."
    };
  } catch (e) {
    return {
      success: false,
      message: "Erro ao restaurar planilha padrão: " + e.message
    };
  }
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
    headers = ['ORIGEM', 'DESCRICAO', 'CATEGORIA', 'TIPO', 'VALOR', 'DATA'];
  } else if (tableName === 'UserRules') {
    headers = ['Padrão (Normalizado)', 'Categoria', 'Estabelecimento'];
  }
  
  if (headers.length > 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

