function getCategoriesSheet() {
  return getTable('Categorias');
}

function ensureCategoriesSheet() {
  const sheet = getCategoriesSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    sheet.appendRow(['id', 'name', 'profile']);
    const defaultCats = [
      [1, "Alimentação", "Despesa Discricionária"],
      [2, "Transporte", "Despesa Discricionária"],
      [3, "Moradia", "Despesa Fixa"],
      [4, "Saúde", "Despesa Fixa"],
      [5, "Lazer", "Despesa Discricionária"],
      [6, "Educação", "Despesa Fixa"],
      [7, "Salário", "Receita Fixa"],
      [8, "Rendimentos", "Receita Variável"],
      [9, "Outros", "Despesa Variável"]
    ];
    sheet.getRange(2, 1, defaultCats.length, 3).setValues(defaultCats);
  }
}

function getGlobalCategories(txDataPreloaded = null) {
  ensureCategoriesSheet();
  const catSheet = getCategoriesSheet();
  const catData = catSheet.getDataRange().getValues();
  const cats = [];
  const existingNames = new Set();
  
  for (let i = 1; i < catData.length; i++) {
    const cName = String(catData[i][1]).trim();
    if (cName) {
      existingNames.add(cName.toLowerCase());
      cats.push({
        id: catData[i][0],
        name: cName,
        profile: catData[i][2]
      });
    }
  }

  // Auto-sync from Transactions sheet
  let txData = txDataPreloaded;
  if (!txData) {
    const txSheet = getTable('Transactions');
    if (txSheet) {
      txData = txSheet.getDataRange().getValues();
    }
  }

  if (txData && txData.length > 1) {
    let newRows = [];
    for (let i = 1; i < txData.length; i++) {
      const row = txData[i];
      const catName = String(row[2] || '').trim();
      const tipo = String(row[3] || '').trim();
      
      if (catName && catName !== 'Outros' && catName !== 'Revisar' && !existingNames.has(catName.toLowerCase())) {
        existingNames.add(catName.toLowerCase());
        const newId = new Date().getTime() + Math.floor(Math.random() * 1000) + i;
        const profile = (tipo === 'Entrada' || catName.toLowerCase().includes('salário') || catName.toLowerCase().includes('receita')) ? 'Receita Variável' : 'Despesa Variável';
        cats.push({
          id: newId,
          name: catName,
          profile: profile
        });
        newRows.push([newId, catName, profile]);
      }
    }
    
    if (newRows.length > 0) {
      catSheet.getRange(catData.length + 1, 1, newRows.length, 3).setValues(newRows);
      SpreadsheetApp.flush();
    }
  }

  return cats;
}

function saveGlobalCategories(categories) {
  const sheet = getCategoriesSheet();
  sheet.clear();
  sheet.appendRow(['id', 'name', 'profile']);
  if (categories && categories.length > 0) {
    const rows = categories.map(c => [c.id, c.name, c.profile]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  
  return true;
}

function addCategory(name, profile) {
  const sheet = getCategoriesSheet();
  const newId = new Date().getTime() + Math.floor(Math.random() * 1000);
  sheet.appendRow([newId, name, profile]);
  return true;
}

function updateCategory(id, name, profile) {
  const sheet = getCategoriesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) == String(id) || String(data[i][1]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      sheet.getRange(i + 1, 2).setValue(name);
      sheet.getRange(i + 1, 3).setValue(profile);
      SpreadsheetApp.flush();
      return true;
    }
  }
  sheet.appendRow([id || Date.now(), name, profile]);
  SpreadsheetApp.flush();
  return true;
}

function deleteCategory(name) {
  const sheet = getCategoriesSheet();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(name).trim().toLowerCase()) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
