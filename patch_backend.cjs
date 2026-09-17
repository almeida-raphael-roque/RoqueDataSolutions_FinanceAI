const fs = require('fs');
let code = fs.readFileSync('gas-project/Transactions.gs', 'utf8');

// 1. Optimize getTransactions
code = code.replace(
`  const catNameMap = {};
  try {
    const cats = getGlobalCategories();
    cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);
  } catch(e) {}`,
`  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catNameMap = {};
  cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);`
);

// 2. Optimize getDashboardData
code = code.replace(
`  const catNameMap = {};
  try {
    const cats = getGlobalCategories();
    cats.forEach(c => catNameMap[String(c.name).toLowerCase()] = c.name);
  } catch(e) {}`,
`  let cats = [];
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
  });`
);

// 3. Optimize getPlanningData
const planningOld = `function getPlanningData(year) {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();`;

const planningNew = `function getPlanningData(year) {
  const sheet = getTable('Transactions');
  const data = sheet.getDataRange().getValues();
  
  let cats = [];
  try { cats = getGlobalCategories(); } catch(e) {}
  const catMacroMap = {};
  cats.forEach(c => {
    if (c.profile) {
      if (c.profile.includes('Fixa')) catMacroMap[String(c.name).toLowerCase()] = 'Fixas';
      else if (c.profile.includes('Variável')) catMacroMap[String(c.name).toLowerCase()] = 'Variáveis Essenciais';
      else catMacroMap[String(c.name).toLowerCase()] = 'Discricionárias';
    }
  });`;

code = code.replace(planningOld, planningNew);

// 4. Update getMacroCategory to accept catMacroMap
const macroOld = `function getMacroCategory(categoria, descricao) {
  const catLower = String(categoria || '').toLowerCase().trim();
  const descLower = String(descricao || '').toLowerCase().trim();
  
  // 1. Check Global Categories (Single Source of Truth)
  try {
    const cats = getGlobalCategories();
    const existing = cats.find(c => String(c.name).toLowerCase().trim() === catLower);
    if (existing && existing.profile) {
      if (existing.profile.includes('Fixa')) return 'Fixas';
      if (existing.profile.includes('Variável')) return 'Variáveis Essenciais';
      return 'Discricionárias';
    }
  } catch(e) {}`;

const macroNew = `function getMacroCategory(categoria, descricao, catMacroMap) {
  const catLower = String(categoria || '').toLowerCase().trim();
  const descLower = String(descricao || '').toLowerCase().trim();
  
  // 1. Check Global Categories (Single Source of Truth)
  if (catMacroMap && catMacroMap[catLower]) {
      return catMacroMap[catLower];
  }`;

code = code.replace(macroOld, macroNew);

// 5. Update calls to getMacroCategory in Transactions.gs
code = code.replace(/getMacroCategory\(([^,]+),\s*([^)]+)\)/g, 'getMacroCategory($1, $2, typeof catMacroMap !== "undefined" ? catMacroMap : null)');

fs.writeFileSync('gas-project/Transactions.gs', code);
console.log("Patched Transactions.gs");
