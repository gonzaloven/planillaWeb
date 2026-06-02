// ═══════════════════════════════════════════════
//  CONFIG — edit category lists here
// ═══════════════════════════════════════════════
const ESSENTIAL_CATS = [
  'Supermercado','Servicio','Expensas','Salud e Higiene','Monotributo',
  'Internet','Gimnasio/Deporte','Pagos y Honorarios','Hogar','Alquiler','Peluqueria'
];
const NON_ESSENTIAL_CATS = [
  'Bolucompra','Entretenimiento','Transporte','Ropa','Tecnologia','Vacaciones','Comida'
];

// ═══════════════════════════════════════════════
//  AUTH
// ═══════════════════════════════════════════════
function verifyAuth(e) {
  const username = e.parameter.username;
  const password = e.parameter.password;
  if (!username || !password) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Missing Authentication Credentials" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  const validUsername = "gonzaloven";
  const validPassword = "Aezakmibaguvix11!";
  if (username !== validUsername || password !== validPassword) {
    return ContentService.createTextOutput(JSON.stringify({ error: "Unauthorized" }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return null;
}

// ═══════════════════════════════════════════════
//  doGet — routes all GET actions
// ═══════════════════════════════════════════════
function doGet(e) {
  const authCheck = verifyAuth(e);
  if (authCheck) return authCheck;

  const action = e.parameter.action;

  if (action === "getData")      return getData();
  if (action === "getHistory")   return getHistory(e);
  if (action === "getPortfolio") return getPortfolioData();

  return json({ error: "Invalid action" });
}

// ═══════════════════════════════════════════════
//  doPost — routes all POST actions
// ═══════════════════════════════════════════════
function doPost(e) {
  const authCheck = verifyAuth(e);
  if (authCheck) return authCheck;

  const params = JSON.parse(e.postData.contents);
  const action = params.action;

  if (action === "addExpense")   return addExpense(params);
  if (action === "logPortfolio") return logPortfolio(params);
  if (action === "deleteEntry")  return deleteEntry(params);
  if (action === "editEntry")    return editEntry(params);

  return json({ error: "Invalid action" });
}

// ═══════════════════════════════════════════════
//  getData — dashboard data (extended)
// ═══════════════════════════════════════════════
function getData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Principal");
  const db    = ss.getSheetByName("Database");

  const dolar   = parseFloat(sheet.getRange("A3").getValue()) || 0;
  const balance = parseFloat(sheet.getRange("D3").getValue()) || 0;

  // Categories from data validation
  const categorias = getCategorias(sheet);

  // Salary: last Sueldo entry in Database
  const dbData = db.getDataRange().getValues();
  let salaryUSD = 0;
  for (let i = dbData.length - 1; i >= 1; i--) {
    if (dbData[i][1] === "Sueldo") { salaryUSD = parseFloat(dbData[i][3]) || 0; break; }
  }

  // This month's spend (USD, excluding income and Correccion)
  const now = new Date();
  const y = now.getFullYear(), m = now.getMonth();
  let thisMonthUSD = 0, thisMonthCount = 0;
  dbData.forEach((row, i) => {
    if (i === 0) return;
    const d = new Date(row[0]);
    if (d.getFullYear() !== y || d.getMonth() !== m) return;
    if (row[1] === "Sueldo" || row[1] === "Correccion") return;
    thisMonthUSD += parseFloat(row[3]) || 0;
    thisMonthCount++;
  });

  // Speedometer: budget = 20% of salary, minus non-essential this month
  const budgetMax = salaryUSD * 0.20;
  let nonEssentialSpent = 0;
  dbData.forEach((row, i) => {
    if (i === 0) return;
    const d = new Date(row[0]);
    if (d.getFullYear() !== y || d.getMonth() !== m) return;
    if (NON_ESSENTIAL_CATS.includes(row[1])) {
      nonEssentialSpent += parseFloat(row[3]) || 0;
    }
  });
  const speedometerValue = Math.max(0, budgetMax - nonEssentialSpent);

  return json({
    dolar,
    balance,
    categorias,
    salaryUSD,
    thisMonthUSD,
    thisMonthCount,
    speedometerValue,
    speedometerMax: budgetMax
  });
}

// ═══════════════════════════════════════════════
//  getHistory — paginated transaction list
// ═══════════════════════════════════════════════
function getHistory(e) {
  const db    = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  const rows  = db.getDataRange().getValues();

  // Optional filters from query params
  const filterCat   = e.parameter.category || '';
  const filterYear  = parseInt(e.parameter.year)  || 0;
  const filterMonth = parseInt(e.parameter.month) || 0; // 1-based

  const transactions = [];
  for (let i = rows.length - 1; i >= 1; i--) {
    const row = rows[i];
    const date = row[0], cat = row[1], ars = row[2], usd = row[3];
    if (!date || !cat) continue;

    if (filterCat && cat !== filterCat) continue;
    if (filterYear || filterMonth) {
      const d = new Date(date);
      if (filterYear  && d.getFullYear() !== filterYear)  continue;
      if (filterMonth && (d.getMonth()+1) !== filterMonth) continue;
    }

    transactions.push([
      Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss"),
      cat,
      parseFloat(ars) || 0,
      parseFloat(usd) || 0
    ]);
  }

  return json({ transactions });
}

// ═══════════════════════════════════════════════
//  addExpense — append row to Database
// ═══════════════════════════════════════════════
function addExpense(params) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const db      = ss.getSheetByName("Database");
  const sheet   = ss.getSheetByName("Principal");
  const date    = new Date();
  const cat     = params.categoria;
  const value   = parseFloat(params.valor);
  const dolar   = parseFloat(sheet.getRange("A3").getValue()) || 1;
  const usd     = cat === "Sueldo" ? value : value / dolar;
  const bgColor = cat === "Sueldo" ? "#77DD77" : "#FF6961";

  db.appendRow([date, cat, value, usd]);

  const lastRow = db.getLastRow();
  const range   = db.getRange(lastRow, 1, 1, 4);
  range.setBackground(bgColor);
  range.setFontSize(12);
  range.setFontFamily("Google Sans");

  return json({ success: true });
}

// ═══════════════════════════════════════════════
//  Portfolio snapshots
// ═══════════════════════════════════════════════
function getPortfolioData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let sheet   = ss.getSheetByName("Portfolio");
  if (!sheet) return json({ snapshots: [] });
  const rows  = sheet.getDataRange().getValues();
  const snapshots = rows.slice(1).map(r => [
    Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
    parseFloat(r[1]) || 0
  ]);
  return json({ snapshots });
}

// ═══════════════════════════════════════════════
//  deleteEntry — remove a row from Database by ISO date key
// ═══════════════════════════════════════════════
function deleteEntry(params) {
  const db   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Database");
  const rows = db.getDataRange().getValues();
  const tz   = Session.getScriptTimeZone();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (!rows[i][0]) continue;
    const dateStr = Utilities.formatDate(new Date(rows[i][0]), tz, "yyyy-MM-dd'T'HH:mm:ss");
    if (dateStr === params.rowDate) {
      db.deleteRow(i + 1);
      return json({ success: true });
    }
  }
  return json({ error: "Row not found" });
}

// ═══════════════════════════════════════════════
//  editEntry — update category + amount for a row
// ═══════════════════════════════════════════════
function editEntry(params) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const db    = ss.getSheetByName("Database");
  const sheet = ss.getSheetByName("Principal");
  const dolar = parseFloat(sheet.getRange("A3").getValue()) || 1;
  const rows  = db.getDataRange().getValues();
  const tz    = Session.getScriptTimeZone();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (!rows[i][0]) continue;
    const dateStr = Utilities.formatDate(new Date(rows[i][0]), tz, "yyyy-MM-dd'T'HH:mm:ss");
    if (dateStr === params.rowDate) {
      const cat = params.categoria;
      const val = parseFloat(params.valor);
      const usd = cat === "Sueldo" ? val : val / dolar;
      db.getRange(i + 1, 2).setValue(cat);
      db.getRange(i + 1, 3).setValue(val);
      db.getRange(i + 1, 4).setValue(usd);
      db.getRange(i + 1, 1, 1, 4).setBackground(cat === "Sueldo" ? "#77DD77" : "#FF6961");
      return json({ success: true });
    }
  }
  return json({ error: "Row not found" });
}

function logPortfolio(params) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Portfolio");
  if (!sheet) {
    sheet = ss.insertSheet("Portfolio");
    sheet.appendRow(["Fecha", "Valor USD"]);
  }
  sheet.appendRow([new Date(), parseFloat(params.value) || 0]);
  return json({ success: true });
}

// ═══════════════════════════════════════════════
//  Helpers
// ═══════════════════════════════════════════════
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getCategorias(sheet) {
  const categoriaRange = sheet.getRange("G3");
  const validation     = categoriaRange.getDataValidation();
  if (!validation) return [];
  const rule = validation.getCriteriaValues();
  return (rule && rule.length > 0) ? rule[0] : [];
}

// Keep legacy functions working
function generateEntry() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const db    = ss.getSheetByName("Database");
  const sheet = ss.getSheetByName("Principal");
  const date  = new Date();
  const cat   = sheet.getRange("G3").getValue();
  const value = sheet.getRange("F3").getValue();
  const dolar = parseInt(sheet.getRange("A3").getValue());
  if (value === "") return;
  const usd     = cat === "Sueldo" ? value : value / dolar;
  const bgColor = cat === "Sueldo" ? "#77DD77" : "#FF6961";
  db.appendRow([date, cat, value, usd]);
  const lastRow = db.getLastRow();
  const range   = db.getRange(lastRow, 1, 1, 4);
  range.setBackground(bgColor); range.setFontSize(12); range.setFontFamily("Google Sans");
  sheet.getRange("F3").setValue("0");
}

function hide() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Principal");
  const active = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  if (active.getValue() === true) {
    sheet.getRange("D2").setValue("***************");
  } else {
    sheet.getRange("D2").setFormula("=D3");
  }
}

function IMPORTJSON(url, query) {
  const response = UrlFetchApp.fetch(url);
  const data     = JSON.parse(response.getContentText());
  const parts    = query.split("/");
  let result     = data;
  for (let i = 1; i < parts.length; i++) result = result[parts[i]];
  return [[result]];
}

function onEdit() {
  const cell  = SpreadsheetApp.getActiveSpreadsheet().getActiveCell();
  const ref   = cell.getA1Notation();
  const sheet = cell.getSheet().getName();
  const val   = cell.getValue();
  if (ref === "I2" && sheet === "Principal" && val === true) generateEntry();
  if (ref === "E2" && sheet === "Principal") hide();
}
