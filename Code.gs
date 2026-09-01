/**
 * Tata's Batters — Cost Tracker web app (backend)
 * Paste this into Extensions ▸ Apps Script as "Code.gs"
 */

const SHEET_ING  = 'Ingredient Prices';
const SHEET_REC  = 'Recipe Costing';
const SHEET_PROD = 'Product Cost Summary';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle("Tata's Batters — Cost Tracker")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(name) {
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet not found: ' + name);
  return s;
}
function isNoteRow_(v) {
  const t = String(v || '');
  return t.indexOf('Quick conversions') === 0 ||
         t.indexOf('Total Ingredient Cost') === 0;
}

/* ============================ READ ============================ */

function getData() {
  const ing  = sh_(SHEET_ING);
  const prod = sh_(SHEET_PROD);
  const rec  = sh_(SHEET_REC);

  // ---- ingredients ----
  const ingredients = [];
  const iLast = ing.getLastRow();
  if (iLast >= 2) {
    const v = ing.getRange(2, 1, iLast - 1, 11).getValues();
    v.forEach(function (r, i) {
      const name = String(r[0] || '').trim();
      if (!name || isNoteRow_(name)) return;
      ingredients.push({
        row: i + 2, name: name, ar: r[1] || '', type: r[2] || '',
        size: r[3] || '', price: r[4], qty: r[5], unit: r[6] || '',
        costPerUnit: r[7], supplier: r[8] || '', notes: r[10] || ''
      });
    });
  }

  // ---- products ----
  const products = [];
  const pLast = prod.getLastRow();
  if (pLast >= 2) {
    const v = prod.getRange(2, 1, pLast - 1, 12).getValues();
    v.forEach(function (r, i) {
      const name = String(r[0] || '').trim();
      if (!name || isNoteRow_(name)) return;
      products.push({
        row: i + 2, name: name, ar: r[1] || '', yieldQty: r[2],
        yieldUnit: r[3] || '', totalCost: r[4], costPerUnit: r[5],
        packaging: r[6], labor: r[7], totalPerUnit: r[8],
        price: r[9], margin: r[10], status: r[11] || ''
      });
    });
  }

  // ---- recipe lines (only rows that actually have an ingredient) ----
  const lines = [];
  const rLast = rec.getLastRow();
  if (rLast >= 2) {
    const v = rec.getRange(2, 1, rLast - 1, 10).getValues();
    v.forEach(function (r, i) {
      const ingName = String(r[2] || '').trim();
      if (!ingName) return;
      lines.push({
        row: i + 2, product: String(r[9] || '').trim(), ingredient: ingName,
        qty: r[4], unit: r[5] || '', cost: r[7], notes: r[8] || ''
      });
    });
  }

  // ---- ingredient types already in use (for the dropdown) ----
  const types = [];
  ingredients.forEach(function (x) {
    if (x.type && types.indexOf(x.type) === -1) types.push(x.type);
  });
  const DEFAULT_TYPES = ['Flour & Grains','Prepared Doughs','Dairy & Eggs','Nuts & Seeds',
    'Dried Fruit','Produce','Meat & Poultry','Spices & Herbs','Sweeteners','Fats & Oils',
    'Leavening','Liquids & Extracts','Baking Add-ins','Packaging','Other'];
  DEFAULT_TYPES.forEach(function (t) { if (types.indexOf(t) === -1) types.push(t); });
  types.sort();

  return { ingredients: ingredients, products: products, lines: lines, types: types };
}

/* ======================= ADD AN INGREDIENT ======================= */

function addIngredient(d) {
  const sheet = sh_(SHEET_ING);
  const name = String(d.name || '').trim();
  if (!name) throw new Error('Ingredient name is required.');

  // reject duplicates
  const existing = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
      throw new Error('"' + name + '" is already in your ingredient list.');
    }
  }

  const row = firstBlankRow_(sheet, 1, 2);
  sheet.getRange(row, 1).setValue(name);
  sheet.getRange(row, 2).setValue(d.ar || '');
  sheet.getRange(row, 3).setValue(d.type || '');
  sheet.getRange(row, 4).setValue(d.size || '');
  sheet.getRange(row, 5).setValue(numOrBlank_(d.price));
  sheet.getRange(row, 6).setValue(numOrBlank_(d.qty));
  sheet.getRange(row, 7).setValue(d.unit || '');
  sheet.getRange(row, 8).setFormula('=IF(F' + row + '=0,0,E' + row + '/F' + row + ')');
  sheet.getRange(row, 9).setValue(d.supplier || '');
  sheet.getRange(row, 10).setValue(new Date());
  sheet.getRange(row, 11).setValue(d.notes || '');

  SpreadsheetApp.flush();
  return 'Added ingredient: ' + name;
}

/* ======================= UPDATE A PRICE ======================= */

function updateIngredientPrice(d) {
  const sheet = sh_(SHEET_ING);
  const row = findRowByValue_(sheet, 1, d.name);
  if (!row) throw new Error('Could not find ingredient: ' + d.name);

  if (d.price !== '' && d.price !== null && d.price !== undefined) {
    sheet.getRange(row, 5).setValue(Number(d.price));
  }
  if (d.qty !== '' && d.qty !== null && d.qty !== undefined) {
    sheet.getRange(row, 6).setValue(Number(d.qty));
  }
  if (d.size)     sheet.getRange(row, 4).setValue(d.size);
  if (d.unit)     sheet.getRange(row, 7).setValue(d.unit);
  if (d.supplier) sheet.getRange(row, 9).setValue(d.supplier);
  sheet.getRange(row, 10).setValue(new Date());

  // make sure the cost formula is intact
  sheet.getRange(row, 8).setFormula('=IF(F' + row + '=0,0,E' + row + '/F' + row + ')');

  SpreadsheetApp.flush();
  const cpu = sheet.getRange(row, 8).getValue();
  return 'Updated ' + d.name + ' — now $' + Number(cpu).toFixed(4) + ' per ' +
         sheet.getRange(row, 7).getValue();
}

/* ========================= ADD A PRODUCT ========================= */

function addProduct(d) {
  const prod = sh_(SHEET_PROD);
  const name = String(d.name || '').trim();
  if (!name) throw new Error('Product name is required.');

  const existing = prod.getRange(2, 1, Math.max(prod.getLastRow() - 1, 1), 1).getValues();
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i][0] || '').trim().toLowerCase() === name.toLowerCase()) {
      throw new Error('"' + name + '" already exists on the summary sheet.');
    }
  }

  const row = firstBlankRow_(prod, 1, 2);
  prod.getRange(row, 1).setValue(name);
  prod.getRange(row, 2).setValue(d.ar || '');
  prod.getRange(row, 3).setValue(numOrBlank_(d.yieldQty));
  prod.getRange(row, 4).setValue(d.yieldUnit || '');
  prod.getRange(row, 5).setFormula(
    "=SUMIF('" + SHEET_REC + "'!$J:$J,A" + row + ",'" + SHEET_REC + "'!$H:$H)");
  prod.getRange(row, 6).setFormula('=IF(C' + row + '=0,0,E' + row + '/C' + row + ')');
  prod.getRange(row, 7).setValue(numOrZero_(d.packaging));
  prod.getRange(row, 8).setValue(numOrZero_(d.labor));
  prod.getRange(row, 9).setFormula('=F' + row + '+G' + row + '+H' + row);
  prod.getRange(row, 10).setValue(numOrBlank_(d.price));
  prod.getRange(row, 11).setFormula(
    '=IF(OR(J' + row + '=0,E' + row + '=0),"—",(J' + row + '-I' + row + ')/J' + row + ')');
  prod.getRange(row, 12).setValue(d.status || 'Add recipe rows');

  // start a block for it on Recipe Costing so recipe lines have somewhere to go
  const rec = sh_(SHEET_REC);
  const blockRow = firstBlankRow_(rec, 10, 2);   // first row with no resolved product
  ensureRecipeFormulas_(rec, blockRow);
  rec.getRange(blockRow, 1).setValue(name);

  SpreadsheetApp.flush();
  return 'Added product: ' + name;
}

/* ====================== ADD A RECIPE LINE ====================== */

function addRecipeLine(d) {
  const rec = sh_(SHEET_REC);
  const product = String(d.product || '').trim();
  const ingredient = String(d.ingredient || '').trim();
  if (!product)    throw new Error('Pick a product.');
  if (!ingredient) throw new Error('Pick an ingredient.');

  const lastRow = Math.max(rec.getLastRow(), 2);
  const resolved = rec.getRange(2, 10, lastRow - 1, 1).getValues();  // col J
  const ingCol   = rec.getRange(2, 3,  lastRow - 1, 1).getValues();  // col C

  var firstFree = 0, blockEnd = 0;
  for (var i = 0; i < resolved.length; i++) {
    if (String(resolved[i][0] || '').trim() === product) {
      blockEnd = i + 2;
      if (!firstFree && !String(ingCol[i][0] || '').trim()) firstFree = i + 2;
    }
  }

  var targetRow;
  if (firstFree) {
    // reuse an empty row already inside this product's block
    targetRow = firstFree;
  } else if (blockEnd) {
    // block is full — insert a fresh row at the end of it
    rec.insertRowAfter(blockEnd);
    targetRow = blockEnd + 1;
    rec.getRange(blockEnd, 1, 1, 10).copyTo(rec.getRange(targetRow, 1, 1, 10));
    rec.getRange(targetRow, 1).clearContent();  // stay blank = inherit product
    rec.getRange(targetRow, 3).clearContent();
    rec.getRange(targetRow, 5).clearContent();
    rec.getRange(targetRow, 9).clearContent();
    ensureRecipeFormulas_(rec, targetRow);
  } else {
    // no block exists yet — start one at the first free row
    targetRow = firstBlankRow_(rec, 10, 2);
    ensureRecipeFormulas_(rec, targetRow);
    rec.getRange(targetRow, 1).setValue(product);
  }

  rec.getRange(targetRow, 3).setValue(ingredient);
  rec.getRange(targetRow, 5).setValue(numOrBlank_(d.qty));
  rec.getRange(targetRow, 9).setValue(d.notes || '');
  ensureRecipeFormulas_(rec, targetRow);

  SpreadsheetApp.flush();
  const cost = rec.getRange(targetRow, 8).getValue();
  return 'Added ' + ingredient + ' to ' + product + ' — $' + Number(cost).toFixed(2);
}

/* ==================== DELETE A RECIPE LINE ==================== */

function clearRecipeLine(row) {
  const rec = sh_(SHEET_REC);
  const r = Number(row);
  if (!r || r < 2) throw new Error('Bad row.');
  // clear only the typed cells; formulas and block structure stay intact
  rec.getRange(r, 3).clearContent();
  rec.getRange(r, 5).clearContent();
  rec.getRange(r, 9).clearContent();
  SpreadsheetApp.flush();
  return 'Removed that ingredient line.';
}

/* =========================== HELPERS =========================== */

function firstBlankRow_(sheet, col, startRow) {
  const last = Math.max(sheet.getLastRow(), startRow);
  const vals = sheet.getRange(startRow, col, last - startRow + 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    const v = String(vals[i][0] || '').trim();
    if (!v) return startRow + i;
    if (isNoteRow_(v)) return startRow + i;   // don't write past the note row
  }
  return last + 1;
}

function findRowByValue_(sheet, col, value) {
  const last = Math.max(sheet.getLastRow(), 2);
  const vals = sheet.getRange(2, col, last - 1, 1).getValues();
  const target = String(value || '').trim().toLowerCase();
  for (var i = 0; i < vals.length; i++) {
    if (String(vals[i][0] || '').trim().toLowerCase() === target) return i + 2;
  }
  return 0;
}

function ensureRecipeFormulas_(rec, row) {
  const prev = row - 1;
  rec.getRange(row, 10).setFormula(
    row === 2 ? '=IF(A2<>"",A2,"")'
              : '=IF(A' + row + '<>"",A' + row + ',J' + prev + ')');
  rec.getRange(row, 2).setFormula(
    '=IFERROR(VLOOKUP(J' + row + ",'" + SHEET_PROD + "'!$A:$L,2,0),\"\")");
  rec.getRange(row, 4).setFormula(
    '=IFERROR(VLOOKUP(C' + row + ",'" + SHEET_ING + "'!$A:$K,2,0),\"\")");
  rec.getRange(row, 6).setFormula(
    '=IFERROR(VLOOKUP(C' + row + ",'" + SHEET_ING + "'!$A:$K,7,0),\"\")");
  rec.getRange(row, 7).setFormula(
    '=IFERROR(VLOOKUP(C' + row + ",'" + SHEET_ING + "'!$A:$K,8,0),0)");
  rec.getRange(row, 8).setFormula(
    '=IF(C' + row + '="",0,E' + row + '*G' + row + ')');
}

function numOrBlank_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}
function numOrZero_(v) {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
}
