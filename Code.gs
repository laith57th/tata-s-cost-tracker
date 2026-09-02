/**
 * Tata's Batters — Cost Tracker web app (backend)
 * Paste this into Extensions ▸ Apps Script as "Code.gs"
 */

const SHEET_ING  = 'Ingredient Prices';
const SHEET_REC  = 'Recipe Costing';
const SHEET_PROD = 'Product Cost Summary';

/**
 * Only these Google accounts can open the app. Add/remove emails here —
 * no redeploy needed for the list itself to take effect on next load,
 * though code CHANGES still require a redeploy (see SETUP.md).
 */
var ALLOWED_EMAILS = [
  'laith57th@gmail.com',
  'leenhussin25@gmail.com'
];

function doGet() {
  var email = Session.getActiveUser().getEmail();

  if (!email) {
    // Deployment "Who has access" isn't set to require login — see SETUP.md.
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif;padding:40px;text-align:center;color:#333">' +
      '<h2>Sign-in required</h2>' +
      '<p>This app needs to know who you are. Make sure you\'re signed into a ' +
      'Google account in this browser, then reload the page.</p></body>');
  }

  if (ALLOWED_EMAILS.indexOf(email.toLowerCase()) === -1 &&
      ALLOWED_EMAILS.indexOf(email) === -1) {
    return HtmlService.createHtmlOutput(
      '<body style="font-family:sans-serif;padding:40px;text-align:center;color:#333">' +
      '<h2>Not authorized</h2>' +
      '<p>' + email + ' doesn\'t have access to this tool.</p>' +
      '<p style="color:#888;font-size:13px">If this should be allowed, add this email ' +
      'to ALLOWED_EMAILS in the script.</p></body>');
  }

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
    const v = ing.getRange(2, 1, iLast - 1, 10).getValues();
    v.forEach(function (r, i) {
      const name = String(r[0] || '').trim();
      if (!name || isNoteRow_(name)) return;
      ingredients.push({
        row: i + 2, name: name, type: r[1] || '',
        size: r[2] || '', price: r[3], qty: r[4], unit: r[5] || '',
        costPerUnit: r[6], supplier: r[7] || '', notes: r[9] || ''
      });
    });
  }

  // ---- products ----
  const products = [];
  const pLast = prod.getLastRow();
  if (pLast >= 2) {
    const v = prod.getRange(2, 1, pLast - 1, 11).getValues();
    v.forEach(function (r, i) {
      const name = String(r[0] || '').trim();
      if (!name || isNoteRow_(name)) return;
      products.push({
        row: i + 2, name: name, yieldQty: r[1],
        yieldUnit: r[2] || '', totalCost: r[3], costPerUnit: r[4],
        packaging: r[5], labor: r[6], totalPerUnit: r[7],
        price: r[8], margin: r[9], status: r[10] || ''
      });
    });
  }

  // ---- recipe lines (only rows that actually have an ingredient) ----
  const lines = [];
  const rLast = rec.getLastRow();
  if (rLast >= 2) {
    const v = rec.getRange(2, 1, rLast - 1, 8).getValues();
    v.forEach(function (r, i) {
      const ingName = String(r[1] || '').trim();
      if (!ingName) return;
      lines.push({
        row: i + 2, product: String(r[7] || '').trim(), ingredient: ingName,
        qty: r[2], unit: r[3] || '', cost: r[5], notes: r[6] || ''
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
  sheet.getRange(row, 2).setValue(d.type || '');
  sheet.getRange(row, 3).setValue(d.size || '');
  sheet.getRange(row, 4).setValue(numOrBlank_(d.price));
  sheet.getRange(row, 5).setValue(numOrBlank_(d.qty));
  sheet.getRange(row, 6).setValue(d.unit || '');
  sheet.getRange(row, 7).setFormula('=IF(E' + row + '=0,0,D' + row + '/E' + row + ')');
  sheet.getRange(row, 8).setValue(d.supplier || '');
  sheet.getRange(row, 9).setValue(new Date());
  sheet.getRange(row, 10).setValue(d.notes || '');

  SpreadsheetApp.flush();
  return 'Added ingredient: ' + name;
}

/* ======================= UPDATE A PRICE ======================= */

function updateIngredientPrice(d) {
  const sheet = sh_(SHEET_ING);
  const row = findRowByValue_(sheet, 1, d.name);
  if (!row) throw new Error('Could not find ingredient: ' + d.name);

  if (d.price !== '' && d.price !== null && d.price !== undefined) {
    sheet.getRange(row, 4).setValue(Number(d.price));
  }
  if (d.qty !== '' && d.qty !== null && d.qty !== undefined) {
    sheet.getRange(row, 5).setValue(Number(d.qty));
  }
  if (d.size)     sheet.getRange(row, 3).setValue(d.size);
  if (d.unit)     sheet.getRange(row, 6).setValue(d.unit);
  if (d.supplier) sheet.getRange(row, 8).setValue(d.supplier);
  sheet.getRange(row, 9).setValue(new Date());

  // make sure the cost formula is intact
  sheet.getRange(row, 7).setFormula('=IF(E' + row + '=0,0,D' + row + '/E' + row + ')');

  SpreadsheetApp.flush();
  const cpu = sheet.getRange(row, 7).getValue();
  return 'Updated ' + d.name + ' — now $' + Number(cpu).toFixed(4) + ' per ' +
         sheet.getRange(row, 6).getValue();
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
  prod.getRange(row, 2).setValue(numOrBlank_(d.yieldQty));
  prod.getRange(row, 3).setValue(d.yieldUnit || '');
  prod.getRange(row, 4).setFormula(
    "=SUMIF('" + SHEET_REC + "'!$H:$H,A" + row + ",'" + SHEET_REC + "'!$F:$F)");
  prod.getRange(row, 5).setFormula('=IF(B' + row + '=0,0,D' + row + '/B' + row + ')');
  prod.getRange(row, 6).setValue(numOrZero_(d.packaging));
  prod.getRange(row, 7).setValue(numOrZero_(d.labor));
  prod.getRange(row, 8).setFormula('=E' + row + '+F' + row + '+G' + row);
  prod.getRange(row, 9).setValue(numOrBlank_(d.price));
  prod.getRange(row, 10).setFormula(
    '=IF(OR(I' + row + '=0,D' + row + '=0),"—",(I' + row + '-H' + row + '*B' + row + ')/I' + row + ')');
  prod.getRange(row, 11).setValue(d.status || 'Add recipe rows');

  // start a block for it on Recipe Costing so recipe lines have somewhere to go
  const rec = sh_(SHEET_REC);
  const blockRow = firstBlankRow_(rec, 8, 2);   // first row with no resolved product
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
  const resolved = rec.getRange(2, 8, lastRow - 1, 1).getValues();  // col H
  const ingCol   = rec.getRange(2, 2, lastRow - 1, 1).getValues();  // col B

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
    rec.getRange(blockEnd, 1, 1, 8).copyTo(rec.getRange(targetRow, 1, 1, 8));
    rec.getRange(targetRow, 1).clearContent();  // stay blank = inherit product
    rec.getRange(targetRow, 2).clearContent();
    rec.getRange(targetRow, 3).clearContent();
    rec.getRange(targetRow, 7).clearContent();
    ensureRecipeFormulas_(rec, targetRow);
  } else {
    // no block exists yet — start one at the first free row
    targetRow = firstBlankRow_(rec, 8, 2);
    ensureRecipeFormulas_(rec, targetRow);
    rec.getRange(targetRow, 1).setValue(product);
  }

  rec.getRange(targetRow, 2).setValue(ingredient);
  rec.getRange(targetRow, 3).setValue(numOrBlank_(d.qty));
  rec.getRange(targetRow, 7).setValue(d.notes || '');
  ensureRecipeFormulas_(rec, targetRow);

  SpreadsheetApp.flush();
  const cost = rec.getRange(targetRow, 6).getValue();
  return 'Added ' + ingredient + ' to ' + product + ' — $' + Number(cost).toFixed(2);
}

/* ==================== DELETE A RECIPE LINE ==================== */

function clearRecipeLine(row) {
  const rec = sh_(SHEET_REC);
  const r = Number(row);
  if (!r || r < 2) throw new Error('Bad row.');
  // clear only the typed cells; formulas and block structure stay intact
  rec.getRange(r, 2).clearContent();
  rec.getRange(r, 3).clearContent();
  rec.getRange(r, 7).clearContent();
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
  rec.getRange(row, 8).setFormula(
    row === 2 ? '=IF(A2<>"",A2,"")'
              : '=IF(A' + row + '<>"",A' + row + ',H' + prev + ')');
  rec.getRange(row, 4).setFormula(
    '=IFERROR(VLOOKUP(B' + row + ",'" + SHEET_ING + "'!$A:$J,6,0),\"\")");
  rec.getRange(row, 5).setFormula(
    '=IFERROR(VLOOKUP(B' + row + ",'" + SHEET_ING + "'!$A:$J,7,0),0)");
  rec.getRange(row, 6).setFormula(
    '=IF(B' + row + '="",0,C' + row + '*E' + row + ')');
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


/* ================================================================
 *  ORDERS  /  INCOME & EXPENSE
 * ================================================================ */

const SHEET_ORD = 'Orders';
const SHEET_ITM = 'Order Items';
const SHEET_FIN = 'Income & Expense';

/** Next sequential order ID, e.g. TB-0007 */
function nextOrderId_() {
  const s = sh_(SHEET_ORD);
  const last = Math.max(s.getLastRow(), 1);
  var max = 0;
  if (last >= 2) {
    const ids = s.getRange(2, 1, last - 1, 1).getValues();
    ids.forEach(function (r) {
      const m = String(r[0] || '').match(/TB-(\d+)/i);
      if (m) max = Math.max(max, parseInt(m[1], 10));
    });
  }
  return 'TB-' + ('0000' + (max + 1)).slice(-4);
}

function getOrders() {
  const so = sh_(SHEET_ORD), si = sh_(SHEET_ITM);

  const orders = [];
  const oLast = so.getLastRow();
  if (oLast >= 2) {
    const v = so.getRange(2, 1, oLast - 1, 24).getValues();
    v.forEach(function (r, i) {
      const id = String(r[0] || '').trim();
      if (!id) return;
      orders.push({
        row: i + 2, id: id, dateReceived: fmtDate_(r[1]), customer: r[2] || '',
        contact: r[3] || '', channel: r[4] || '', fulfillment: r[5] || '',
        neededDate: fmtDate_(r[6]), neededTime: r[7] || '', address: r[8] || '',
        status: r[9] || '', payStatus: r[10] || '', payMethod: r[11] || '',
        subtotal: r[12], deliveryFee: r[13], discount: r[14], total: r[15],
        deposit: r[16], balance: r[17], cost: r[18], profit: r[19],
        margin: r[20], allergy: r[21] || '', notes: r[22] || '',
        logged: String(r[23] || '').trim()
      });
    });
  }

  const items = [];
  const iLast = si.getLastRow();
  if (iLast >= 2) {
    const v = si.getRange(2, 1, iLast - 1, 9).getValues();
    v.forEach(function (r, i) {
      const prodName = String(r[1] || '').trim();
      if (!prodName) return;                       // skip empty rows
      const id = String(r[8] || r[0] || '').trim(); // resolved col I, fallback col A
      if (!id) return;
      items.push({
        row: i + 2, orderId: id, product: prodName, qty: r[2],
        unitPrice: r[3], lineTotal: r[4], lineCost: r[6], notes: r[7] || ''
      });
    });
  }

  return { orders: orders, items: items, nextId: nextOrderId_() };
}

function fmtDate_(d) {
  if (!d) return '';
  if (Object.prototype.toString.call(d) === '[object Date]') {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(d);
}

/** Create an order plus its line items in one go. */
function createOrder(d) {
  const so = sh_(SHEET_ORD), si = sh_(SHEET_ITM);
  if (!d.customer || !String(d.customer).trim()) throw new Error('Customer name is required.');
  if (!d.items || !d.items.length) throw new Error('Add at least one item to the order.');

  const id = nextOrderId_();
  const row = firstBlankRow_(so, 1, 2);

  so.getRange(row, 1).setValue(id);
  so.getRange(row, 2).setValue(d.dateReceived ? new Date(d.dateReceived) : new Date());
  so.getRange(row, 3).setValue(d.customer);
  so.getRange(row, 4).setValue(d.contact || '');
  so.getRange(row, 5).setValue(d.channel || '');
  so.getRange(row, 6).setValue(d.fulfillment || '');
  if (d.neededDate) so.getRange(row, 7).setValue(new Date(d.neededDate));
  so.getRange(row, 8).setValue(d.neededTime || '');
  so.getRange(row, 9).setValue(d.address || '');
  so.getRange(row, 10).setValue(d.status || 'New');
  so.getRange(row, 11).setValue(d.payStatus || 'Unpaid');
  so.getRange(row, 12).setValue(d.payMethod || '');
  so.getRange(row, 14).setValue(numOrZero_(d.deliveryFee));
  so.getRange(row, 15).setValue(numOrZero_(d.discount));
  so.getRange(row, 17).setValue(numOrZero_(d.deposit));
  so.getRange(row, 22).setValue(d.allergy || '');
  so.getRange(row, 23).setValue(d.notes || '');
  ensureOrderFormulas_(so, row);

  var firstLine = true;
  d.items.forEach(function (it) {
    if (!it.product) return;
    const ir = firstBlankRow_(si, 2, 2);        // first row with no Product
    if (firstLine) {                            // ID shown once per order only
      si.getRange(ir, 1).setValue(id);
      firstLine = false;
    }
    si.getRange(ir, 2).setValue(it.product);
    si.getRange(ir, 3).setValue(numOrZero_(it.qty));
    si.getRange(ir, 8).setValue(it.notes || '');
    ensureItemFormulas_(si, ir);
  });

  SpreadsheetApp.flush();
  const total = so.getRange(row, 16).getValue();
  return { id: id, message: 'Created ' + id + ' — ' + d.customer + ' — $' + Number(total).toFixed(2) };
}

function updateOrderStatus(d) {
  const so = sh_(SHEET_ORD);
  const row = findRowByValue_(so, 1, d.id);
  if (!row) throw new Error('Order not found: ' + d.id);
  if (d.status)    so.getRange(row, 10).setValue(d.status);
  if (d.payStatus) so.getRange(row, 11).setValue(d.payStatus);
  if (d.payMethod) so.getRange(row, 12).setValue(d.payMethod);
  if (d.deposit !== '' && d.deposit !== null && d.deposit !== undefined) {
    so.getRange(row, 17).setValue(Number(d.deposit));
  }
  SpreadsheetApp.flush();
  return 'Updated ' + d.id + '.';
}

/**
 * Post an order's revenue to the Income & Expense log.
 * Guarded so the same order can't be logged twice.
 */
function logOrderToIncome(orderId) {
  const so = sh_(SHEET_ORD), wf = sh_(SHEET_FIN);
  const row = findRowByValue_(so, 1, orderId);
  if (!row) throw new Error('Order not found: ' + orderId);

  if (String(so.getRange(row, 24).getValue() || '').trim().toLowerCase() === 'yes') {
    throw new Error(orderId + ' has already been logged to Income & Expense.');
  }

  const customer = so.getRange(row, 3).getValue();
  const total    = Number(so.getRange(row, 16).getValue()) || 0;
  const method   = so.getRange(row, 12).getValue();
  if (total <= 0) throw new Error('Order total is $0 — nothing to log.');

  const fr = firstBlankRow_(wf, 1, 2);
  wf.getRange(fr, 1).setValue(new Date());
  wf.getRange(fr, 2).setValue('Income');
  wf.getRange(fr, 3).setValue('Order Revenue');
  wf.getRange(fr, 4).setValue('Order ' + orderId + ' — ' + customer);
  wf.getRange(fr, 5).setValue(total);
  wf.getRange(fr, 6).setValue(method || '');
  wf.getRange(fr, 7).setValue(orderId);

  so.getRange(row, 24).setValue('Yes');
  so.getRange(row, 11).setValue('Paid in Full');
  SpreadsheetApp.flush();
  return 'Logged $' + total.toFixed(2) + ' from ' + orderId + ' to Income & Expense.';
}

function addExpense(d) {
  const wf = sh_(SHEET_FIN);
  if (!d.amount) throw new Error('Enter an amount.');
  const fr = firstBlankRow_(wf, 1, 2);
  wf.getRange(fr, 1).setValue(d.date ? new Date(d.date) : new Date());
  wf.getRange(fr, 2).setValue(d.type || 'Expense');
  wf.getRange(fr, 3).setValue(d.category || 'Other Expense');
  wf.getRange(fr, 4).setValue(d.description || '');
  wf.getRange(fr, 5).setValue(Number(d.amount));
  wf.getRange(fr, 6).setValue(d.method || '');
  wf.getRange(fr, 8).setValue(d.notes || '');
  SpreadsheetApp.flush();
  return 'Logged ' + (d.type || 'Expense') + ': $' + Number(d.amount).toFixed(2);
}

function getFinance() {
  const wf = sh_(SHEET_FIN);
  const rows = [];
  const last = wf.getLastRow();
  if (last >= 2) {
    const v = wf.getRange(2, 1, last - 1, 8).getValues();
    v.forEach(function (r, i) {
      if (!r[1] && !r[4]) return;
      rows.push({ row: i + 2, date: fmtDate_(r[0]), type: r[1] || '', category: r[2] || '',
                  description: r[3] || '', amount: r[4], method: r[5] || '',
                  orderId: r[6] || '', notes: r[7] || '' });
    });
  }
  var inc = 0, exp = 0;
  rows.forEach(function (r) {
    const a = Number(r.amount) || 0;
    if (String(r.type).toLowerCase() === 'income') inc += a; else exp += a;
  });
  return { rows: rows, totalIncome: inc, totalExpense: exp, net: inc - exp };
}

function ensureOrderFormulas_(so, row) {
  so.getRange(row, 13).setFormula("=SUMIF('" + SHEET_ITM + "'!$I:$I,A" + row + ",'" + SHEET_ITM + "'!$E:$E)");
  so.getRange(row, 16).setFormula('=M' + row + '+N' + row + '-O' + row);
  so.getRange(row, 18).setFormula('=P' + row + '-Q' + row);
  so.getRange(row, 19).setFormula("=SUMIF('" + SHEET_ITM + "'!$I:$I,A" + row + ",'" + SHEET_ITM + "'!$G:$G)");
  so.getRange(row, 20).setFormula('=P' + row + '-S' + row);
  so.getRange(row, 21).setFormula('=IF(P' + row + '=0,"—",T' + row + '/P' + row + ')');
}

function ensureItemFormulas_(si, row) {
  const prev = row - 1;
  // resolved Order ID: blank col A means "same order as the row above"
  si.getRange(row, 9).setFormula(
    row === 2 ? '=IF(A2<>"",A2,"")'
              : '=IF(A' + row + '<>"",A' + row + ',I' + prev + ')');
  si.getRange(row, 4).setFormula(
    "=IFERROR(VLOOKUP(B" + row + ",'" + SHEET_PROD + "'!$A:$I,9,0),0)");
  si.getRange(row, 5).setFormula('=IF(B' + row + '="",0,C' + row + '*D' + row + ')');
  si.getRange(row, 6).setFormula(
    "=IFERROR(VLOOKUP(B" + row + ",'" + SHEET_PROD + "'!$A:$H,8,0)*VLOOKUP(B" + row +
    ",'" + SHEET_PROD + "'!$A:$B,2,0),0)");
  si.getRange(row, 7).setFormula('=IF(B' + row + '="",0,C' + row + '*F' + row + ')');
}


