// ================== CONFIG ==================
const SS_ID = '1CbAUtD12wYFXdx5ltTY699_KSdkn1497zwmoy48AJb8';

// Sheets
const INWARD_SHEET    = 'Inward BSM';
const INVENTORY_SHEET = 'Inventory';
const SALES_SHEET     = 'Sales';

const MASTER_SHEET = 'Master';
const LOG_SHEET    = 'Log';

// Drive folder for PDFs
const INVOICE_FOLDER_ID = '1TnDKdopwk0xUMqhgMvSJuhWsU8MksUSR';

// ================== UTILITIES ==================

// Get or create sheet
function getSheetByName_(name) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

// Master!A1 email
function getMasterEmail_() {
  const ss = SpreadsheetApp.openById(SS_ID);
  const sh = ss.getSheetByName(MASTER_SHEET);
  if (!sh) return '';
  const v = sh.getRange('A1').getValue();
  return String(v || '').trim();
}

// Log sheet append
function logEvent_(type, detail, meta) {
  const ss = SpreadsheetApp.openById(SS_ID);
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) sh = ss.insertSheet(LOG_SHEET);

  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'User', 'Type', 'Detail', 'Meta JSON']);
  }

  const userEmail = Session.getActiveUser().getEmail() || '';
  sh.appendRow([
    new Date(),
    userEmail,
    type,
    detail,
    meta ? JSON.stringify(meta) : ''
  ]);
}

// ===== INVENTORY HELPER (Inventory sheet structure) =====
// Inventory columns:
// A Date
// B Supplier
// C Ref No
// D BaleNo
// E Fabric Name
// F Fabric Quality
// G Design No
// H Color
// I One pcs Qty
// J Qty (pcs)
// K UOM
// L Rate   (₦ per Yard – from "Rate per Qty in Naira")
// M Total  (Qty Yard * Rate)
// N Location
// O Available Qty  (Yard)
const INV_COL = {
  BALE_NO:        4,  // D
  FABRIC_NAME:    5,  // E
  FABRIC_QUALITY: 6,  // F
  DESIGN_NO:      7,  // G
  COLOR:          8,  // H
  ONE_PCS_QTY:    9,  // I
  QTY:           10,  // J
  UOM:           11,  // K
  RATE:          12,  // L
  TOTAL:         13,  // M
  LOCATION:      14,  // N
  AVAIL_QTY:     15   // O
};

// ---------- SALES SHEET HEADER HELPER ----------
function getSalesSheet_() {
  const sh = getSheetByName_(SALES_SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Outward ID',        // 1  A
      'Date',              // 2  B
      'Customer Name',     // 3  C
      'Company Name',      // 4  D
      'Email',             // 5  E
      'Website',           // 6  F
      'VAT No',            // 7  G
      'Customer Image',    // 8  H
      'Invoice No',        // 9  I
      'Bale Nos',          // 10 J
      'Fabric Name',       // 11 K
      'Fabric Quality',    // 12 L
      'Design No',         // 13 M
      'Color',             // 14 N
      'Colour Varient',    // 15 O
      'One pcs Qty',       // 16 P
      'Qty',               // 17 Q (Yard sold)
      'UOM',               // 18 R
      'Inward Rate',       // 19 S (₦/Yard)
      'Selling Rate',      // 20 T
      'Line Amount',       // 21 U
      'Invoice PDF URL',   // 22 V
      'Created By',        // 23 W
      'Created At'         // 24 X
    ]);
  }
  return sh;
}

// Case-insensitive header lookup (0-based)
function findColIndex_(headerRow, candidates) {
  const upper = headerRow.map(c => String(c || '').trim().toUpperCase());
  for (let i = 0; i < upper.length; i++) {
    for (let j = 0; j < candidates.length; j++) {
      if (upper[i] === candidates[j].toUpperCase()) return i;
    }
  }
  return -1;
}

// Generate next invoice no: INV-YYYYMMDD-XXX
function nextInvoiceNo_() {
  const sh = getSalesSheet_();
  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idxInvoice = findColIndex_(headerRow, ['Invoice No']);
  if (idxInvoice < 0) throw new Error('Sales sheet header missing "Invoice No"');

  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const today  = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const prefix = 'INV-' + today + '-';

  const last = sh.getLastRow();
  let seq = 1;

  for (let r = last; r >= 2; r--) {
    const val = sh.getRange(r, idxInvoice + 1).getValue();
    if (!val) continue;
    const m = String(val).match(/INV-(\d{8})-(\d+)/);
    if (m && m[1] === today) {
      seq = Number(m[2]) + 1;
      break;
    }
  }

  return prefix + Utilities.formatString('%03d', seq);
}

// ============= 1. BUILD INVENTORY STATE (INWARD - SOLD) =============
function buildInventoryState_() {
  const invSh = getSheetByName_(INVENTORY_SHEET);
  const lastInv = invSh.getLastRow();
  if (lastInv < 2) return { rows: [] };

  // Read A:O for all inventory rows
  const invValues = invSh.getRange(2, 1, lastInv - 1, INV_COL.AVAIL_QTY).getValues();

  // ---- Build SOLD map (in Yard) from Sales ----
  const salesSh   = getSalesSheet_();
  const lastSales = salesSh.getLastRow();
  const soldMap   = {}; // key = bale|design|color => qty sold (Yard)

  if (lastSales >= 2) {
    const lastCol = salesSh.getLastColumn();
    const sHeader = salesSh.getRange(1, 1, 1, lastCol).getValues()[0];

    const idxBale   = findColIndex_(sHeader, ['Bale Nos', 'Bale No']);
    const idxDesign = findColIndex_(sHeader, ['Design No']);
    const idxColor  = findColIndex_(sHeader, ['Color']);
    const idxQty    = findColIndex_(sHeader, ['Qty (Sold)', 'Qty', 'Quantity']);

    if (idxBale >= 0 && idxDesign >= 0 && idxColor >= 0 && idxQty >= 0) {
      const sVals = salesSh.getRange(2, 1, lastSales - 1, lastCol).getValues();

      sVals.forEach(r => {
        const baleStr  = String(r[idxBale]   || '').trim();
        const designNo = String(r[idxDesign] || '').trim();
        const color    = String(r[idxColor]  || '').trim();
        const qtySoldY = Number(r[idxQty]    || 0);    // Yard

        if (!baleStr || !designNo || !color || !qtySoldY) return;

        baleStr
          .split(',')
          .map(b => b.trim())
          .filter(Boolean)
          .forEach(bale => {
            const key = `${bale}|${designNo}|${color}`;
            soldMap[key] = (soldMap[key] || 0) + qtySoldY;
          });
      });
    }
  }

  // ---- Build row objects + compute Available Qty (Yard) ----
  const rows         = [];
  const availToWrite = [];

  invValues.forEach((r, i) => {
    const rowIndex = i + 2;

    const baleNo        = String(r[INV_COL.BALE_NO        - 1] || '').trim();
    const fabricName    = String(r[INV_COL.FABRIC_NAME    - 1] || '').trim();
    const fabricQuality = String(r[INV_COL.FABRIC_QUALITY - 1] || '').trim();
    const designNo      = String(r[INV_COL.DESIGN_NO      - 1] || '').trim();
    const color         = String(r[INV_COL.COLOR          - 1] || '').trim();
    const onePcsQty     = Number(r[INV_COL.ONE_PCS_QTY    - 1] || 0); // Yard/pcs
    const qtyPcs        = Number(r[INV_COL.QTY            - 1] || 0); // pcs
    const uom           = String(r[INV_COL.UOM            - 1] || '').trim();
    const rate          = Number(r[INV_COL.RATE           - 1] || 0); // ₦/Yard
    const location      = String(r[INV_COL.LOCATION       - 1] || '').trim();

    const inwardYard = onePcsQty * qtyPcs;           // total yard inward
    const key        = `${baleNo}|${designNo}|${color}`;
    const soldYard   = soldMap[key] || 0;
    const availYard  = Math.max(0, inwardYard - soldYard);

    availToWrite.push([availYard]);

    rows.push({
      rowIndex,
      baleNo,
      fabricName,
      fabricQuality,
      designNo,
      color,
      onePcsQty,          // Yard per pcs
      qtyInwardYard: inwardYard,
      uom,
      inwardRate: rate,   // ₦ per Yard
      location,
      soldQtyYard: soldYard,
      availableQty: availYard
    });
  });

  // Write Available Qty (Yard) into column O
  invSh
    .getRange(2, INV_COL.AVAIL_QTY, availToWrite.length, 1)
    .setValues(availToWrite);

  return { rows };
}

// Optional: Bale-only sold map
function getBaleSoldMap_() {
  const sh   = getSheetByName_(SALES_SHEET);
  const last = sh.getLastRow();
  const soldMap = {};
  if (last < 2) return soldMap;

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  let idxBale = findColIndex_(header, ['Bale Nos', 'Bale No']);
  let idxQty  = findColIndex_(header, ['Qty (Sold)', 'Qty']);
  if (idxBale === -1 || idxQty === -1) return soldMap;

  const values = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  values.forEach(r => {
    const baleStr = String(r[idxBale] || '').trim();
    const qty     = Number(r[idxQty] || 0);
    if (!baleStr || !qty) return;

    baleStr
      .split(',')
      .map(b => b.trim())
      .filter(Boolean)
      .forEach(bale => {
        if (!soldMap[bale]) soldMap[bale] = 0;
        soldMap[bale] += qty;
      });
  });

  return soldMap;
}

// ================== 2. PENDING INWARD (Inward BSM) ==================
function getPendingInward() {
  const sh   = getSheetByName_(INWARD_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const lastCol = sh.getLastColumn();
  const header  = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values  = sh.getRange(2, 1, last - 1, lastCol).getValues();

  const idxBale    = header.indexOf('Bale No');
  const idxFabric  = header.indexOf('Fabric Name');
  const idxQuality = header.indexOf('Fabric Quality');
  const idxDesign  = header.indexOf('Design No');
  const idxQtyYard = header.indexOf('Qty'); // total Yard
  const idxUom     = header.indexOf('UOM');
  const idxStatus  = header.indexOf('Inward By Nigeria Status');

  if (idxBale === -1 || idxQtyYard === -1 || idxStatus === -1) {
    throw new Error(
      'Inward BSM sheet is missing Bale No / Qty / Inward By Nigeria Status headers'
    );
  }

  // Colour columns (White, Wine, Yellow, Color-1..Color-15, etc)
  const colorColIdxs = [];
  header.forEach((name, i) => {
    const n = String(name || '').trim();
    if (!n) return;
    if (['White', 'Wine', 'Yellow'].includes(n) || /^Color/i.test(n)) {
      colorColIdxs.push(i);
    }
  });

  const baleMap = {};

  values.forEach((row, rIndex) => {
    const status = String(row[idxStatus] || '').trim();
    if (status === 'Verified') return;

    const baleNo = String(row[idxBale] || '').trim();
    if (!baleNo) return;

    if (!baleMap[baleNo]) {
      baleMap[baleNo] = {
        baleNo,
        fabricName:    String(row[idxFabric]  || ''),
        fabricQuality: String(row[idxQuality] || ''),
        totalQtyYard:  0,
        uom:           String(row[idxUom]     || ''),
        designSet: new Set(),
        colorSet:  new Set(),
        rowIndexes: []
      };
    }

    const rec = baleMap[baleNo];

    rec.totalQtyYard += Number(row[idxQtyYard] || 0);

    if (idxDesign !== -1 && row[idxDesign] !== '') {
      rec.designSet.add(String(row[idxDesign]));
    }

    colorColIdxs.forEach(ci => {
      const q = Number(row[ci] || 0);
      if (q > 0) {
        const colourName = String(header[ci] || '').trim();
        if (colourName) rec.colorSet.add(colourName);
      }
    });

    rec.rowIndexes.push(rIndex + 2);
  });

  return Object.values(baleMap).map(rec => ({
    baleNo:        rec.baleNo,
    fabricName:    rec.fabricName,
    fabricQuality: rec.fabricQuality,
    totalQtyYard:  rec.totalQtyYard,
    designCount:   rec.designSet.size,
    colorCount:    rec.colorSet.size,
    rowIndexes:    rec.rowIndexes
  }));
}

// ================== 3. VERIFY BALE & PUSH TO INVENTORY ==================
// Uses "Rate per Qty in Naira" from Inward BSM as ₦ per Yard.
function verifyBaleRow(payload) {
  const baleNo = String(payload && payload.baleNo || '').trim();
  if (!baleNo) throw new Error('Missing baleNo');

  const ss     = SpreadsheetApp.openById(SS_ID);
  const inward = ss.getSheetByName(INWARD_SHEET);
  const inv    = getSheetByName_(INVENTORY_SHEET);

  const last = inward.getLastRow();
  if (last < 2) return { ok: false, message: 'No inward rows' };

  const lastCol = inward.getLastColumn();
  const header  = inward.getRange(1, 1, 1, lastCol).getValues()[0];
  const values  = inward.getRange(2, 1, last - 1, lastCol).getValues();

  // --- find columns using helper (case-insensitive, safer) ---
  const idxBale    = findColIndex_(header, ['Bale No']);
  const idxFabric  = findColIndex_(header, ['Fabric Name']);
  const idxQuality = findColIndex_(header, ['Fabric Quality']);
  const idxDesign  = findColIndex_(header, ['Design No']);
  const idxColor   = findColIndex_(header, ['Color']);
  const idxPcs     = findColIndex_(header, ['Pcs']);
  const idxQtyYard = findColIndex_(header, ['Qty']); // total Yard
  const idxUom     = findColIndex_(header, ['UOM']);
  const idxRateNg  = findColIndex_(header, [
    'Rate per Qty in Naira',
    'Rate per Qty in Naira (₦)',
    'Rate/Qty in Naira'
  ]);
  const idxStatus  = findColIndex_(header, ['Inward By Nigeria Status']);

  if (idxBale === -1 || idxPcs === -1 || idxQtyYard === -1 || idxStatus === -1) {
    throw new Error(
      'Inward BSM header missing one of: Bale No / Pcs / Qty / Inward By Nigeria Status'
    );
  }

  const rowsToVerify = [];
  values.forEach((row, i) => {
    const thisBale = String(row[idxBale] || '').trim();
    const st       = String(row[idxStatus] || '').trim();
    if (thisBale === baleNo && st !== 'Verified') {
      rowsToVerify.push({ sheetRow: i + 2, row });
    }
  });

  if (!rowsToVerify.length) {
    return { ok: true, message: 'Nothing to verify for bale ' + baleNo };
  }

  // Ensure Inventory header exists
  if (inv.getLastRow() === 0) {
    inv.appendRow([
      'Date', 'Supplier', 'Ref No', 'BaleNo',
      'Fabric Name', 'Fabric Quality', 'Design No', 'Color',
      'One pcs Qty', 'Qty', 'UOM', 'Rate', 'Total',
      'Location', 'Available Qty'
    ]);
  }

  const invRows = [];

  rowsToVerify.forEach(({ row }) => {
    const fabric  = idxFabric  >= 0 ? String(row[idxFabric]  || '') : '';
    const quality = idxQuality >= 0 ? String(row[idxQuality] || '') : '';
    const design  = idxDesign  >= 0 ? String(row[idxDesign]  || '') : '';
    const color   = idxColor   >= 0 ? String(row[idxColor]   || '') : '';
    const pcs     = Number(idxPcs     >= 0 ? (row[idxPcs]     || 0) : 0);   // pieces
    const qtyYard = Number(idxQtyYard >= 0 ? (row[idxQtyYard] || 0) : 0);   // total Yard
    const uom     = idxUom    >= 0 ? String(row[idxUom]      || 'Yard') : 'Yard';

    // 🔹 Treat "Rate per Qty in Naira" as Naira per Yard
    const rateNg  = Number(idxRateNg >= 0 ? (row[idxRateNg] || 0) : 0);    // ₦/Yard

    // Yard per piece (fallback: entire qty if pcs is 0)
    const onePcsQty = pcs ? (qtyYard / pcs) : qtyYard;

    // 🔹 Total = Qty (Yard) * Rate per Yard
    const lineTotal = qtyYard * rateNg;

    invRows.push([
      new Date(),          // Date
      'BSM-1',             // Supplier
      '',                  // Ref No
      baleNo,              // BaleNo
      fabric,
      quality,
      design,
      color,
      onePcsQty,           // One pcs Qty (Yard/pcs)
      pcs,                 // Qty (pcs)
      uom,
      rateNg,              // Rate (₦/Yard)
      lineTotal,           // Total (₦)
      'Nigeria Godown',    // Location
      qtyYard              // Available Qty (Yard)
    ]);
  });

  // Append to Inventory
  if (invRows.length) {
    inv
      .getRange(inv.getLastRow() + 1, 1, invRows.length, invRows[0].length)
      .setValues(invRows);
  }

  // Mark inward rows as Verified
  rowsToVerify.forEach(({ sheetRow }) => {
    inward.getRange(sheetRow, idxStatus + 1).setValue('Verified');
  });

  return {
    ok: true,
    baleNo,
    verifiedRows: rowsToVerify.length,
    inventoryRowsAdded: invRows.length
  };
}

// ================== 4. INVENTORY SUMMARY & DETAILS ==================
function getInventorySummary() {
  const state = buildInventoryState_();
  const rows  = state.rows || [];

  const fabricMap = {};
  const designSet = new Set();

  rows.forEach(r => {
    if (!r.fabricName) return;
    if (!r.availableQty || r.availableQty <= 0) return;

    if (!fabricMap[r.fabricName]) {
      fabricMap[r.fabricName] = {
        fabricName: r.fabricName,
        baleCount: 0,
        totalQty: 0
      };
    }

    const fm = fabricMap[r.fabricName];
    fm.baleCount = (fm.baleCount || 0) + 1;
    fm.totalQty  = (fm.totalQty  || 0) + (Number(r.availableQty) || 0);

    if (r.designNo) {
      designSet.add(r.fabricName + '|' + r.designNo);
    }
  });

  const fabrics = Object.values(fabricMap).sort((a, b) =>
    a.fabricName.localeCompare(b.fabricName)
  );

  const totalBales = fabrics.reduce((s, f) => s + (f.baleCount || 0), 0);
  const totalQty   = fabrics.reduce((s, f) => s + (f.totalQty  || 0), 0);

  return {
    totalBales,
    totalQty,
    designCount: designSet.size,
    fabrics
  };
}

function getInventoryDetailsByFabric(payload) {
  const fabricName = String(payload.fabricName || '').trim();
  const state = buildInventoryState_();

  return state.rows
    .filter(r => r.fabricName === fabricName && r.availableQty > 0)
    .map(r => ({
      baleNo:       r.baleNo,
      designNo:     r.designNo,
      color:        r.color,
      inwardQty:    r.qtyInwardYard,            // Yard
      soldQty:      r.soldQtyYard,              // Yard
      availableQty: r.availableQty,             // Yard
      uom:          r.uom,
      rate:         r.inwardRate,               // ₦/Yard
      total:        r.qtyInwardYard * r.inwardRate
    }));
}

// ================== 5. STOCK BY BALE (for Sales) ==================
function getStockByBale(payload) {
  const baleNo = String(payload.baleNo || '').trim();
  if (!baleNo) return [];

  const state = buildInventoryState_();
  return state.rows
    .filter(r => r.baleNo === baleNo && r.availableQty > 0)
    .map(r => ({
      baleNos:       r.baleNo,
      fabricName:    r.fabricName,
      fabricQuality: r.fabricQuality,
      designNo:      r.designNo,
      color:         r.color,
      onePcsQty:     r.onePcsQty,    // Yard/pcs
      uom:           r.uom,
      inwardRate:    r.inwardRate,   // ₦/Yard
      availableQty:  r.availableQty, // Yard
      rowIndex:      r.rowIndex
    }));
}

// ================== 6. PDF GENERATION ==================
function generateInvoicePdf_(header, items) {
  // ❗ Only skip if folder ID is actually missing
  if (!INVOICE_FOLDER_ID) {
    return '';
  }

  items = items || [];

  const baleSet = new Set();
  items.forEach(it => {
    if (it.baleNos) {
      String(it.baleNos)
        .split(',')
        .map(b => b.trim())
        .filter(b => b)
        .forEach(b => baleSet.add(b));
    }
  });
  const baleSummary = Array.from(baleSet).join(', ');

  let grandTotal = 0;
  items.forEach(item => {
    const qty  = Number(item.qty) || 0;
    const rate = Number(item.sellingRate) || 0;
    grandTotal += qty * rate;
  });

  const CURRENCY = '₦';

  let html = ''
    + '<html><head><meta charset="utf-8">'
    + '<style>'
    + 'body{font-family:Arial,Helvetica,sans-serif;font-size:11px;margin:24px;color:#111;}'
    + '.header{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:16px;border-bottom:2px solid #222;padding-bottom:8px;}'
    + '.company-block{max-width:55%;}'
    + '.company-name{font-size:20px;font-weight:bold;margin-bottom:4px;letter-spacing:0.5px;text-transform:uppercase;}'
    + '.subtitle{font-size:13px;font-weight:600;margin-bottom:4px;}'
    + '.meta-block{font-size:11px;text-align:right;}'
    + '.meta-block div{margin-bottom:2px;}'
    + '.section-title{font-size:12px;font-weight:bold;margin:14px 0 6px 0;text-transform:uppercase;}'
    + '.box{border:1px solid #ddd;border-radius:4px;padding:8px;margin-bottom:10px;}'
    + '.box p{margin:0 0 3px 0;}'
    + 'table.items{border-collapse:collapse;width:100%;margin-top:8px;}'
    + 'table.items th,table.items td{border:1px solid #ddd;padding:5px 6px;}'
    + 'table.items th{background:#f3f4f6;font-weight:600;font-size:11px;text-align:center;}'
    + 'table.items td.num{text-align:right;}'
    + 'table.items td.center{text-align:center;}'
    + 'table.items tr:nth-child(even){background:#fafafa;}'
    + 'tfoot td{font-weight:bold;background:#f9fafb;}'
    + '.total-label{text-align:right;}'
    + '.grand-total{font-size:12px;font-weight:bold;}'
    + '.footer{margin-top:20px;font-size:10px;color:#444;border-top:1px solid #ddd;padding-top:8px;}'
    + '.terms-list{margin:4px 0 0 14px;padding:0;}'
    + '.terms-list li{margin:2px 0;}'
    + '</style></head><body>';

  html += '<div class="header">';
  html +=   '<div class="company-block">'
         +   '<div class="company-name">Aryan Export</div>'
         +   '<div class="subtitle">Invoice</div>'
         +   '<p style="margin:0;">(Nigeria Godown Sales)</p>'
         + '</div>';

  html +=   '<div class="meta-block">'
         +     '<div><b>Invoice No:</b> ' + (header.invoiceNo || '') + '</div>'
         +     '<div><b>Date:</b> ' + (header.date || '') + '</div>'
         +     '<div><b>Customer:</b> ' + (header.customerName || '') + '</div>';
  if (baleSummary) {
    html +=   '<div><b>Bales:</b> ' + baleSummary + '</div>';
  }
  html +=   '</div>';
  html += '</div>';

  html += '<div class="section-title">Bill To</div>';
  html += '<div class="box">';
  html +=   '<p><b>Name:</b> ' + (header.customerName || '') + '</p>';
  if (header.companyName) {
    html += '<p><b>Company:</b> ' + header.companyName + '</p>';
  }
  if (header.email) {
    html += '<p><b>Email:</b> ' + header.email + '</p>';
  }
  if (header.website) {
    html += '<p><b>Website:</b> ' + header.website + '</p>';
  }
  if (header.vatNo) {
    html += '<p><b>VAT No:</b> ' + header.vatNo + '</p>';
  }
  html += '</div>';

  if (header.customerImageUrl) {
    html += '<img src="' + header.customerImageUrl
         +  '" style="margin-top:6px;max-width:100px;max-height:100px;border-radius:4px;border:1px solid #eee;" alt="Customer Image" />';
  }

  html += '<div class="section-title">Invoice Details</div>';
  html += '<table class="items"><thead><tr>'
       +    '<th>Bale</th>'
       +    '<th>Design</th>'
       +    '<th>Colour Variant</th>'
       +    '<th>Color</th>'
       +    '<th>Qty</th>'
       +    '<th>UOM</th>'
       +    '<th>Rate (' + CURRENCY + ')</th>'
       +    '<th>Amount (' + CURRENCY + ')</th>'
       +  '</tr></thead><tbody>';

  items.forEach(item => {
    const qty  = Number(item.qty) || 0;
    const rate = Number(item.sellingRate) || 0;
    const amt  = qty * rate;

    html += '<tr>'
         +    '<td class="center">' + (item.baleNos || '') + '</td>'
         +    '<td class="center">' + (item.designNo || '') + '</td>'
         +    '<td class="center">' + (item.colourVariant || '') + '</td>'
         +    '<td class="center">' + (item.color || '') + '</td>'
         +    '<td class="num">' + qty + '</td>'
         +    '<td class="center">' + (item.uom || '') + '</td>'
         +    '<td class="num">' + CURRENCY + ' ' + rate.toFixed(2) + '</td>'
         +    '<td class="num">' + CURRENCY + ' ' + amt.toFixed(2) + '</td>'
         +  '</tr>';
  });

  html += '</tbody><tfoot><tr>'
       +    '<td colspan="7" class="total-label grand-total">Total</td>'
       +    '<td class="num grand-total">' + CURRENCY + ' ' + grandTotal.toFixed(2) + '</td>'
       +  '</tr></tfoot></table>';

  html += '<div class="section-title">Terms & Conditions</div>';
  html += '<div class="box footer">'
       +   '<ul class="terms-list">'
       +     '<li>All payments to be made in Nigerian Naira (₦).</li>'
       +     '<li>Goods once sold are not returnable.</li>'
       +     '<li>Any discrepancies must be reported within 24 hours of delivery.</li>'
       +   '</ul>'
       + '</div>';

  html += '</body></html>';

  const blob = Utilities.newBlob(html, 'text/html', 'invoice.html')
    .getAs('application/pdf');

  const folder = DriveApp.getFolderById(INVOICE_FOLDER_ID);
  const file   = folder.createFile(blob)
                      .setName((header.invoiceNo || 'Invoice') + '.pdf');

  return file.getUrl();
}

// ================== 7. SAVE / UPDATE SALES ==================
function saveOutwardEntry(payloadHeader, items) {
  if (!payloadHeader || !items || !items.length) {
    throw new Error('No items to save');
  }

  const sh  = getSalesSheet_();
  const now = new Date();

  let invoiceNo = String(payloadHeader.invoiceNo || '').trim();
  if (!invoiceNo) invoiceNo = nextInvoiceNo_();

  const customerName = String(payloadHeader.customerName || '').trim();
  if (!customerName) {
    throw new Error('Customer Name is required.');
  }

  const header = {
    date:             payloadHeader.date ? new Date(payloadHeader.date) : now,
    customerName:     customerName,
    companyName:      payloadHeader.companyName || '',
    email:            payloadHeader.email || '',
    website:          payloadHeader.website || '',
    vatNo:            payloadHeader.vatNo || '',
    customerImageUrl: payloadHeader.customerImageUrl || '',
    invoiceNo
  };


  const outwardId    = 'OUT-' + now.getTime();
  const rowsToAppend = [];

  items.forEach(item => {
    const inwardRate  = Number(item.inwardRate)  || 0; // ₦/Yard
    const sellingRate = Number(item.sellingRate) || 0;

    if (sellingRate < inwardRate) {
      throw new Error(
        'Selling rate for Design ' + item.designNo +
        ' (' + sellingRate + ') is below inward rate (' + inwardRate + ').'
      );
    }

    const qty        = Number(item.qty) || 0;   // Yard
    const lineAmount = qty * sellingRate;

    rowsToAppend.push([
      outwardId,                               // 1  Outward ID
      header.date,                             // 2  Date
      header.customerName,                     // 3  Customer Name
      header.companyName,                      // 4  Company Name
      header.email,                            // 5  Email
      header.website,                          // 6  Website
      header.vatNo,                            // 7  VAT No
      header.customerImageUrl,                 // 8  Customer Image
      invoiceNo,                               // 9  Invoice No
      item.baleNos || '',                      // 10 Bale Nos
      item.fabricName || '',                   // 11 Fabric Name
      item.fabricQuality || '',                // 12 Fabric Quality
      item.designNo || '',                     // 13 Design No
      item.color || '',                        // 14 Color
      item.colourVariant || '',                // 15 Colour Varient
      item.onePcsQty || 0,                     // 16 One pcs Qty (Yard/pcs)
      qty,                                     // 17 Qty (Yard sold)
      item.uom || 'Yard',                      // 18 UOM
      inwardRate,                              // 19 Inward Rate (₦/Yard)
      sellingRate,                             // 20 Selling Rate
      lineAmount,                              // 21 Line Amount
      '',                                      // 22 Invoice PDF URL (fill later)
      Session.getActiveUser().getEmail() || '',// 23 Created By
      now                                      // 24 Created At
    ]);
  });

  // Generate PDF
  const pdfUrl = generateInvoicePdf_(header, items);

  // Fill PDF URL in col 22
  if (rowsToAppend.length) {
    const startRow = sh.getLastRow() + 1;
    rowsToAppend.forEach(r => {
      r[21] = pdfUrl || '';   // index 21 -> Invoice PDF URL
    });
    sh.getRange(startRow, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }

  // Email copy
  sendInvoiceMail_(header, pdfUrl);

  // Log
  const totalAmount = rowsToAppend.reduce((sum, r) => sum + Number(r[20 + 1] || 0), 0); // col 21
  logEvent_(
    'SAVE_INVOICE',
    'Saved invoice ' + invoiceNo,
    {
      outwardId,
      invoiceNo,
      customer: header.customerName,
      lineCount: items.length,
      totalAmount
    }
  );

  return { success: true, outwardId, invoiceNo, pdfUrl };
}

// -------- Recent invoices (for app UI) --------
function getRecentInvoices() {
  const sh   = getSalesSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];

  const lastCol   = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values    = sh.getRange(2, 1, last - 1, lastCol).getValues();

  const idxDate     = findColIndex_(headerRow, ['Date']);
  const idxInvoice  = findColIndex_(headerRow, ['Invoice No']);
  const idxCustomer = findColIndex_(headerRow, ['Customer Name', 'Customer']);
  const idxLineAmt  = findColIndex_(headerRow, ['Line Amount', 'Amount', 'Total']);
  const idxPdf      = findColIndex_(headerRow, ['Invoice PDF URL', 'PDF URL']);

  const byInvoice = {};

  values.forEach(r => {
    const invNo = idxInvoice >= 0 ? r[idxInvoice] : '';
    if (!invNo) return;

    if (!byInvoice[invNo]) {
      byInvoice[invNo] = {
        invoiceNo: String(invNo),
        date: idxDate >= 0 ? r[idxDate] : '',
        customerName: idxCustomer >= 0 ? (r[idxCustomer] || '') : '',
        total: 0,
        pdfUrl: idxPdf >= 0 ? (r[idxPdf] || '') : ''
      };
    }

    const lineAmount = idxLineAmt >= 0 ? Number(r[idxLineAmt] || 0) : 0;
    byInvoice[invNo].total += lineAmount;
  });

  const tz = Session.getScriptTimeZone();
  const list = Object.values(byInvoice)
    .sort((a, b) => {
      const da = a.date instanceof Date ? a.date.getTime() : 0;
      const db = b.date instanceof Date ? b.date.getTime() : 0;
      return db - da;
    })
    .slice(0, 10)
    .map(rec => ({
      invoiceNo: rec.invoiceNo,
      date: rec.date instanceof Date
        ? Utilities.formatDate(rec.date, tz, 'yyyy-MM-dd')
        : rec.date,
      customerName: rec.customerName,
      total: rec.total,
      pdfUrl: rec.pdfUrl
    }));

  return list;
}

// -------- Invoice details for editing --------
function getInvoiceDetails(payload) {
  const invoiceNo = String(payload && payload.invoiceNo || '').trim();
  if (!invoiceNo) throw new Error('Missing invoiceNo');

  const sh   = getSalesSheet_();
  const last = sh.getLastRow();
  if (last < 2) throw new Error('No invoices yet');

  const lastCol   = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const values    = sh.getRange(2, 1, last - 1, lastCol).getValues();

  const idxDate     = findColIndex_(headerRow, ['Date']);
  const idxCustomer = findColIndex_(headerRow, ['Customer Name']);
  const idxInvoice  = findColIndex_(headerRow, ['Invoice No']);
  const idxBale     = findColIndex_(headerRow, ['Bale Nos', 'Bale No']);
  const idxFabric   = findColIndex_(headerRow, ['Fabric Name']);
  const idxQuality  = findColIndex_(headerRow, ['Fabric Quality']);
  const idxDesign   = findColIndex_(headerRow, ['Design No']);
  const idxColor    = findColIndex_(headerRow, ['Color']);
  const idxColVar   = findColIndex_(headerRow, ['Colour Varient']);
  const idxOnePcs   = findColIndex_(headerRow, ['One pcs Qty']);
  const idxQty      = findColIndex_(headerRow, ['Qty']);
  const idxUom      = findColIndex_(headerRow, ['UOM']);
  const idxInward   = findColIndex_(headerRow, ['Inward Rate']);
  const idxSelling  = findColIndex_(headerRow, ['Selling Rate']);

  const header = {};
  const items  = [];

  values.forEach(r => {
    if (String(r[idxInvoice] || '').trim() !== invoiceNo) return;

    if (!header.invoiceNo) {
      header.invoiceNo    = invoiceNo;
      header.date         = idxDate     >= 0 ? r[idxDate]     : '';
      header.customerName = idxCustomer >= 0 ? r[idxCustomer] : '';
      header.baleNos      = idxBale     >= 0 ? r[idxBale]     : '';

      header.companyName      = '';
      header.email            = '';
      header.website          = '';
      header.vatNo            = '';
      header.customerImageUrl = '';
    }

    items.push({
      baleNos:        idxBale    >= 0 ? r[idxBale]    : '',
      fabricName:     idxFabric  >= 0 ? r[idxFabric]  : '',
      fabricQuality:  idxQuality >= 0 ? r[idxQuality] : '',
      designNo:       idxDesign  >= 0 ? r[idxDesign]  : '',
      color:          idxColor   >= 0 ? r[idxColor]   : '',
      colourVariant:  idxColVar  >= 0 ? r[idxColVar]  : '',
      onePcsQty:      idxOnePcs  >= 0 ? r[idxOnePcs]  : '',
      qty:            idxQty     >= 0 ? r[idxQty]     : 0,
      uom:            idxUom     >= 0 ? r[idxUom]     : '',
      inwardRate:     idxInward  >= 0 ? r[idxInward]  : 0,
      sellingRate:    idxSelling >= 0 ? r[idxSelling] : 0
    });
  });

  if (!header.invoiceNo) throw new Error('Invoice not found');

  return { header, items };
}

// -------- Update invoice: delete & re-save --------
function updateInvoice(payload) {
  const invoiceNo = String(payload && payload.header && payload.header.invoiceNo || '').trim();
  if (!invoiceNo) throw new Error('Missing invoiceNo for update');

  const sh   = getSalesSheet_();
  const last = sh.getLastRow();
  if (last < 2) throw new Error('No invoices yet');

  const lastCol   = sh.getLastColumn();
  const headerRow = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxInvoice = findColIndex_(headerRow, ['Invoice No']);
  if (idxInvoice < 0) throw new Error('Sales sheet missing Invoice No column');

  const values = sh.getRange(2, 1, last - 1, lastCol).getValues();
  const rowsToDelete = [];

  values.forEach((r, i) => {
    if (String(r[idxInvoice] || '').trim() === invoiceNo) {
      rowsToDelete.push(i + 2);
    }
  });

  rowsToDelete.sort((a, b) => b - a).forEach(row => sh.deleteRow(row));

  return saveOutwardEntry(payload.header, payload.items);
}

// ================== 8. DASHBOARD DATA ==================
function getDashboardData_(payload) {
  const sh   = getSalesSheet_();
  const last = sh.getLastRow();
  if (last < 2) {
    return {
      totals: {
        totalInvoices: 0,
        totalQty: 0,
        totalAmount: 0,
        todayAmount: 0
      },
      topCustomers: [],
      recentInvoices: [],
      salesByDate: []
    };
  }

  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const values    = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();

  const idxDate       = findColIndex_(headerRow, ['Date']);
  const idxCustomer   = findColIndex_(headerRow, ['Customer Name', 'Customer']);
  const idxInvoice    = findColIndex_(headerRow, ['Invoice No', 'Invoice']);
  const idxQty        = findColIndex_(headerRow, ['Qty (Sold)', 'Qty', 'Quantity']);
  const idxRate       = findColIndex_(headerRow, ['Selling Rate', 'Rate']);
  const idxLineAmount = findColIndex_(headerRow, ['Line Amount', 'Amount', 'Total']);

  const invoiceMap = {};

  values.forEach(r => {
    const invoiceNo = idxInvoice >= 0 ? r[idxInvoice] : '';
    if (!invoiceNo) return;

    const d        = idxDate     >= 0 ? r[idxDate]     : '';
    const customer = idxCustomer >= 0 ? r[idxCustomer] : '';
    const qty      = idxQty      >= 0 ? Number(r[idxQty])  || 0 : 0;
    const rate     = idxRate     >= 0 ? Number(r[idxRate]) || 0 : 0;

    let amount = 0;
    if (idxLineAmount >= 0 && r[idxLineAmount] !== '') {
      amount = Number(r[idxLineAmount]) || 0;
    } else {
      amount = qty * rate;
    }

    if (!invoiceMap[invoiceNo]) {
      invoiceMap[invoiceNo] = {
        invoiceNo,
        date: d,
        customer,
        qty: 0,
        amount: 0
      };
    }
    invoiceMap[invoiceNo].qty    += qty;
    invoiceMap[invoiceNo].amount += amount;
  });

  const invoices = Object.values(invoiceMap);
  const tz       = Session.getScriptTimeZone();

  let dateFrom = null;
  let dateTo   = null;

  if (payload && payload.dateFrom) {
    dateFrom = new Date(payload.dateFrom);
  }
  if (payload && payload.dateTo) {
    dateTo = new Date(payload.dateTo);
    dateTo.setHours(23, 59, 59, 999);
  }

  function inRange(d) {
    if (!(d instanceof Date)) return true;
    if (dateFrom && d < dateFrom) return false;
    if (dateTo && d > dateTo)   return false;
    return true;
  }

  let totalQty      = 0;
  let totalAmount   = 0;
  let todayAmount   = 0;
  let totalInvoices = 0;
  const customerTotals = {};
  const salesByDateMap = {};

  const today  = new Date();
  const todayY = today.getFullYear();
  const todayM = today.getMonth();
  const todayD = today.getDate();

  invoices.forEach(inv => {
    const d = inv.date;
    if (!inRange(d)) return;

    totalInvoices++;
    totalQty    += inv.qty;
    totalAmount += inv.amount;

    if (d instanceof Date) {
      if (
        d.getFullYear() === todayY &&
        d.getMonth()    === todayM &&
        d.getDate()     === todayD
      ) {
        todayAmount += inv.amount;
      }

      const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      salesByDateMap[key] = (salesByDateMap[key] || 0) + inv.amount;
    }

    if (inv.customer) {
      customerTotals[inv.customer] =
        (customerTotals[inv.customer] || 0) + inv.amount;
    }
  });

  const topCustomers = Object.keys(customerTotals)
    .map(name => ({ customer: name, amount: customerTotals[name] }))
    .sort((a, b) => b.amount - a.amount)
    .slice(0, 5);

  const recentInvoices = invoices
    .slice()
    .filter(inv => inRange(inv.date))
    .sort((a, b) => {
      const ad = a.date instanceof Date ? a.date.getTime() : 0;
      const bd = b.date instanceof Date ? b.date.getTime() : 0;
      return bd - ad;
    })
    .slice(0, 10)
    .map(r => ({
      invoiceNo: r.invoiceNo,
      date: r.date instanceof Date
        ? Utilities.formatDate(r.date, tz, 'yyyy-MM-dd')
        : r.date,
      customer: r.customer,
      qty: r.qty,
      amount: r.amount
    }));

  let salesByDate = Object.keys(salesByDateMap)
    .map(dateStr => ({ date: dateStr, amount: salesByDateMap[dateStr] }))
    .sort((a, b) => a.date.localeCompare(b.date));

  if (salesByDate.length > 7) {
    salesByDate = salesByDate.slice(salesByDate.length - 7);
  }

  return {
    totals: {
      totalInvoices,
      totalQty,
      totalAmount,
      todayAmount
    },
    topCustomers,
    recentInvoices,
    salesByDate
  };
}

// ================== 9. EMAIL & TEST HELPERS ==================
function sendInvoiceMail_(header, pdfUrl) {
  try {
    const to = getMasterEmail_();
    if (!to || !pdfUrl) return;

    const m = String(pdfUrl).match(/[-\w]{25,}/);
    if (!m) return;
    const fileId = m[0];

    const file    = DriveApp.getFileById(fileId);
    const subject = 'Nigeria Invoice ' + (header.invoiceNo || '');
    const body =
      'Dear Team,\n\n' +
      'Please find attached Nigeria Sales invoice ' + (header.invoiceNo || '') +
      ' for customer ' + (header.customerName || '') + '.\n\n' +
      'Regards,\nAryan Export';

    MailApp.sendEmail({
      to,
      subject,
      body,
      attachments: [file.getAs(MimeType.PDF)]
    });

    logEvent_(
      'EMAIL_INVOICE',
      'Sent invoice ' + (header.invoiceNo || '') + ' to ' + to,
      { invoiceNo: header.invoiceNo, email: to }
    );
  } catch (err) {
    logEvent_(
      'EMAIL_INVOICE_ERROR',
      'Failed email for ' + (header && header.invoiceNo),
      { error: String(err), pdfUrl }
    );
  }
}

function testSendEmailPermission() {
  const me = Session.getActiveUser().getEmail() || 'your-email@example.com';
  MailApp.sendEmail({
    to: me,
    subject: 'Test – Nigeria Sales app',
    htmlBody: 'This is a test email to enable MailApp permission.'
  });
}

function testDrivePermission() {
  const folder = DriveApp.getFolderById(INVOICE_FOLDER_ID);
  Logger.log('Folder name: ' + folder.getName());
}

// ================== 10. WEB API via doGet ==================
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action)
      ? String(e.parameter.action).trim()
      : '';

    const payloadStr = (e && e.parameter && e.parameter.payload)
      ? e.parameter.payload
      : '{}';

    let payload = {};
    try {
      payload = JSON.parse(payloadStr);
    } catch (err) {
      payload = {};
    }

    let result;

    switch (action) {
      case 'getPendingInward':
        result = getPendingInward();
        break;

      case 'verifyBaleRow':
        result = verifyBaleRow(payload);
        break;

      case 'getInventorySummary':
        result = getInventorySummary();
        break;

      case 'getInventoryDetailsByFabric':
        result = getInventoryDetailsByFabric(payload);
        break;

      case 'getStockByBale':
        result = getStockByBale(payload);
        break;

      case 'saveOutwardEntry':
        result = saveOutwardEntry(payload.header, payload.items);
        break;

      case 'getRecentInvoices':
        result = getRecentInvoices();
        break;

      case 'getInvoiceDetails':
        result = getInvoiceDetails(payload);
        break;

      case 'updateInvoice':
        result = updateInvoice(payload);
        break;

      case 'dashboard':
        result = getDashboardData_(payload);
        break;

      case 'debugAction':
        result = {
          actionReceived: action,
          params: e && e.parameter ? e.parameter : {}
        };
        break;

      default:
        result = { message: 'OK from doGet (no action)' };
        break;
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
