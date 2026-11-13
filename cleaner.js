/* 進銷項資料清洗 — 瀏覽器端：讀檔、欄位對應、畫面與 Excel 輸出
 *
 * 清洗邏輯本身在 embed/invoice-cleaner-core.js（window.InvoiceCleaner），
 * 網站與嵌入版共用同一份；本檔只負責把 xlsx 轉成核心庫的輸入格式、再把結果畫出來。
 */

const Core = window.InvoiceCleaner;

// ============================================================
// 工具函式
// ============================================================

function $(id) { return document.getElementById(id); }

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), {
          type: 'array',
          cellDates: true,
          cellNF: false,
        });
        resolve(wb);
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function sheetToRows(ws) {
  // 回傳二維陣列（含空白）
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

function toNumber(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    const cleaned = v.replace(/,/g, '').trim();
    if (cleaned === '') return 0;
    const n = Number(cleaned);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}

function toCleanString(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

function fmtNum(n) {
  if (!n && n !== 0) return '';
  return Math.round(n).toLocaleString('en-US');
}

function findColumns(headerRow, map) {
  const result = {};
  for (const key in map) {
    const aliases = map[key];
    let idx = -1;
    for (const alias of aliases) {
      idx = headerRow.findIndex(h => h !== null && String(h).trim().replace(/\s+/g, '') === alias.replace(/\s+/g, ''));
      if (idx >= 0) break;
    }
    result[key] = idx;
  }
  return result;
}

// 在前 maxScan 列中找含有 marker 儲存格的標頭列
function findHeaderRow(rows, marker, maxScan = 10) {
  for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
    if (rows[i] && rows[i].some(c => c === marker)) return i;
  }
  return -1;
}

function todayStamp() {
  const d = new Date();
  return `${d.getFullYear() - 1911}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

// ============================================================
// 公司名稱／統編 抓取
// ============================================================

// 從文中系統檔（印刷格式）找「用戶：」列
function findCompanyFromWenzhong(rows) {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i];
    if (!row) continue;
    const idx = row.findIndex(c => c !== null && c !== undefined && /^用戶[:：]/.test(String(c).trim()));
    if (idx >= 0) {
      // 同一儲存格 '用戶：陳桓有限公司'?
      const sameCell = String(row[idx]).replace(/^用戶[:：]\s*/, '').trim();
      if (sameCell) return sameCell;
      // 公司名在 idx 之後第一個非空、非「頁/製表」的儲存格
      for (let j = idx + 1; j < row.length; j++) {
        const v = toCleanString(row[j]);
        if (!v) continue;
        if (/^(頁|製表|第|序號|金額|稅額|日期|發票|送件)/.test(v)) continue;
        return v;
      }
    }
  }
  return '';
}

// 從平台檔的 Invoice 表抓買方/賣方第一筆有效的統編+名稱
function findCompanyFromPlatform(wb, side /* 'buyer' | 'seller' */) {
  const ws = wb.Sheets['Invoice'] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) return null;
  const rows = sheetToRows(ws);
  if (!rows.length) return null;
  const header = rows[0];
  const idTaxId = header.findIndex(c => c === (side === 'buyer' ? '買方統一編號' : '賣方統一編號'));
  const idName = header.findIndex(c => c === (side === 'buyer' ? '買方名稱' : '賣方名稱'));
  if (idTaxId < 0 || idName < 0) return null;
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const taxId = toCleanString(row[idTaxId]).replace(/\.0+$/, '').replace(/\D/g, '');
    const name = toCleanString(row[idName]);
    if (taxId && name && !/^0+$/.test(taxId)) {
      return { taxId, name };
    }
  }
  return null;
}

// 從工作簿/頁首抓統編（8 碼純數字）
function extractTaxIdFromFile(wb, rows) {
  // 1) 工作表名稱（如 "60689392_進項憑證明細表..."）
  if (wb && wb.SheetNames) {
    for (const sn of wb.SheetNames) {
      const m = String(sn).match(/(?:^|\D)(\d{8})(?:$|\D)/);
      if (m) return m[1];
    }
  }
  // 2) 工作簿元資料 Title
  if (wb && wb.Props && wb.Props.Title) {
    const m = String(wb.Props.Title).match(/(\d{8})/);
    if (m) return m[1];
  }
  // 3) 前 5 列任意儲存格找 8 碼純數字
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const row = rows[i] || [];
    for (const c of row) {
      const s = toCleanString(c);
      const m = s.match(/(?:^|\D)(\d{8})(?:$|\D)/);
      if (m) return m[1];
    }
  }
  return '';
}

// ============================================================
// 折讓明細（銷項／進項共用）
// ============================================================

const ALLOWANCE_HEADER_MAP = {
  allowanceNo: ['折讓單號碼'],
  status: ['折讓單狀態'],
  category: ['折讓單類別'],
  invoiceNo: ['發票號碼'],
  itemName: ['品項名稱', '品名'],
  itemAmount: ['品項折讓金額(不含稅)', '品項折讓金額'],
  amount: ['折讓金額(不含稅)', '折讓金額'],
  tax: ['折讓稅額'],
  allowanceDate: ['折讓單日期'],
};

// 讀平台下載的折讓明細（一列一品項）。expected = '銷項' | '進項'，用來擋上傳錯檔。
function parseAllowanceFile(wb, expected) {
  const ws = wb.Sheets['allowance'] || wb.Sheets[wb.SheetNames[0]];
  const rows = ws ? sheetToRows(ws) : [];
  const h = findHeaderRow(rows, '折讓單號碼');
  if (h < 0) throw new Error(`找不到「折讓單號碼」標頭欄，請確認是否為平台下載的${expected}折讓明細`);
  const cols = findColumns(rows[h], ALLOWANCE_HEADER_MAP);
  const missing = ['amount', 'tax'].filter(k => cols[k] < 0);
  if (missing.length) {
    throw new Error('折讓檔缺少必要欄位：' + missing.map(k => ALLOWANCE_HEADER_MAP[k][0]).join(', '));
  }
  const items = [];
  for (let i = h + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || toCleanString(row[cols.allowanceNo]) === '') continue;
    const item = {};
    for (const key in cols) if (cols[key] >= 0) item[key] = row[cols[key]];
    if (item.category && !String(item.category).includes(expected)) {
      throw new Error(`折讓檔的折讓單類別是「${item.category}」，這裡需要${expected}折讓明細`);
    }
    items.push(item);
  }
  return items;
}

// ============================================================
// 銷項清洗
// ============================================================

const OUTPUT_HEADER_MAP = {
  invoiceNo: ['發票號碼'],
  status: ['發票狀態'],
  invoiceDate: ['發票日期'],
  buyerTaxId: ['買方統一編號', '買方統編'],
  salesAmount: ['銷售額合計'],
  taxAmount: ['營業稅'],
  totalAmount: ['總計'],
  // 以下為選擇性欄位：有就用來拆 應稅／零稅率／免稅
  taxType: ['課稅別'],
  zeroRateAmount: ['零稅銷售額', '零稅率銷售額'],
  exemptAmount: ['免稅銷售額'],
};

function cleanOutput(wb, allowanceWb) {
  // 嘗試找到 Invoice 工作表（或第一張）
  const ws = wb.Sheets['Invoice'] || wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('找不到工作表');

  const rows = sheetToRows(ws);
  if (rows.length === 0) throw new Error('檔案無資料');

  const headerRowIdx = findHeaderRow(rows, '發票號碼');
  if (headerRowIdx < 0) throw new Error('找不到「發票號碼」標頭欄，請確認檔案是否為平台下載的銷項發票明細');

  const cols = findColumns(rows[headerRowIdx], OUTPUT_HEADER_MAP);
  const required = ['invoiceNo', 'totalAmount', 'salesAmount', 'taxAmount'];
  const missing = required.filter(k => cols[k] < 0);
  if (missing.length) throw new Error('缺少必要欄位：' + missing.join(', '));

  const get = (row, key) => (cols[key] >= 0 ? row[cols[key]] : undefined);
  // 零稅／免稅欄存在就一律給數字（空白視為 0），整份檔案一致走「欄位拆分」
  const getNum = (row, key) => (cols[key] >= 0 ? toNumber(row[cols[key]]) : undefined);

  const invoices = [];
  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    invoices.push({
      invoiceNo: row[cols.invoiceNo],
      status: get(row, 'status'),
      buyerTaxId: get(row, 'buyerTaxId'),
      salesAmount: row[cols.salesAmount],
      taxAmount: row[cols.taxAmount],
      totalAmount: row[cols.totalAmount],
      taxType: get(row, 'taxType'),
      zeroRateAmount: getNum(row, 'zeroRateAmount'),
      exemptAmount: getNum(row, 'exemptAmount'),
      invoiceDate: get(row, 'invoiceDate'),
    });
  }

  return Core.cleanOutputData(invoices, {
    // 公司資訊：銷項取賣方
    company: findCompanyFromPlatform(wb, 'seller') || { taxId: '', name: '' },
    allowances: allowanceWb ? parseAllowanceFile(allowanceWb, '銷項') : undefined,
  });
}

const OUTPUT_COLUMNS = [
  { key: 'label', title: '類別' },
  { key: 'count', title: '筆數' },
  { key: 'sales', title: '應稅銷售額' },
  { key: 'tax', title: '稅額' },
  { key: 'zero', title: '零稅率銷售額' },
  { key: 'exempt', title: '免稅銷售額' },
  { key: 'total', title: '合計' },
];

// 結果下方的說明：對應申報書欄位、需人工確認的項目
function outputFootnotes(result) {
  const notes = ['小計列對應申報書（401／403）「收銀機發票(三聯式)及電子發票」列：應稅銷售額→第 5 欄、稅額→第 6 欄、零稅率銷售額→第 7 欄、免稅銷售額→第 8 欄。'];
  const alw = result.allowance;
  if (alw) {
    notes.push('「減：退回及折讓」列依序對應第 17、18、19、20 欄（筆數為折讓單張數）；如另有紙本發票的退回及折讓，須合併填報。');
    if (alw.unverified.length) {
      notes.push(`※ 折讓單 ${alw.unverified.length} 張稅額為 0、查無原發票課稅別（多為前期發票），暫列免稅，請確認：${alw.unverified.join('、')}`);
    }
  }
  if (result.special.length) notes.push(`※ 特種稅率發票 ${result.special.length} 筆未納入本表，請另行處理。`);
  return notes;
}

function renderOutputResult(result) {
  $('output-period').textContent = result.period || '';
  const c = result.company || {};
  $('output-company').innerHTML = c.name
    ? `${c.name}${c.taxId ? `<span class="tax-id">統一編號 ${c.taxId}</span>` : ''}`
    : '';
  const alw = result.allowance;
  const notes = [];
  if (result.skipped.length) notes.push(`已略過非開立確認 ${result.skipped.length} 筆`);
  if (result.special.length) notes.push(`特種稅率 ${result.special.length} 筆未納入本表`);
  if (alw && alw.skipped.length) notes.push(`已略過作廢折讓單 ${alw.skipped.length} 張`);
  $('output-meta').textContent = `製表日期：${todayStamp()}　共 ${result.subtotal.count} 筆${notes.length ? `（${notes.join('；')}）` : ''}`;

  const cell = (r, col) => col.key === 'label'
    ? `<td class="center">${r.label}</td>`
    : `<td class="num">${col.key === 'count' ? r.count : fmtNum(r[col.key])}</td>`;
  const tr = (r, cls) => `<tr${cls ? ` class="${cls}"` : ''}>${OUTPUT_COLUMNS.map(col => cell(r, col)).join('')}</tr>`;

  let html = `<thead><tr>${OUTPUT_COLUMNS.map(col =>
    `<th class="${col.key === 'label' ? 'center' : 'num'}">${col.title}</th>`).join('')}</tr></thead><tbody>`;
  for (const r of result.rows) html += tr(r);
  html += tr({ label: '小計', ...result.subtotal }, 'grand-total');
  if (alw) html += tr(alw.row, 'subtotal');
  html += `</tbody>`;
  $('output-table').innerHTML = html;
  $('output-note').textContent = outputFootnotes(result).join('\n');

  $('output-result').hidden = false;
  $('output-print-btn').disabled = false;
  $('output-download-btn').disabled = false;
}

// ---------- Excel 樣式工具 ----------
const FONT = { name: '微軟正黑體', sz: 11 };
const BORDER_THIN = {
  top: { style: 'thin', color: { rgb: '999999' } },
  bottom: { style: 'thin', color: { rgb: '999999' } },
  left: { style: 'thin', color: { rgb: '999999' } },
  right: { style: 'thin', color: { rgb: '999999' } },
};
const STYLE = {
  sectionTitle: {
    font: { name: '微軟正黑體', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  },
  periodLine: {
    font: { name: '微軟正黑體', sz: 12, bold: true, color: { rgb: '1E3A8A' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  },
  companyLine: {
    font: { name: '微軟正黑體', sz: 12, bold: true, color: { rgb: '1F2937' } },
    alignment: { horizontal: 'left', vertical: 'center' },
  },
  header: {
    font: { name: '微軟正黑體', sz: 11, bold: true, color: { rgb: '1F2937' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'E5E7EB' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: BORDER_THIN,
  },
  cellText: {
    font: FONT,
    alignment: { horizontal: 'left', vertical: 'center', wrapText: false },
    border: BORDER_THIN,
  },
  cellCenter: {
    font: FONT,
    alignment: { horizontal: 'center', vertical: 'center' },
    border: BORDER_THIN,
  },
  cellNum: {
    font: FONT,
    alignment: { horizontal: 'right', vertical: 'center' },
    border: BORDER_THIN,
    numFmt: '#,##0',
  },
  subtotalLabel: {
    font: { name: '微軟正黑體', sz: 11, bold: true, color: { rgb: '1F2937' } },
    fill: { patternType: 'solid', fgColor: { rgb: 'FEF3C7' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: BORDER_THIN,
  },
  subtotalNum: {
    font: { name: '微軟正黑體', sz: 11, bold: true },
    fill: { patternType: 'solid', fgColor: { rgb: 'FEF3C7' } },
    alignment: { horizontal: 'right', vertical: 'center' },
    border: BORDER_THIN,
    numFmt: '#,##0',
  },
  grandTotalLabel: {
    font: { name: '微軟正黑體', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1D4ED8' } },
    alignment: { horizontal: 'center', vertical: 'center' },
    border: BORDER_THIN,
  },
  grandTotalNum: {
    font: { name: '微軟正黑體', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
    fill: { patternType: 'solid', fgColor: { rgb: '1D4ED8' } },
    alignment: { horizontal: 'right', vertical: 'center' },
    border: BORDER_THIN,
    numFmt: '#,##0',
  },
  footnote: {
    font: { name: '微軟正黑體', sz: 10, color: { rgb: '4B5563' } },
  },
};

function makeCell(value, style) {
  if (value === null || value === undefined || value === '') {
    return { v: '', t: 's', s: style };
  }
  if (typeof value === 'number') {
    return { v: value, t: 'n', s: style };
  }
  return { v: String(value), t: 's', s: style };
}

function colLetter(n) {
  // 0-based → 'A','B',...,'Z','AA',...
  let s = '';
  n = n + 1;
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function setCell(ws, r, c, cell) {
  ws[colLetter(c) + (r + 1)] = cell;
}

function ensureRange(ws, maxR, maxC) {
  ws['!ref'] = `A1:${colLetter(maxC)}${maxR + 1}`;
}

function downloadOutputXlsx(result) {
  const ws = {};
  let r = 0;
  // 標題
  setCell(ws, r, 0, makeCell('電子銷項發票清洗結果', STYLE.sectionTitle));
  r++;
  // 公司名稱
  if (result.company && result.company.name) {
    const cn = result.company;
    const txt = `公司名稱：${cn.name}${cn.taxId ? `（統一編號 ${cn.taxId}）` : ''}`;
    setCell(ws, r, 0, makeCell(txt, STYLE.companyLine));
    r++;
  }
  // 申報期別（如有）
  if (result.period) {
    setCell(ws, r, 0, makeCell(`申報期別：${result.period}`, STYLE.periodLine));
    r++;
  }
  r++;
  // 表頭
  const lastCol = OUTPUT_COLUMNS.length - 1;
  OUTPUT_COLUMNS.forEach((c, i) => setCell(ws, r, i, makeCell(c.title, STYLE.header)));
  r++;
  const writeRow = (row, labelStyle, numStyle) => {
    OUTPUT_COLUMNS.forEach((c, i) =>
      setCell(ws, r, i, makeCell(row[c.key], c.key === 'label' ? labelStyle : numStyle)));
    r++;
  };
  for (const row of result.rows) writeRow(row, STYLE.cellCenter, STYLE.cellNum);
  writeRow({ label: '小計', ...result.subtotal }, STYLE.grandTotalLabel, STYLE.grandTotalNum);
  if (result.allowance) writeRow(result.allowance.row, STYLE.subtotalLabel, STYLE.subtotalNum);
  // 備註
  for (const t of outputFootnotes(result)) {
    r++;
    setCell(ws, r, 0, makeCell(t, STYLE.footnote));
  }

  ensureRange(ws, r, lastCol);
  ws['!cols'] = [{ wch: 16 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];
  ws['!rows'] = [{ hpt: 22 }];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '清洗後');
  XLSX.writeFile(wb, '電子銷項清洗_' + Date.now() + '.xlsx');
}

// ============================================================
// 進項清洗
// ============================================================

const WENZHONG_HEADER_MAP = {
  no: ['序號'],
  date: ['日期'],
  invoiceNo: ['發票\\稅單\\報單', '發票/稅單/報單', '發票'],
  amount: ['憑證金額', '金額'],
  taxAmount: ['稅額'],
  deductible: ['扣抵否', '扣抵'],
  // 以下為選擇性欄位（找不到不報錯），用於品名 fallback
  creditMemo: ['貸方摘要'],
  debitMemo: ['借方摘要'],
};

function findWenzhongHeader(rows) {
  for (let i = 0; i < Math.min(rows.length, 20); i++) {
    const row = rows[i];
    if (!row) continue;
    const cells = row.map(c => c === null ? '' : String(c).trim());
    const hasNo = cells.includes('序號');
    const hasDate = cells.includes('日期');
    const hasInvoice = cells.some(c => c.includes('發票') && (c.includes('稅單') || c.includes('報單') || c === '發票'));
    if (hasNo && hasDate && hasInvoice) return i;
  }
  return -1;
}

function buildItemNameMap(platformWb) {
  // 從平台檔的 Invoice_details 工作表，建立 發票號碼 → 第一筆品名 的對應
  const map = new Map();
  const sheetNames = ['Invoice_details', 'Sheet1', ...platformWb.SheetNames];
  let ws = null;
  for (const name of sheetNames) {
    if (platformWb.Sheets[name]) {
      const candidate = platformWb.Sheets[name];
      const test = sheetToRows(candidate);
      if (test.length > 0 && test[0] && test[0].some(c => c === '發票號碼') && test[0].some(c => c === '品名')) {
        ws = candidate;
        break;
      }
    }
  }
  if (!ws) return map;
  const rows = sheetToRows(ws);
  if (rows.length === 0) return map;
  const header = rows[0];
  const idxInv = header.findIndex(c => c === '發票號碼');
  const idxName = header.findIndex(c => c === '品名');
  const idxSeq = header.findIndex(c => c === '序號');
  if (idxInv < 0 || idxName < 0) return map;

  // 收集每張發票的所有明細
  const collected = new Map(); // invoiceNo → [{seq, name}, ...]
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const inv = toCleanString(row[idxInv]);
    const name = toCleanString(row[idxName]);
    if (!inv || !name) continue;
    const seq = idxSeq >= 0 ? toNumber(row[idxSeq]) : i;
    if (!collected.has(inv)) collected.set(inv, []);
    collected.get(inv).push({ seq, name });
  }
  // 依序號排序，取第一筆作為代表品名
  for (const [inv, list] of collected) {
    list.sort((a, b) => a.seq - b.seq);
    map.set(inv, list[0].name);
  }
  return map;
}

// 從平台 Invoice 表建立電子發票號碼集合，用於過濾文中系統檔中的紙本發票
function buildPlatformInvoiceSet(platformWb) {
  const set = new Set();
  const ws = platformWb.Sheets['Invoice'] || platformWb.Sheets[platformWb.SheetNames[0]];
  if (!ws) return set;
  const rows = sheetToRows(ws);
  if (!rows.length) return set;
  const idx = rows[0].findIndex(c => c === '發票號碼');
  if (idx < 0) return set;
  for (let i = 1; i < rows.length; i++) {
    const inv = toCleanString((rows[i] || [])[idx]);
    if (inv) set.add(inv);
  }
  return set;
}

function cleanInput(wenzhongWb, platformWb, allowanceWb) {
  // 主資料：文中系統 進項憑證明細表
  // 1) 優先選 sheet 名含「進項憑證明細表」的 sheet
  // 2) 否則掃所有 sheet 找有正確標頭列的
  let ws = null;
  const preferred = wenzhongWb.SheetNames.filter(n => n.includes('進項憑證明細表'));
  for (const name of preferred) {
    const candidate = wenzhongWb.Sheets[name];
    if (findWenzhongHeader(sheetToRows(candidate)) >= 0) {
      ws = candidate;
      break;
    }
  }
  if (!ws) {
    for (const name of wenzhongWb.SheetNames) {
      if (name === '清洗後' || name === '清洗結果') continue;
      const candidate = wenzhongWb.Sheets[name];
      if (findWenzhongHeader(sheetToRows(candidate)) >= 0) {
        ws = candidate;
        break;
      }
    }
  }
  if (!ws) ws = wenzhongWb.Sheets[wenzhongWb.SheetNames[0]];
  const rows = sheetToRows(ws);
  const headerIdx = findWenzhongHeader(rows);
  if (headerIdx < 0) throw new Error('找不到文中系統「進項憑證明細表」標頭列（須含「序號」「日期」「發票」欄位）');

  // 標頭列可能有空白合併欄，需找實際有名稱的欄位位置
  const headerRow = rows[headerIdx];
  const cols = {};
  for (const key in WENZHONG_HEADER_MAP) {
    const aliases = WENZHONG_HEADER_MAP[key];
    let idx = -1;
    for (const alias of aliases) {
      const target = alias.replace(/\\/g, '');
      idx = headerRow.findIndex(h => {
        if (h === null || h === undefined) return false;
        return String(h).replace(/\\/g, '').replace(/\s+/g, '') === target.replace(/\s+/g, '');
      });
      if (idx >= 0) break;
    }
    cols[key] = idx;
  }
  const required = ['no', 'date', 'invoiceNo', 'amount', 'taxAmount', 'deductible'];
  const missing = required.filter(k => cols[k] < 0);
  if (missing.length) {
    throw new Error('文中系統檔缺少必要欄位：' + missing.map(k => WENZHONG_HEADER_MAP[k][0]).join(', '));
  }

  const memo = (row, key) => (cols[key] >= 0 ? toCleanString(row[cols[key]]) : '');
  const certs = [];
  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    // 跳過區塊註記列（單純文字、無金額）
    const noVal = row[cols.no];
    if (noVal === null || noVal === undefined || noVal === '') continue;
    if (typeof noVal === 'string' && !/^\d/.test(noVal.trim())) continue;
    certs.push({
      invoiceNo: row[cols.invoiceNo],
      date: row[cols.date],
      amount: row[cols.amount],
      taxAmount: row[cols.taxAmount],
      deductible: row[cols.deductible],
      itemName: memo(row, 'creditMemo') || memo(row, 'debitMemo'), // 平台查無品名時的 fallback
    });
  }

  // 公司資訊：先從文中檔的「用戶：」列抓，否則從平台檔的買方名稱抓
  const wzCompanyName = findCompanyFromWenzhong(rows);
  const platformBuyer = findCompanyFromPlatform(platformWb, 'buyer');
  let companyName = wzCompanyName;
  if (!companyName && platformBuyer && platformBuyer.name && platformBuyer.name !== platformBuyer.taxId) {
    companyName = platformBuyer.name;
  }
  const company = {
    name: companyName,
    taxId: (platformBuyer && platformBuyer.taxId) || extractTaxIdFromFile(wenzhongWb, rows),
  };

  return Core.cleanInputData(certs, {
    // 過濾紙本發票：發票號碼不在電子發票平台下載清單中 → 排除
    platformInvoiceNos: buildPlatformInvoiceSet(platformWb),
    platformItemNames: buildItemNameMap(platformWb),
    company,
    allowances: allowanceWb ? parseAllowanceFile(allowanceWb, '進項') : undefined,
  });
}

// 進項結果要列出的區段：[標題, {rows, subtotal}]
function inputSections(result) {
  const list = [['可扣抵進項', result.deductible], ['不可扣抵進項', result.nondeductible]];
  const alw = result.allowance;
  if (alw) {
    list.push(['可扣抵進項之退出及折讓', alw.deductible]);
    if (alw.nondeductible.rows.length) list.push(['不可扣抵進項之退出及折讓', alw.nondeductible]);
  }
  return list;
}

function inputFootnotes(result) {
  const alw = result.allowance;
  if (!alw) return [];
  const notes = [
    '「可扣抵進項之退出及折讓」小計對應申報書（401／403）「減：退出、折讓及海關退還溢繳稅款」第 40／41 欄（屬固定資產者為第 42／43 欄）。',
    '折讓區段的發票號碼為被折讓的原發票號碼，扣抵別依原發票在文中明細中的扣抵欄。',
  ];
  if (alw.unmatched.length) {
    notes.push(`* 原發票不在本期文中明細（前期或紙本發票），扣抵別預設為 Y，請確認（${alw.unmatched.length} 張）。`);
  }
  return notes;
}

// 扣抵欄顯示：查不到原發票的折讓加註 *
function deductibleLabel(r) {
  return (r.deductible || '') + (r.matched === false ? '*' : '');
}

function renderInputSection(title, section) {
  const rows = section.rows;
  const sub = section.subtotal;
  let html = `<h3 class="result-section-title">${title}（${rows.length} 筆）</h3>`;
  html += `<table class="result-table">
    <thead><tr>
      <th class="num" style="width:50px">序號</th>
      <th class="center" style="width:90px">日期</th>
      <th class="center" style="width:120px">發票號碼</th>
      <th>品名</th>
      <th class="num" style="width:90px">金額</th>
      <th class="num" style="width:80px">稅額</th>
      <th class="center" style="width:60px">扣抵</th>
    </tr></thead>
    <tbody>`;
  for (const r of rows) {
    html += `<tr>
      <td class="num">${r.seq}</td>
      <td class="center">${r.date}</td>
      <td class="center">${r.invoiceNo}</td>
      <td>${r.itemName || '<span style="color:#9ca3af">—</span>'}</td>
      <td class="num">${fmtNum(r.amount)}</td>
      <td class="num">${fmtNum(r.tax)}</td>
      <td class="center">${deductibleLabel(r)}</td>
    </tr>`;
  }
  html += `<tr class="subtotal">
    <td colspan="4" class="center">小計</td>
    <td class="num">${fmtNum(sub.amount)}</td>
    <td class="num">${fmtNum(sub.tax)}</td>
    <td></td>
  </tr></tbody></table>`;
  return html;
}

function renderInputResult(result) {
  const totalCount = result.deductible.rows.length + result.nondeductible.rows.length;
  $('input-period').textContent = result.period || '';
  const c = result.company || {};
  $('input-company').innerHTML = c.name
    ? `${c.name}${c.taxId ? `<span class="tax-id">統一編號 ${c.taxId}</span>` : ''}`
    : '';
  const notes = [];
  if (result.paperCount) notes.push(`已排除非電子發票 ${result.paperCount} 筆`);
  if (result.missingNames) notes.push(`其中 ${result.missingNames} 筆查無平台明細品名`);
  const alw = result.allowance;
  if (alw) {
    const n = alw.deductible.rows.length + alw.nondeductible.rows.length;
    notes.push(`進項折讓 ${n} 張${alw.skipped.length ? `、已略過作廢 ${alw.skipped.length} 張` : ''}`);
  }
  $('input-meta').textContent = `製表日期：${todayStamp()}　共 ${totalCount} 筆${notes.length ? `（${notes.join('；')}）` : ''}`;

  const footnotes = inputFootnotes(result);
  $('input-tables').innerHTML =
    inputSections(result).map(([title, section]) => renderInputSection(title, section)).join('') +
    (footnotes.length ? `<p class="result-note">${footnotes.join('\n')}</p>` : '');

  $('input-result').hidden = false;
  $('input-print-btn').disabled = false;
  $('input-download-btn').disabled = false;
}

function downloadInputXlsx(result) {
  const ws = { '!merges': [] };
  const rowHeights = [];
  let r = 0;
  const mergeRow = (c0, c1) => ws['!merges'].push({ s: { r, c: c0 }, e: { r, c: c1 } });

  // 公司名稱（如有）
  if (result.company && result.company.name) {
    const cn = result.company;
    const txt = `公司名稱：${cn.name}${cn.taxId ? `（統一編號 ${cn.taxId}）` : ''}`;
    setCell(ws, r, 0, makeCell(txt, STYLE.companyLine));
    mergeRow(0, 6);
    rowHeights[r] = { hpt: 22 };
    r++;
  }
  // 申報期別（如有）
  if (result.period) {
    setCell(ws, r, 0, makeCell(`申報期別：${result.period}`, STYLE.periodLine));
    mergeRow(0, 6);
    rowHeights[r] = { hpt: 22 };
    r++;
  }
  if (r > 0) r++; // 空一列

  function writeSection(title, rows, subtotal) {
    // 區段標題（合併 7 欄）
    setCell(ws, r, 0, makeCell(title, STYLE.sectionTitle));
    mergeRow(0, 6);
    rowHeights[r] = { hpt: 24 };
    r++;
    // 欄標頭
    const headers = ['序號', '日期', '發票號碼', '品名', '金額', '稅額', '扣抵'];
    headers.forEach((h, c) => setCell(ws, r, c, makeCell(h, STYLE.header)));
    rowHeights[r] = { hpt: 22 };
    r++;
    // 資料列
    for (const row of rows) {
      setCell(ws, r, 0, makeCell(row.seq, STYLE.cellCenter));
      setCell(ws, r, 1, makeCell(row.date, STYLE.cellCenter));
      setCell(ws, r, 2, makeCell(row.invoiceNo, STYLE.cellCenter));
      setCell(ws, r, 3, makeCell(row.itemName, STYLE.cellText));
      setCell(ws, r, 4, makeCell(row.amount, STYLE.cellNum));
      setCell(ws, r, 5, makeCell(row.tax, STYLE.cellNum));
      setCell(ws, r, 6, makeCell(deductibleLabel(row), STYLE.cellCenter));
      r++;
    }
    // 小計列：合併前 4 欄為「小計」標籤
    setCell(ws, r, 0, makeCell('小計', STYLE.subtotalLabel));
    setCell(ws, r, 1, makeCell('', STYLE.subtotalLabel));
    setCell(ws, r, 2, makeCell('', STYLE.subtotalLabel));
    setCell(ws, r, 3, makeCell('', STYLE.subtotalLabel));
    mergeRow(0, 3);
    setCell(ws, r, 4, makeCell(subtotal.amount, STYLE.subtotalNum));
    setCell(ws, r, 5, makeCell(subtotal.tax, STYLE.subtotalNum));
    setCell(ws, r, 6, makeCell('', STYLE.subtotalLabel));
    rowHeights[r] = { hpt: 22 };
    r++;
    // 區段間空一列
    r++;
  }

  for (const [title, section] of inputSections(result)) {
    writeSection(title, section.rows, section.subtotal);
  }
  for (const t of inputFootnotes(result)) {
    setCell(ws, r, 0, makeCell(t, STYLE.footnote));
    r++;
  }

  ensureRange(ws, r - 1, 6);
  ws['!cols'] = [
    { wch: 6 },   // 序號
    { wch: 11 },  // 日期
    { wch: 14 },  // 發票號碼
    { wch: 38 },  // 品名
    { wch: 12 },  // 金額
    { wch: 10 },  // 稅額
    { wch: 7 },   // 扣抵
  ];
  ws['!rows'] = rowHeights;

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '清洗後');
  XLSX.writeFile(wb, '電子進項清洗_' + Date.now() + '.xlsx');
}

// ============================================================
// 事件綁定
// ============================================================

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    $('tab-' + btn.dataset.tab).classList.add('active');
  });
});

function bindFileInput(inputId, nameId, onChange) {
  const input = $(inputId);
  const nameEl = $(nameId);
  input.addEventListener('change', () => {
    const file = input.files[0];
    if (file) {
      nameEl.textContent = file.name;
      input.parentElement.classList.add('has-file');
    } else {
      nameEl.textContent = '尚未選擇檔案';
      input.parentElement.classList.remove('has-file');
    }
    onChange(file);
  });
}

// 銷項
let outputResult = null;
let outputFile = null;
let outputAlwFile = null;
bindFileInput('output-file', 'output-file-name', file => {
  outputFile = file;
  $('output-clean-btn').disabled = !file;
  $('output-error').textContent = '';
});
bindFileInput('output-alw-file', 'output-alw-file-name', file => {
  outputAlwFile = file;
  $('output-error').textContent = '';
});

$('output-clean-btn').addEventListener('click', async () => {
  if (!outputFile) return;
  $('output-error').textContent = '';
  try {
    const [wb, alwWb] = await Promise.all([
      readWorkbook(outputFile),
      outputAlwFile ? readWorkbook(outputAlwFile) : null,
    ]);
    outputResult = cleanOutput(wb, alwWb);
    renderOutputResult(outputResult);
  } catch (err) {
    console.error(err);
    $('output-error').textContent = '清洗失敗：' + err.message;
    $('output-result').hidden = true;
    $('output-print-btn').disabled = true;
    $('output-download-btn').disabled = true;
  }
});

$('output-print-btn').addEventListener('click', () => {
  // 切換到銷項分頁，再列印
  document.querySelector('.tab-btn[data-tab=output]').click();
  setTimeout(() => window.print(), 100);
});

$('output-download-btn').addEventListener('click', () => {
  if (outputResult) downloadOutputXlsx(outputResult);
});

// 進項
let inputResult = null;
let wenzhongFile = null;
let platformFile = null;
let inputAlwFile = null;

function updateInputCleanBtn() {
  $('input-clean-btn').disabled = !(wenzhongFile && platformFile);
}

bindFileInput('input-wenzhong', 'input-wenzhong-name', file => {
  wenzhongFile = file;
  updateInputCleanBtn();
  $('input-error').textContent = '';
});
bindFileInput('input-platform', 'input-platform-name', file => {
  platformFile = file;
  updateInputCleanBtn();
  $('input-error').textContent = '';
});
bindFileInput('input-alw', 'input-alw-name', file => {
  inputAlwFile = file;
  $('input-error').textContent = '';
});

$('input-clean-btn').addEventListener('click', async () => {
  if (!wenzhongFile || !platformFile) return;
  $('input-error').textContent = '';
  try {
    const [wzWb, pfWb, alwWb] = await Promise.all([
      readWorkbook(wenzhongFile),
      readWorkbook(platformFile),
      inputAlwFile ? readWorkbook(inputAlwFile) : null,
    ]);
    inputResult = cleanInput(wzWb, pfWb, alwWb);
    renderInputResult(inputResult);
  } catch (err) {
    console.error(err);
    $('input-error').textContent = '清洗失敗：' + err.message;
    $('input-result').hidden = true;
    $('input-print-btn').disabled = true;
    $('input-download-btn').disabled = true;
  }
});

$('input-print-btn').addEventListener('click', () => {
  document.querySelector('.tab-btn[data-tab=input]').click();
  setTimeout(() => window.print(), 100);
});

$('input-download-btn').addEventListener('click', () => {
  if (inputResult) downloadInputXlsx(inputResult);
});
