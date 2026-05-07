/* 進銷項資料清洗 — 瀏覽器端清洗邏輯 */

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
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return String(v);
    return String(v);
  }
  return String(v).trim();
}

function fmtNum(n) {
  if (!n && n !== 0) return '';
  return Math.round(n).toLocaleString('en-US');
}

// 民國年日期格式化：支援 '0115-03-01' / '1150114' / Date 物件
function fmtMingguoDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear() - 1911;
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  const s = String(v).trim();
  // 格式: 0115-03-01 或 115-03-01 或 115/03/01
  let m = s.match(/^0?(\d{3})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) {
    return `${m[1]}/${String(m[2]).padStart(2,'0')}/${String(m[3]).padStart(2,'0')}`;
  }
  // 格式: 1150114 (民國年7位純數字)
  m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
  if (m) {
    return `${m[1]}/${m[2]}/${m[3]}`;
  }
  return s;
}

// 解析任意日期值為民國年月 {y, m}；無法解析回傳 null
function dateToMingguoYM(v) {
  if (v === null || v === undefined || v === '') return null;
  if (v instanceof Date) {
    return { y: v.getFullYear() - 1911, m: v.getMonth() + 1 };
  }
  const s = String(v).trim();
  // 民國: 0115-03-01 / 115-03-01 / 115/03/01
  let m = s.match(/^0?(\d{3})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return { y: +m[1], m: +m[2] };
  // 民國: 1150114 (7碼)
  m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
  if (m) return { y: +m[1], m: +m[2] };
  // 西元: 2026-03-01 / 2026/03/01
  m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return { y: +m[1] - 1911, m: +m[2] };
  return null;
}

// 從一組日期推算營業稅申報期別字串，例 '115年1-2月' 或 '115年3月'
function inferPeriod(yms) {
  const valid = yms.filter(Boolean);
  if (!valid.length) return '';
  valid.sort((a, b) => a.y === b.y ? a.m - b.m : a.y - b.y);
  const first = valid[0], last = valid[valid.length - 1];
  if (first.y === last.y && first.m === last.m) return `${first.y}年${first.m}月`;
  if (first.y === last.y) return `${first.y}年${first.m}-${last.m}月`;
  return `${first.y}年${first.m}月 ~ ${last.y}年${last.m}月`;
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
// 銷項清洗
// ============================================================

const OUTPUT_HEADER_MAP = {
  invoiceNo: ['發票號碼'],
  formatCode: ['格式代號'],
  status: ['發票狀態'],
  invoiceDate: ['發票日期'],
  buyerTaxId: ['買方統一編號', '買方統編'],
  buyerName: ['買方名稱'],
  sellerTaxId: ['賣方統一編號', '賣方統編'],
  sellerName: ['賣方名稱'],
  salesAmount: ['銷售額合計'],
  taxAmount: ['營業稅'],
  totalAmount: ['總計'],
};

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

function cleanOutput(wb) {
  // 嘗試找到 Invoice 工作表（或第一張）
  let ws = wb.Sheets['Invoice'];
  if (!ws) ws = wb.Sheets[wb.SheetNames[0]];
  if (!ws) throw new Error('找不到工作表');

  const rows = sheetToRows(ws);
  if (rows.length === 0) throw new Error('檔案無資料');

  // 找標頭列
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    if (rows[i] && rows[i].some(c => c === '發票號碼')) {
      headerRowIdx = i;
      break;
    }
  }
  if (headerRowIdx < 0) throw new Error('找不到「發票號碼」標頭欄，請確認檔案是否為平台下載的銷項發票明細');

  const cols = findColumns(rows[headerRowIdx], OUTPUT_HEADER_MAP);
  const required = ['invoiceNo', 'totalAmount', 'salesAmount', 'taxAmount'];
  const missing = required.filter(k => cols[k] < 0);
  if (missing.length) throw new Error('缺少必要欄位：' + missing.join(', '));

  let twoTotalSum = 0, twoCount = 0;
  const three = { sales: 0, tax: 0, total: 0, count: 0 };
  const skipped = [];
  const dates = [];

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c === null || c === '')) continue;
    const invoiceNo = toCleanString(row[cols.invoiceNo]);
    if (!invoiceNo) continue;

    const status = cols.status >= 0 ? toCleanString(row[cols.status]) : '';
    if (status && status !== '開立已確認') {
      skipped.push({ invoiceNo, status });
      continue;
    }

    const buyerTaxRaw = cols.buyerTaxId >= 0 ? toCleanString(row[cols.buyerTaxId]).replace(/\.0+$/, '') : '';
    const buyerTax = buyerTaxRaw.replace(/\D/g, ''); // 只留數字
    const total = toNumber(row[cols.totalAmount]);
    const rawSales = toNumber(row[cols.salesAmount]);
    const rawTax = toNumber(row[cols.taxAmount]);

    if (cols.invoiceDate >= 0) {
      const ym = dateToMingguoYM(row[cols.invoiceDate]);
      if (ym) dates.push(ym);
    }

    // 無統編 = 空字串、純 0（如 '0'、'00000000'）— 視為電子二聯（B2C）
    const isTwoCopy = !buyerTax || /^0+$/.test(buyerTax);

    if (isTwoCopy) {
      // 電子二聯：先把總計累加，最後一次反推
      twoTotalSum += total;
      twoCount++;
    } else {
      three.sales += rawSales;
      three.tax += rawTax;
      three.total += total;
      three.count++;
    }
  }

  // 電子二聯彙總：銷售額 = round(總計合計 / 1.05)、稅額 = round(銷售額 × 5%)
  const twoSales = Math.round(twoTotalSum / 1.05);
  const twoTax = Math.round(twoSales * 0.05);
  const two = {
    sales: twoSales,
    tax: twoTax,
    total: twoTotalSum,
    count: twoCount,
  };

  // 公司資訊：銷項取賣方
  const company = findCompanyFromPlatform(wb, 'seller') || { taxId: '', name: '' };

  return {
    rows: [
      { label: '電子二聯', ...two },
      { label: '電子三聯', ...three },
    ],
    subtotal: {
      sales: two.sales + three.sales,
      tax: two.tax + three.tax,
      total: two.total + three.total,
      count: two.count + three.count,
    },
    period: inferPeriod(dates),
    company,
    skipped,
  };
}

function renderOutputResult(result) {
  const meta = $('output-meta');
  const today = new Date();
  const stamp = `${today.getFullYear() - 1911}/${String(today.getMonth()+1).padStart(2,'0')}/${String(today.getDate()).padStart(2,'0')}`;
  $('output-period').textContent = result.period || '';
  const c = result.company || {};
  $('output-company').innerHTML = c.name
    ? `${c.name}${c.taxId ? `<span class="tax-id">統一編號 ${c.taxId}</span>` : ''}`
    : '';
  meta.textContent = `製表日期：${stamp}　共 ${result.subtotal.count} 筆${result.skipped.length ? `（已略過非開立確認 ${result.skipped.length} 筆）` : ''}`;

  const table = $('output-table');
  let html = `
    <thead>
      <tr>
        <th class="center">類別</th>
        <th class="num">筆數</th>
        <th class="num">銷售額</th>
        <th class="num">稅額</th>
        <th class="num">合計</th>
      </tr>
    </thead>
    <tbody>
  `;
  for (const r of result.rows) {
    html += `<tr>
      <td class="center">${r.label}</td>
      <td class="num">${r.count}</td>
      <td class="num">${fmtNum(r.sales)}</td>
      <td class="num">${fmtNum(r.tax)}</td>
      <td class="num">${fmtNum(r.total)}</td>
    </tr>`;
  }
  html += `<tr class="grand-total">
    <td class="center">小計</td>
    <td class="num">${result.subtotal.count}</td>
    <td class="num">${fmtNum(result.subtotal.sales)}</td>
    <td class="num">${fmtNum(result.subtotal.tax)}</td>
    <td class="num">${fmtNum(result.subtotal.total)}</td>
  </tr>`;
  html += `</tbody>`;
  table.innerHTML = html;

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
    font: { name: '微軟正黑體', sz: 12, bold: true, color: { rgb: 'B91C1C' } },
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
  const headers = ['類別', '筆數', '銷售額', '稅額', '合計'];
  headers.forEach((h, c) => setCell(ws, r, c, makeCell(h, STYLE.header)));
  r++;
  // 資料列
  for (const row of result.rows) {
    setCell(ws, r, 0, makeCell(row.label, STYLE.cellCenter));
    setCell(ws, r, 1, makeCell(row.count, STYLE.cellNum));
    setCell(ws, r, 2, makeCell(row.sales, STYLE.cellNum));
    setCell(ws, r, 3, makeCell(row.tax, STYLE.cellNum));
    setCell(ws, r, 4, makeCell(row.total, STYLE.cellNum));
    r++;
  }
  // 小計
  setCell(ws, r, 0, makeCell('小計', STYLE.grandTotalLabel));
  setCell(ws, r, 1, makeCell(result.subtotal.count, STYLE.grandTotalNum));
  setCell(ws, r, 2, makeCell(result.subtotal.sales, STYLE.grandTotalNum));
  setCell(ws, r, 3, makeCell(result.subtotal.tax, STYLE.grandTotalNum));
  setCell(ws, r, 4, makeCell(result.subtotal.total, STYLE.grandTotalNum));

  ensureRange(ws, r, 4);
  ws['!cols'] = [{ wch: 12 }, { wch: 8 }, { wch: 14 }, { wch: 12 }, { wch: 14 }];
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

function cleanInput(wenzhongWb, platformWb) {
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

  const nameMap = buildItemNameMap(platformWb);
  const platformInvoiceSet = buildPlatformInvoiceSet(platformWb);

  const deductible = []; // 可扣抵
  const nondeductible = []; // 不可扣抵
  let dedSeq = 0, ndSeq = 0;
  let paperCount = 0; // 排除的紙本/非電子發票筆數
  const dates = [];

  for (let i = headerIdx + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    // 跳過區塊註記列（單純文字、無金額）
    const noVal = row[cols.no];
    if (noVal === null || noVal === undefined || noVal === '') continue;
    if (typeof noVal === 'string' && !/^\d/.test(noVal.trim())) continue;

    const invoiceNo = toCleanString(row[cols.invoiceNo]);
    if (!invoiceNo) continue;

    // 過濾紙本發票：發票號碼不在電子發票平台下載清單中 → 排除
    // （若平台 Invoice 集合為空，跳過此過濾以策安全）
    if (platformInvoiceSet.size > 0 && !platformInvoiceSet.has(invoiceNo)) {
      paperCount++;
      continue;
    }

    const date = fmtMingguoDate(row[cols.date]);
    const ym = dateToMingguoYM(row[cols.date]);
    if (ym) dates.push(ym);
    const amount = toNumber(row[cols.amount]);
    const tax = toNumber(row[cols.taxAmount]);
    const ded = toCleanString(row[cols.deductible]).toUpperCase();
    let itemName = nameMap.get(invoiceNo) || '';
    if (!itemName && cols.creditMemo >= 0) itemName = toCleanString(row[cols.creditMemo]);
    if (!itemName && cols.debitMemo >= 0) itemName = toCleanString(row[cols.debitMemo]);

    const record = {
      date, invoiceNo, itemName, amount, tax, deductible: ded,
    };

    if (ded === 'Y') {
      deductible.push({ seq: ++dedSeq, ...record });
    } else if (ded === 'N') {
      nondeductible.push({ seq: ++ndSeq, ...record });
    } else {
      // 預設為可扣抵
      deductible.push({ seq: ++dedSeq, ...record });
    }
  }

  const sumOf = list => list.reduce((acc, r) => ({
    amount: acc.amount + r.amount,
    tax: acc.tax + r.tax,
  }), { amount: 0, tax: 0 });

  // 公司資訊：先從文中檔的「用戶：」列抓，否則從平台檔的買方名稱抓
  const wzRows = sheetToRows(ws);
  const wzCompanyName = findCompanyFromWenzhong(wzRows);
  const platformBuyer = findCompanyFromPlatform(platformWb, 'buyer');
  let companyName = wzCompanyName;
  if (!companyName && platformBuyer && platformBuyer.name && platformBuyer.name !== platformBuyer.taxId) {
    companyName = platformBuyer.name;
  }
  const company = {
    name: companyName,
    taxId: (platformBuyer && platformBuyer.taxId) || extractTaxIdFromFile(wenzhongWb, wzRows),
  };

  return {
    deductible: { rows: deductible, subtotal: sumOf(deductible) },
    nondeductible: { rows: nondeductible, subtotal: sumOf(nondeductible) },
    period: inferPeriod(dates),
    company,
    paperCount,
    missingNames: deductible.concat(nondeductible).filter(r => !r.itemName).length,
  };
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
      <td class="center">${r.deductible || ''}</td>
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
  const meta = $('input-meta');
  const today = new Date();
  const stamp = `${today.getFullYear() - 1911}/${String(today.getMonth()+1).padStart(2,'0')}/${String(today.getDate()).padStart(2,'0')}`;
  const totalCount = result.deductible.rows.length + result.nondeductible.rows.length;
  $('input-period').textContent = result.period || '';
  const c = result.company || {};
  $('input-company').innerHTML = c.name
    ? `${c.name}${c.taxId ? `<span class="tax-id">統一編號 ${c.taxId}</span>` : ''}`
    : '';
  const notes = [];
  if (result.paperCount) notes.push(`已排除非電子發票 ${result.paperCount} 筆`);
  if (result.missingNames) notes.push(`其中 ${result.missingNames} 筆查無平台明細品名`);
  meta.textContent = `製表日期：${stamp}　共 ${totalCount} 筆${notes.length ? `（${notes.join('；')}）` : ''}`;

  const container = $('input-tables');
  container.innerHTML =
    renderInputSection('可扣抵進項', result.deductible) +
    renderInputSection('不可扣抵進項', result.nondeductible);

  $('input-result').hidden = false;
  $('input-print-btn').disabled = false;
  $('input-download-btn').disabled = false;
}

function downloadInputXlsx(result) {
  const ws = {};
  const rowHeights = [];
  let r = 0;

  // 公司名稱（如有）
  if (result.company && result.company.name) {
    const cn = result.company;
    const txt = `公司名稱：${cn.name}${cn.taxId ? `（統一編號 ${cn.taxId}）` : ''}`;
    setCell(ws, r, 0, makeCell(txt, STYLE.companyLine));
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r, c: 0 }, e: { r, c: 6 } });
    rowHeights[r] = { hpt: 22 };
    r++;
  }
  // 申報期別（如有）
  if (result.period) {
    setCell(ws, r, 0, makeCell(`申報期別：${result.period}`, STYLE.periodLine));
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r, c: 0 }, e: { r, c: 6 } });
    rowHeights[r] = { hpt: 22 };
    r++;
  }
  if (r > 0) r++; // 空一列

  function writeSection(title, rows, subtotal) {
    // 區段標題（合併 7 欄）
    setCell(ws, r, 0, makeCell(title, STYLE.sectionTitle));
    if (!ws['!merges']) ws['!merges'] = [];
    ws['!merges'].push({ s: { r, c: 0 }, e: { r, c: 6 } });
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
      setCell(ws, r, 6, makeCell(row.deductible, STYLE.cellCenter));
      r++;
    }
    // 小計列：合併前 4 欄為「小計」標籤
    setCell(ws, r, 0, makeCell('小計', STYLE.subtotalLabel));
    setCell(ws, r, 1, makeCell('', STYLE.subtotalLabel));
    setCell(ws, r, 2, makeCell('', STYLE.subtotalLabel));
    setCell(ws, r, 3, makeCell('', STYLE.subtotalLabel));
    ws['!merges'].push({ s: { r, c: 0 }, e: { r, c: 3 } });
    setCell(ws, r, 4, makeCell(subtotal.amount, STYLE.subtotalNum));
    setCell(ws, r, 5, makeCell(subtotal.tax, STYLE.subtotalNum));
    setCell(ws, r, 6, makeCell('', STYLE.subtotalLabel));
    rowHeights[r] = { hpt: 22 };
    r++;
    // 區段間空一列
    r++;
  }

  writeSection('可扣抵進項', result.deductible.rows, result.deductible.subtotal);
  writeSection('不可扣抵進項', result.nondeductible.rows, result.nondeductible.subtotal);

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
bindFileInput('output-file', 'output-file-name', file => {
  $('output-clean-btn').disabled = !file;
  $('output-error').textContent = '';
});

$('output-clean-btn').addEventListener('click', async () => {
  const file = $('output-file').files[0];
  if (!file) return;
  $('output-error').textContent = '';
  try {
    const wb = await readWorkbook(file);
    outputResult = cleanOutput(wb);
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

$('input-clean-btn').addEventListener('click', async () => {
  if (!wenzhongFile || !platformFile) return;
  $('input-error').textContent = '';
  try {
    const [wzWb, pfWb] = await Promise.all([
      readWorkbook(wenzhongFile),
      readWorkbook(platformFile),
    ]);
    inputResult = cleanInput(wzWb, pfWb);
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
