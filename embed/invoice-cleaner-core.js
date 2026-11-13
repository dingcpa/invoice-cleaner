/*!
 * invoice-cleaner-core.js
 * 進銷項資料清洗核心函式 — 純邏輯，無外部相依（無需 SheetJS／DOM）
 *
 * 來源：https://github.com/dingcpa/invoice-cleaner
 * 線上 demo：https://dingcpa.github.io/invoice-cleaner/
 *
 * 主要 API:
 *   InvoiceCleaner.cleanOutputData(invoices, opts)  — 銷項清洗
 *   InvoiceCleaner.cleanInputData(certs, opts)      — 進項清洗
 *
 * 同時匯出輔助函式：
 *   dateToMingguoYM(v)   — 任意日期 → {y, m} 民國年月，無法解析回傳 null
 *   inferPeriod(yms)     — 一組民國年月 → 申報期別字串
 *   fmtMingguoDate(v)    — 任意日期 → '115/03/01' 格式字串
 *
 * 用法：
 *   <script src="invoice-cleaner-core.js"></script>
 *   <script>
 *     const result = InvoiceCleaner.cleanOutputData(invoiceArray);
 *   </script>
 *
 *   或於 Node/CommonJS:
 *     const { cleanOutputData } = require('./invoice-cleaner-core');
 */

(function (global) {
  'use strict';

  // ============================================================
  // 工具函式
  // ============================================================

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
    if (typeof v === 'number') return String(v);
    return String(v).trim();
  }

  /**
   * 解析任意日期值為民國年月 {y, m}；無法解析回傳 null
   * 支援格式：
   *   - Date 物件
   *   - '0115-03-01' / '115-03-01' / '115/03/01'（民國，含分隔符）
   *   - '1150114'（民國 7 碼純數字）
   *   - '2026-03-01'（西元）
   */
  function dateToMingguoYM(v) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date) {
      return { y: v.getFullYear() - 1911, m: v.getMonth() + 1 };
    }
    const s = String(v).trim();
    let m = s.match(/^0?(\d{3})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return { y: +m[1], m: +m[2] };
    m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
    if (m) return { y: +m[1], m: +m[2] };
    m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return { y: +m[1] - 1911, m: +m[2] };
    return null;
  }

  /**
   * 從一組民國年月推算營業稅申報期別字串
   * 同年同月：'115年3月'
   * 同年跨月：'115年3-4月'
   * 跨年：'115年12月 ~ 116年1月'
   */
  function inferPeriod(yms) {
    const valid = (yms || []).filter(Boolean);
    if (!valid.length) return '';
    valid.sort((a, b) => a.y === b.y ? a.m - b.m : a.y - b.y);
    const first = valid[0], last = valid[valid.length - 1];
    if (first.y === last.y && first.m === last.m) return first.y + '年' + first.m + '月';
    if (first.y === last.y) return first.y + '年' + first.m + '-' + last.m + '月';
    return first.y + '年' + first.m + '月 ~ ' + last.y + '年' + last.m + '月';
  }

  /** 任意日期 → '115/03/01' 民國日期字串 */
  function fmtMingguoDate(v) {
    if (!v) return '';
    if (v instanceof Date) {
      const y = v.getFullYear() - 1911;
      const mm = String(v.getMonth() + 1).padStart(2, '0');
      const dd = String(v.getDate()).padStart(2, '0');
      return y + '/' + mm + '/' + dd;
    }
    const s = String(v).trim();
    let m = s.match(/^0?(\d{3})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return m[1] + '/' + String(m[2]).padStart(2, '0') + '/' + String(m[3]).padStart(2, '0');
    m = s.match(/^(\d{3})(\d{2})(\d{2})$/);
    if (m) return m[1] + '/' + m[2] + '/' + m[3];
    m = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (m) return (+m[1] - 1911) + '/' + String(m[2]).padStart(2, '0') + '/' + String(m[3]).padStart(2, '0');
    return s;
  }

  function isPresent(v) {
    return v !== null && v !== undefined && v !== '';
  }

  // ============================================================
  // 折讓單（銷項／進項共用）
  // ============================================================

  /**
   * 把平台折讓明細（一列一品項，表頭金額在同一張折讓單的每一列重複）
   * 依折讓單號碼彙整成一張張折讓單。
   * 金額一律取表頭 amount／tax（品項金額帶小數，加總會有 1 元差）。
   *
   * 每列支援欄位：allowanceNo, status, invoiceNo（原發票號碼）, itemName,
   *   itemAmount, itemTax（品項金額／稅額，銷項折讓逐品項判斷課稅別用）,
   *   amount, tax（必填，折讓單表頭 折讓金額(不含稅)／折讓稅額）, allowanceDate
   */
  function groupAllowances(items, skipNonConfirmed) {
    const docs = new Map();
    const skipped = new Map();
    (items || []).forEach((it, i) => {
      if (!it) return;
      const no = toCleanString(it.allowanceNo) || ('#' + i);
      const status = toCleanString(it.status);
      if (skipNonConfirmed && status && status !== '折讓已確認') {
        skipped.set(no, status);
        return;
      }
      let doc = docs.get(no);
      if (!doc) {
        doc = { allowanceNo: no, date: it.allowanceDate, amount: toNumber(it.amount), tax: toNumber(it.tax), items: [] };
        docs.set(no, doc);
      }
      doc.items.push({
        invoiceNo: toCleanString(it.invoiceNo),
        itemName: toCleanString(it.itemName),
        amount: toNumber(it.itemAmount),
        tax: isPresent(it.itemTax) ? toNumber(it.itemTax) : null,
      });
    });
    return {
      docs: Array.from(docs.values()),
      skipped: Array.from(skipped, ([allowanceNo, status]) => ({ allowanceNo, status })),
    };
  }

  // 課稅別字串 → 'taxable' | 'zero' | 'exempt'；混合、特種或無法判斷 → null
  function categoryOfTaxType(t) {
    if (!t || /混合|特種/.test(t)) return null;
    if (/免稅/.test(t)) return 'exempt';
    if (/零稅/.test(t)) return 'zero';
    if (/應稅/.test(t)) return 'taxable';
    return null;
  }

  /**
   * 判斷單一折讓品項的課稅別。回傳：
   *   { cat: 'taxable'|'zero'|'exempt', guessed } 或 { cat: 'follow' }（跟著同張其他品項）
   *   或 null（品項沒有稅額資料，交給整張折讓單判斷）
   *
   * 優先序：
   *   1. 原發票在本期銷項中且課稅別明確（應稅／零稅率／免稅）→ 依原發票
   *   2. 品項稅額 ≠ 0 → 應稅
   *   3. 品項稅額 = 0 且金額 < 10（含負數折扣列）→ 5% 稅額本來就會四捨五入成 0，無法判斷 → follow
   *   4. 品項稅額 = 0 且金額 ≥ 10 → 不可能是 5% 應稅，暫列免稅（也可能是零稅率）→ guessed
   */
  function classifyAllowanceItem(it, taxTypeOf) {
    const c = categoryOfTaxType(taxTypeOf.get(it.invoiceNo));
    if (c) return { cat: c, guessed: false };
    if (it.tax === null) return null;
    if (it.tax !== 0) return { cat: 'taxable', guessed: false };
    if (it.amount < 10) return { cat: 'follow' };
    return { cat: 'exempt', guessed: true };
  }

  /**
   * 銷項折讓彙總：逐品項判斷課稅別，拆 應稅／零稅率／免稅。
   * 平台折讓檔的金額一律為未稅、稅額另列（B2C 折讓亦同，與 B2C 發票含稅不同），故不做 5% 反推。
   *
   * 每張折讓單：
   *   - 品項課稅別見 classifyAllowanceItem
   *   - 品項沒有稅額資料又查無原發票 → 整張判斷：表頭有稅額 → 應稅；稅額為 0 → 暫列免稅
   *   - follow 品項：同張有應稅品項 → 應稅；否則跟其他品項；全部都是 follow → 應稅
   *   - 金額以表頭為準：零稅率／免稅取品項金額四捨五入，剩餘歸應稅（無應稅時歸最後一類）
   *   - 暫列免稅者列入 unverified
   *   - 核對：|應稅金額 × 5% − 表頭稅額| 超過容許誤差（0.5 × 品項數 + 1）→ 列入 mismatched
   */
  function summarizeOutputAllowances(docs, taxTypeOf) {
    const row = { label: '減：退回及折讓', sales: 0, tax: 0, zero: 0, exempt: 0, total: 0, count: 0 };
    const unverified = [];
    const mismatched = [];
    for (const doc of docs) {
      const fallback = { cat: doc.tax !== 0 ? 'taxable' : 'exempt', guessed: doc.tax === 0 };
      const cls = doc.items.map(it => classifyAllowanceItem(it, taxTypeOf) || fallback);
      const decided = cls.map(x => x.cat).filter(c => c !== 'follow');
      const followAs = decided.includes('taxable') ? 'taxable' : (decided[0] || 'taxable');
      const cats = cls.map(x => (x.cat === 'follow' ? followAs : x.cat));
      if (cls.some((x, i) => x.guessed && cats[i] !== 'taxable')) unverified.push(doc.allowanceNo);

      const part = { taxable: 0, zero: 0, exempt: 0 };
      doc.items.forEach((it, i) => { part[cats[i]] += it.amount; });
      const order = ['zero', 'exempt', 'taxable'].filter(c => cats.includes(c));
      const last = order.pop() || 'taxable';
      let rest = doc.amount;
      for (const c of order) { part[c] = Math.round(part[c]); rest -= part[c]; }
      part[last] = rest;

      if (Math.abs(part.taxable * 0.05 - doc.tax) > 0.5 * doc.items.length + 1) {
        mismatched.push({ allowanceNo: doc.allowanceNo, sales: part.taxable, tax: doc.tax });
      }

      row.sales += part.taxable;
      row.zero += part.zero;
      row.exempt += part.exempt;
      row.tax += doc.tax;
      row.total += doc.amount + doc.tax;
      row.count++;
    }
    return { row, unverified, mismatched };
  }

  /**
   * 進項折讓：依原發票在憑證明細中的扣抵別分為可扣抵／不可扣抵。
   * 原發票不在明細中（前期或紙本）→ 預設可扣抵，matched = false。
   */
  function buildInputAllowances(docs, deductibleOf) {
    const deductible = [], nondeductible = [], unmatched = [];
    for (const doc of docs) {
      const invNos = Array.from(new Set(doc.items.map(it => it.invoiceNo).filter(Boolean)));
      const matchedNo = invNos.find(n => deductibleOf.has(n));
      if (!matchedNo) unmatched.push(doc.allowanceNo);
      const ded = matchedNo ? deductibleOf.get(matchedNo) : 'Y';
      const names = doc.items.map(it => it.itemName).filter(Boolean);
      const record = {
        date: fmtMingguoDate(doc.date),
        invoiceNo: invNos.join('、'),
        itemName: names.length > 1 ? names[0] + ' 等 ' + names.length + ' 項' : (names[0] || ''),
        amount: doc.amount,
        tax: doc.tax,
        deductible: ded,
        allowanceNo: doc.allowanceNo,
        matched: !!matchedNo,
      };
      const list = ded === 'N' ? nondeductible : deductible;
      list.push(Object.assign({ seq: list.length + 1 }, record));
    }
    return { deductible, nondeductible, unmatched };
  }

  function sumAmountTax(list) {
    return list.reduce(
      (acc, r) => ({ amount: acc.amount + r.amount, tax: acc.tax + r.tax }),
      { amount: 0, tax: 0 }
    );
  }

  // ============================================================
  // 銷項清洗
  // ============================================================

  /**
   * 銷項清洗
   *
   * @param {Array<Object>} invoices 銷項發票陣列。每筆物件支援以下欄位：
   *   - invoiceNo (string)         必填，發票號碼
   *   - status (string?)           發票狀態，預設略過非「開立已確認」
   *   - buyerTaxId (string|number) 買方統一編號（'00000000'/'0'/空 → 個人）
   *   - salesAmount (number)       銷售額合計（電子三聯使用）
   *   - taxAmount (number)         營業稅額（電子三聯使用）
   *   - totalAmount (number)       總計（必填，電子二聯反推用）
   *   - taxType (string?)          課稅別（'應稅'/'零稅率'/'免稅'/'特種稅率'…）
   *   - zeroRateAmount (number?)   零稅銷售額（有提供時優先於 taxType）
   *   - exemptAmount (number?)     免稅銷售額（有提供時優先於 taxType）
   *   - invoiceDate (string|Date?) 發票日期（用於推算申報期別）
   *   - sellerTaxId (string?)      自動推測公司資訊用
   *   - sellerName (string?)       自動推測公司資訊用
   * @param {Object} [opts]
   *   - skipNonConfirmed (boolean = true) 是否略過非「開立已確認」／「折讓已確認」
   *   - company ({taxId, name}?) 直接指定公司資訊，覆寫自動推測
   *   - allowances (Array?) 銷項折讓明細（欄位見 groupAllowances），提供時回傳 allowance
   * @returns {{
   *   rows: Array<{label, count, sales, tax, zero, exempt, total}>,
   *   subtotal: {count, sales, tax, zero, exempt, total},
   *   allowance: null | {
   *     row: {label, count, sales, tax, zero, exempt, total},  // count = 折讓單張數
   *     skipped: Array<{allowanceNo, status}>,
   *     unverified: Array<string>,  // 含稅額為 0、查無原發票課稅別、暫列免稅品項的折讓單號
   *     mismatched: Array<{allowanceNo, sales, tax}>  // 應稅金額 × 5% 與表頭稅額不符，需人工核對
   *   },
   *   period: string,
   *   company: {taxId, name},
   *   skipped: Array<{invoiceNo, status}>,
   *   special: Array<{invoiceNo, taxType}>
   * }}
   *   sales = 應稅銷售額（未稅）、zero = 零稅率銷售額、exempt = 免稅銷售額；
   *   小計對應 401/403「收銀機發票(三聯式)及電子發票」列第 5／6／7／8 欄，
   *   allowance.row 對應「減：退回及折讓」第 17／18／19／20 欄。
   *
   * 演算法：
   *   - 買方統編去除非數字後若全為 0（含空字串、'0'、'00000000'）→ 電子二聯（B2C）
   *   - 否則 → 電子三聯
   *   - 每張發票先拆出 零稅率／免稅 金額：
   *       有 zeroRateAmount / exemptAmount → 直接採用（可處理混合課稅別）
   *       否則依 taxType 整張歸類（免稅 / 零稅率），其餘為應稅
   *   - 課稅別含「特種」者不納入，列入 special
   *   - 電子三聯：應稅銷售額 = 銷售額合計 − 零稅率 − 免稅；稅額直接加總
   *   - 電子二聯：應稅總計（含稅）= 總計 − 零稅率 − 免稅，加總後以 5% 反推
   *       銷售額 = round(應稅總計合計 / 1.05)
   *       稅額   = 應稅總計合計 − 銷售額  （確保銷售+稅必等於應稅總計）
   *     零稅率與免稅不反推。
   */
  function cleanOutputData(invoices, opts) {
    opts = opts || {};
    const skipNonConfirmed = opts.skipNonConfirmed !== false;

    let twoTaxableGross = 0;
    const two = { zero: 0, exempt: 0, total: 0, count: 0 };
    const three = { sales: 0, tax: 0, zero: 0, exempt: 0, total: 0, count: 0 };
    const skipped = [];
    const special = [];
    const dates = [];
    let detectedSeller = null;
    const taxTypeOf = new Map(); // 發票號碼 → 課稅別（供折讓單查原發票）

    for (const inv of (invoices || [])) {
      if (!inv) continue;
      const invoiceNo = toCleanString(inv.invoiceNo);
      if (!invoiceNo) continue;
      taxTypeOf.set(invoiceNo, toCleanString(inv.taxType));

      const status = toCleanString(inv.status);
      if (skipNonConfirmed && status && status !== '開立已確認') {
        skipped.push({ invoiceNo, status });
        continue;
      }

      const buyerTaxRaw = toCleanString(inv.buyerTaxId).replace(/\.0+$/, '');
      const buyerTax = buyerTaxRaw.replace(/\D/g, '');
      const total = toNumber(inv.totalAmount);
      const rawSales = toNumber(inv.salesAmount);
      const rawTax = toNumber(inv.taxAmount);

      const ym = dateToMingguoYM(inv.invoiceDate);
      if (ym) dates.push(ym);

      // 從第一筆有效記錄抓賣方資訊
      if (!detectedSeller) {
        const taxId = toCleanString(inv.sellerTaxId).replace(/\.0+$/, '').replace(/\D/g, '');
        const name = toCleanString(inv.sellerName);
        if (taxId && name && !/^0+$/.test(taxId)) {
          detectedSeller = { taxId, name };
        }
      }

      const taxType = toCleanString(inv.taxType);
      if (/特種/.test(taxType)) {
        special.push({ invoiceNo, taxType });
        continue;
      }
      const { zero, exempt } = splitNonTaxable({
        taxType,
        zeroCol: isPresent(inv.zeroRateAmount) ? toNumber(inv.zeroRateAmount) : null,
        exemptCol: isPresent(inv.exemptAmount) ? toNumber(inv.exemptAmount) : null,
        total,
        tax: rawTax,
      });

      const isTwoCopy = !buyerTax || /^0+$/.test(buyerTax);
      if (isTwoCopy) {
        twoTaxableGross += total - zero - exempt;
        two.zero += zero;
        two.exempt += exempt;
        two.total += total;
        two.count++;
      } else {
        three.sales += rawSales - zero - exempt;
        three.tax += rawTax;
        three.zero += zero;
        three.exempt += exempt;
        three.total += total;
        three.count++;
      }
    }

    two.sales = Math.round(twoTaxableGross / 1.05);
    two.tax = twoTaxableGross - two.sales;

    let allowance = null;
    if (opts.allowances) {
      const g = groupAllowances(opts.allowances, skipNonConfirmed);
      const s = summarizeOutputAllowances(g.docs, taxTypeOf);
      allowance = { row: s.row, skipped: g.skipped, unverified: s.unverified, mismatched: s.mismatched };
    }

    return {
      rows: [
        { label: '電子二聯', ...two },
        { label: '電子三聯', ...three },
      ],
      subtotal: {
        sales: two.sales + three.sales,
        tax: two.tax + three.tax,
        zero: two.zero + three.zero,
        exempt: two.exempt + three.exempt,
        total: two.total + three.total,
        count: two.count + three.count,
      },
      allowance,
      period: inferPeriod(dates),
      company: opts.company || detectedSeller || { taxId: '', name: '' },
      skipped,
      special,
    };
  }

  // 將一張發票的金額拆成 零稅率／免稅 兩部分（其餘即為應稅）。
  // 有 零稅銷售額／免稅銷售額 欄位時直接採用；否則依課稅別整張歸類
  // （混合課稅別而沒有拆分欄位時無從拆起，整張視為應稅）。
  function splitNonTaxable(p) {
    if (p.zeroCol !== null || p.exemptCol !== null) {
      return { zero: p.zeroCol || 0, exempt: p.exemptCol || 0 };
    }
    const c = categoryOfTaxType(p.taxType);
    if (c === 'exempt') return { zero: 0, exempt: p.total - p.tax };
    if (c === 'zero') return { zero: p.total - p.tax, exempt: 0 };
    return { zero: 0, exempt: 0 };
  }

  // ============================================================
  // 進項清洗
  // ============================================================

  /**
   * 進項清洗
   *
   * @param {Array<Object>} certs 文中系統憑證明細陣列。每筆物件支援以下欄位：
   *   - invoiceNo (string)        必填，發票號碼
   *   - date (string|Date?)       憑證日期（民國/西元/Date 物件皆可）
   *   - amount (number)           憑證金額
   *   - taxAmount (number)        稅額
   *   - deductible (string)       'Y' / 'N'，預設 'Y'
   *   - itemName (string?)        品名（若 platformItemNames 沒提供時使用）
   * @param {Object} [opts]
   *   - platformInvoiceNos (Set|Array<string>?) 電子發票平台已下載的發票號碼集合；
   *     若提供，會排除不在此集合內的列（紙本／自行輸入），並計入 paperCount
   *   - platformItemNames (Map|Object<string, string>?) 發票號碼 → 品名 對應；
   *     若有則優先使用平台明細品名（典型來自平台 Invoice_details 第一筆）
   *   - company ({taxId, name}?) 公司資訊（無自動推測能力，建議呼叫端傳入）
   *   - allowances (Array?) 進項折讓明細（欄位見 groupAllowances），提供時回傳 allowance
   *   - skipNonConfirmed (boolean = true) 折讓單是否略過非「折讓已確認」
   * @returns {{
   *   deductible: {rows, subtotal: {amount, tax}},
   *   nondeductible: {rows, subtotal: {amount, tax}},
   *   allowance: null | {
   *     deductible: {rows, subtotal},      // 可扣抵進項之退出及折讓 → 401/403 第 40／41 欄
   *     nondeductible: {rows, subtotal},
   *     skipped: Array<{allowanceNo, status}>,
   *     unmatched: Array<string>           // 原發票不在憑證明細中、預設可扣抵的折讓單號
   *   },
   *   period: string,
   *   company: {taxId, name},
   *   paperCount: number,
   *   missingNames: number
   * }}
   *   折讓 row：{ seq, date, invoiceNo（原發票號碼）, itemName, amount, tax, deductible,
   *              allowanceNo, matched }；扣抵別取原發票在 certs 中的扣抵欄。
   */
  function cleanInputData(certs, opts) {
    opts = opts || {};
    const platformSet = normalizeToSet(opts.platformInvoiceNos);
    const nameMap = normalizeToMap(opts.platformItemNames);

    const deductible = [];
    const nondeductible = [];
    let dedSeq = 0, ndSeq = 0;
    let paperCount = 0;
    const dates = [];
    const deductibleOf = new Map(); // 發票號碼 → 'Y'/'N'（含紙本，供折讓單查原發票）

    for (const cert of (certs || [])) {
      if (!cert) continue;
      const invoiceNo = toCleanString(cert.invoiceNo);
      if (!invoiceNo) continue;
      deductibleOf.set(invoiceNo, toCleanString(cert.deductible).toUpperCase() === 'N' ? 'N' : 'Y');

      // 過濾紙本：發票號碼不在平台集合中 → 排除
      if (platformSet && platformSet.size > 0 && !platformSet.has(invoiceNo)) {
        paperCount++;
        continue;
      }

      const date = fmtMingguoDate(cert.date);
      const ym = dateToMingguoYM(cert.date);
      if (ym) dates.push(ym);

      const amount = toNumber(cert.amount);
      const tax = toNumber(cert.taxAmount);
      const ded = toCleanString(cert.deductible).toUpperCase();

      let itemName = '';
      if (nameMap && nameMap.has(invoiceNo)) itemName = nameMap.get(invoiceNo) || '';
      if (!itemName) itemName = toCleanString(cert.itemName);

      const record = { date, invoiceNo, itemName, amount, tax, deductible: ded };
      if (ded === 'N') {
        nondeductible.push({ seq: ++ndSeq, ...record });
      } else {
        deductible.push({ seq: ++dedSeq, ...record });
      }
    }

    let allowance = null;
    if (opts.allowances) {
      const g = groupAllowances(opts.allowances, opts.skipNonConfirmed !== false);
      const a = buildInputAllowances(g.docs, deductibleOf);
      allowance = {
        deductible: { rows: a.deductible, subtotal: sumAmountTax(a.deductible) },
        nondeductible: { rows: a.nondeductible, subtotal: sumAmountTax(a.nondeductible) },
        skipped: g.skipped,
        unmatched: a.unmatched,
      };
    }

    return {
      deductible: { rows: deductible, subtotal: sumAmountTax(deductible) },
      nondeductible: { rows: nondeductible, subtotal: sumAmountTax(nondeductible) },
      allowance,
      period: inferPeriod(dates),
      company: opts.company || { taxId: '', name: '' },
      paperCount,
      missingNames: deductible.concat(nondeductible).filter(r => !r.itemName).length,
    };
  }

  function normalizeToSet(input) {
    if (!input) return null;
    if (input instanceof Set) return input;
    if (Array.isArray(input)) return new Set(input.map(toCleanString).filter(Boolean));
    return null;
  }

  function normalizeToMap(input) {
    if (!input) return null;
    if (input instanceof Map) return input;
    if (typeof input === 'object') {
      const m = new Map();
      for (const k of Object.keys(input)) m.set(k, input[k]);
      return m;
    }
    return null;
  }

  // ============================================================
  // 匯出
  // ============================================================

  const api = {
    cleanOutputData,
    cleanInputData,
    dateToMingguoYM,
    inferPeriod,
    fmtMingguoDate,
    version: '1.3.0',
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.InvoiceCleaner = api;
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
