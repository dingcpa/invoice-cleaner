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
    return s;
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
   *   - invoiceDate (string|Date?) 發票日期（用於推算申報期別）
   *   - sellerTaxId (string?)      自動推測公司資訊用
   *   - sellerName (string?)       自動推測公司資訊用
   * @param {Object} [opts]
   *   - skipNonConfirmed (boolean = true) 是否略過非「開立已確認」
   *   - company ({taxId, name}?) 直接指定公司資訊，覆寫自動推測
   * @returns {{
   *   rows: Array<{label, count, sales, tax, total}>,
   *   subtotal: {count, sales, tax, total},
   *   period: string,
   *   company: {taxId, name},
   *   skipped: Array<{invoiceNo, status}>
   * }}
   *
   * 演算法：
   *   - 買方統編去除非數字後若全為 0（含空字串、'0'、'00000000'）→ 電子二聯（B2C）
   *   - 否則 → 電子三聯
   *   - 電子三聯：直接加總原始 銷售額 / 稅額 / 合計
   *   - 電子二聯：先加總所有總計，再以 5% 稅率反推
   *       銷售額 = round(總計合計 / 1.05)
   *       稅額   = 總計合計 − 銷售額  （確保銷售+稅必等於合計）
   */
  function cleanOutputData(invoices, opts) {
    opts = opts || {};
    const skipNonConfirmed = opts.skipNonConfirmed !== false;

    let twoTotalSum = 0, twoCount = 0;
    const three = { sales: 0, tax: 0, total: 0, count: 0 };
    const skipped = [];
    const dates = [];
    let detectedSeller = null;

    for (const inv of (invoices || [])) {
      if (!inv) continue;
      const invoiceNo = toCleanString(inv.invoiceNo);
      if (!invoiceNo) continue;

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

      const isTwoCopy = !buyerTax || /^0+$/.test(buyerTax);
      if (isTwoCopy) {
        twoTotalSum += total;
        twoCount++;
      } else {
        three.sales += rawSales;
        three.tax += rawTax;
        three.total += total;
        three.count++;
      }
    }

    const twoSales = Math.round(twoTotalSum / 1.05);
    const twoTax = twoTotalSum - twoSales;
    const two = { sales: twoSales, tax: twoTax, total: twoTotalSum, count: twoCount };

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
      company: opts.company || detectedSeller || { taxId: '', name: '' },
      skipped,
    };
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
   * @returns {{
   *   deductible: {rows, subtotal: {amount, tax}},
   *   nondeductible: {rows, subtotal: {amount, tax}},
   *   period: string,
   *   company: {taxId, name},
   *   paperCount: number,
   *   missingNames: number
   * }}
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

    for (const cert of (certs || [])) {
      if (!cert) continue;
      const invoiceNo = toCleanString(cert.invoiceNo);
      if (!invoiceNo) continue;

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

    const sumOf = list => list.reduce(
      (acc, r) => ({ amount: acc.amount + r.amount, tax: acc.tax + r.tax }),
      { amount: 0, tax: 0 }
    );

    return {
      deductible: { rows: deductible, subtotal: sumOf(deductible) },
      nondeductible: { rows: nondeductible, subtotal: sumOf(nondeductible) },
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
    version: '1.0.0',
  };

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = api;
  } else {
    global.InvoiceCleaner = api;
  }
})(typeof window !== 'undefined' ? window : (typeof globalThis !== 'undefined' ? globalThis : this));
