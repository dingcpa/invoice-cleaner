# invoice-cleaner-core

進銷項資料清洗的核心邏輯，抽出為**單一 JS 檔、無外部相依**版本，供其他發票系統嵌入使用。

不需要 SheetJS、不操作 DOM；輸入 = 純 JS 陣列／物件，輸出 = 結構化結果物件。

線上完整版（含上傳 xlsx／列印／下載 Excel）：<https://dingcpa.github.io/invoice-cleaner/> —— 網站本身也是直接呼叫這個檔案，兩者結果必然一致。

---

## 安裝

只需要一個檔案：[`invoice-cleaner-core.js`](./invoice-cleaner-core.js)

### 瀏覽器（直接 `<script>` 引入）

```html
<script src="invoice-cleaner-core.js"></script>
<script>
  const result = InvoiceCleaner.cleanOutputData(myInvoiceArray);
  console.log(result);
</script>
```

### Node / CommonJS

```js
const { cleanOutputData, cleanInputData } = require('./invoice-cleaner-core');
```

### ES Module（如需）

把檔案末段的匯出改成 `export const InvoiceCleaner = api;` 即可。

---

## API

### 1. `cleanOutputData(invoices, opts?)` — 銷項清洗

把銷項發票陣列依買方統編分為電子二聯／電子三聯，並依課稅別拆出應稅／零稅率／免稅，彙總筆數、應稅銷售額、稅額、零稅率銷售額、免稅銷售額、合計。

#### 輸入：`invoices` — Array

每筆物件可含：

| 欄位 | 型別 | 說明 |
|---|---|---|
| `invoiceNo` | string | **必填**，發票號碼 |
| `status` | string? | 發票狀態，預設略過非「開立已確認」 |
| `buyerTaxId` | string\|number | 買方統編，全為 0（含空字串、`'00000000'`）→ 個人 |
| `salesAmount` | number | 銷售額合計（電子三聯使用） |
| `taxAmount` | number | 營業稅額（電子三聯使用） |
| `totalAmount` | number | 總計（必填，電子二聯反推用） |
| `taxType` | string? | 課稅別（`'應稅'`／`'零稅率'`／`'免稅'`／`'特種稅率'`…） |
| `zeroRateAmount` | number? | 零稅銷售額；有提供時優先於 `taxType`（可處理混合課稅別發票） |
| `exemptAmount` | number? | 免稅銷售額；有提供時優先於 `taxType` |
| `invoiceDate` | string\|Date? | 發票日期，用於推算申報期別 |
| `sellerTaxId` | string? | 自動推測公司資訊用 |
| `sellerName` | string? | 自動推測公司資訊用 |

#### 選項：`opts`

| 欄位 | 預設 | 說明 |
|---|---|---|
| `skipNonConfirmed` | `true` | 是否略過非「開立已確認」的發票／非「折讓已確認」的折讓單 |
| `company` | – | 直接指定 `{taxId, name}`，覆寫自動推測 |
| `allowances` | – | 銷項折讓明細陣列（格式見下方「[折讓明細](#3-折讓明細allowances)」）；提供時回傳 `allowance` |

#### 回傳

```js
{
  rows: [
    { label: '電子二聯', count, sales, tax, zero, exempt, total },
    { label: '電子三聯', count, sales, tax, zero, exempt, total },
  ],
  subtotal: { count, sales, tax, zero, exempt, total },
  allowance: {                             // 未傳 opts.allowances 時為 null
    row: { label: '減：退回及折讓', count, sales, tax, zero, exempt, total },  // count = 折讓單張數
    skipped: [{ allowanceNo, status }],    // 被略過的作廢折讓單
    unverified: ['折讓單號', ...],          // 含稅額為 0、查無原發票課稅別、暫列免稅品項者
    mismatched: [{ allowanceNo, sales, tax }],  // 應稅金額 × 5% 與表頭稅額不符，需人工核對
  },
  period:  '115年3-4月',                  // 自動從 invoiceDate 推算
  company: { taxId: '...', name: '...' },
  skipped: [{ invoiceNo, status }],       // 被略過的非開立確認列
  special: [{ invoiceNo, taxType }],      // 特種稅率發票，不納入本表
}
```

- `sales` = 應稅銷售額（未稅）、`tax` = 稅額、`zero` = 零稅率銷售額、`exempt` = 免稅銷售額、`total` = 總計合計
- `subtotal` 對應 401／403 申報書「收銀機發票(三聯式)及電子發票」列：第 5／6／7／8 欄
- `allowance.row` 對應「減：退回及折讓」：第 17／18／19／20 欄（另有紙本發票折讓須自行合併）
- ⚠️ 1.1.0 起 `sales` 不再含免稅／零稅率金額（1.0.0 的電子三聯 `sales` 為銷售額合計，含免稅）

#### 演算法

- 買方統編去除非數字後若全為 0（含 `''`、`'0'`、`'00000000'`）→ **電子二聯（B2C）**
- 否則 → **電子三聯**
- 每張發票先拆出零稅率／免稅金額：有 `zeroRateAmount`／`exemptAmount` 就直接採用；否則依 `taxType` 整張歸類，其餘為應稅
- 課稅別含「特種」者不納入，列入 `special`
- 電子三聯：應稅銷售額 = `銷售額合計 − 零稅率 − 免稅`；稅額直接加總
- 電子二聯：個人買受人的應稅銷售額欄是**含稅**金額（稅額欄為 0），所以只有應稅部分要反推
  - 應稅總計 = `總計 − 零稅率 − 免稅`，全部加總後：
  - 銷售額 = `round(應稅總計合計 / 1.05)`
  - 稅額 = `應稅總計合計 − 銷售額`（保證銷售+稅必等於應稅總計）
  - 零稅率與免稅**不反推**，原額列入
- 銷項折讓：**逐品項**判斷課稅別，拆應稅／零稅率／免稅（一張折讓單可能同時有應稅與免稅品項）
  - 平台折讓檔的金額一律是**未稅**、稅額另列，B2C 折讓也一樣（與 B2C 發票含稅不同），所以折讓**不做 5% 反推**
  - 每個品項依序判斷：
    1. 原發票在 `invoices` 中且課稅別明確 → 依原發票（應稅品項即使小額稅額四捨五入為 0，仍列應稅）
    2. 品項稅額（`itemTax`）≠ 0 → 應稅
    3. 品項稅額 = 0 且品項金額 < 10（含負數折扣列）→ 5% 本來就會四捨五入成 0，跟著同張其他品項（有應稅品項就歸應稅）
    4. 品項稅額 = 0 且品項金額 ≥ 10 → 不可能是 5% 應稅，暫列免稅並列入 `unverified`（也可能是零稅率）
    5. 沒傳 `itemTax` 又查無原發票 → 整張判斷：表頭有稅額 → 應稅；稅額為 0 → 暫列免稅並列入 `unverified`
  - 金額以表頭為準：零稅率／免稅取品項金額四捨五入，剩餘歸應稅
  - 核對：`|應稅金額 × 5% − 表頭稅額| > 0.5 × 品項數 + 1` 的折讓單列入 `mismatched`

---

### 2. `cleanInputData(certs, opts?)` — 進項清洗

把文中系統憑證明細依扣抵欄分為可扣抵／不可扣抵；可選擇用平台檔的發票集合過濾紙本、用平台明細補品名。

#### 輸入：`certs` — Array

每筆物件可含：

| 欄位 | 型別 | 說明 |
|---|---|---|
| `invoiceNo` | string | **必填**，發票號碼 |
| `date` | string\|Date? | 憑證日期（民國 `'115/03/01'`／西元／7 碼純數字／Date 皆可） |
| `amount` | number | 憑證金額 |
| `taxAmount` | number | 稅額 |
| `deductible` | string | `'Y'` / `'N'`，預設 `'Y'` |
| `itemName` | string? | 品名（若 `platformItemNames` 沒提供時使用） |

#### 選項：`opts`

| 欄位 | 說明 |
|---|---|
| `platformInvoiceNos` | `Set<string>` 或 `Array<string>`：電子發票平台已下載的發票號碼集合。若提供，**會排除不在此集合內的列**（紙本／自行輸入／補傳票） |
| `platformItemNames` | `Map<string,string>` 或 `Object`：發票號碼 → 品名。若有則**優先**使用平台明細品名 |
| `company` | 公司資訊 `{taxId, name}` |
| `allowances` | 進項折讓明細陣列（格式見下方「[折讓明細](#3-折讓明細allowances)」）；提供時回傳 `allowance` |
| `skipNonConfirmed` | 預設 `true`：折讓單略過非「折讓已確認」 |

#### 回傳

```js
{
  deductible:    { rows: [...], subtotal: { amount, tax } },
  nondeductible: { rows: [...], subtotal: { amount, tax } },
  allowance: {                                  // 未傳 opts.allowances 時為 null
    deductible:    { rows: [...], subtotal },   // 可扣抵進項之退出及折讓 → 第 40／41 欄
    nondeductible: { rows: [...], subtotal },
    skipped:   [{ allowanceNo, status }],       // 作廢折讓單
    unmatched: ['折讓單號', ...],                // 原發票不在 certs 中、預設可扣抵者
  },
  period: '115年3-4月',
  company: { taxId, name },
  paperCount: 5,        // 排除的紙本/非電子發票筆數
  missingNames: 0,      // 仍查無品名的筆數
}
```

每筆 row：`{ seq, date, invoiceNo, itemName, amount, tax, deductible }`，`seq` 在每個區段內從 1 重編號。

折讓 row 另含 `allowanceNo`、`matched`；`invoiceNo` 為被折讓的**原發票號碼**，扣抵別取原發票在 `certs` 中的扣抵欄（不分是否為紙本），查不到 → `'Y'`、`matched: false`。

---

### 3. 折讓明細（`allowances`）

銷項／進項共用格式，對應電子發票平台下載的折讓明細（**一列一品項**，表頭金額在同一張折讓單的每一列重複）。核心庫會依 `allowanceNo` 彙整成一張張折讓單，金額一律取**表頭**的 `amount`／`tax`（品項金額帶小數，加總會有 1 元差）。

| 欄位 | 平台欄名 | 說明 |
|---|---|---|
| `allowanceNo` | 折讓單號碼 | 同號視為同一張；缺少時每列各自一張 |
| `status` | 折讓單狀態 | 預設略過非「折讓已確認」 |
| `invoiceNo` | 發票號碼 | 被折讓的原發票號碼（用來回查課稅別／扣抵別） |
| `itemName` | 品項名稱 | 進項折讓顯示用 |
| `itemAmount` | 品項折讓金額(不含稅) | 銷項折讓：同一張混有不同課稅別時用來拆分 |
| `itemTax` | 品項折讓稅額 | 銷項折讓：逐品項判斷應稅／免稅（建議提供，前期發票的折讓全靠它判斷） |
| `amount` | 折讓金額(不含稅) | **必填**，表頭金額 |
| `tax` | 折讓稅額 | **必填**，表頭稅額 |
| `allowanceDate` | 折讓單日期 | 進項折讓顯示用 |

---

### 輔助函式

```js
InvoiceCleaner.dateToMingguoYM('0115-03-01');  // → { y: 115, m: 3 }
InvoiceCleaner.inferPeriod([{y:115,m:3},{y:115,m:4}]);  // → '115年3-4月'
InvoiceCleaner.fmtMingguoDate(new Date('2026-03-01'));  // → '115/03/01'
```

---

## 範例：完整整合

```js
// 假設你的發票系統把銷項資料從 DB 撈出來
const dbRows = await db.query(`
  SELECT invoice_no, status, buyer_tax_id, sales_amount, tax_amount,
         total_amount, tax_type, zero_rate_amount, exempt_amount,
         invoice_date, seller_tax_id, seller_name
  FROM invoices_out WHERE period_yyyymm = '11503'
`);

// 轉成 cleanOutputData 期望的格式
const invoices = dbRows.map(r => ({
  invoiceNo:   r.invoice_no,
  status:      r.status,
  buyerTaxId:  r.buyer_tax_id,
  salesAmount: r.sales_amount,
  taxAmount:   r.tax_amount,
  totalAmount: r.total_amount,
  taxType:     r.tax_type,
  zeroRateAmount: r.zero_rate_amount,
  exemptAmount:   r.exempt_amount,
  invoiceDate: r.invoice_date,
  sellerTaxId: r.seller_tax_id,
  sellerName:  r.seller_name,
}));

const result = InvoiceCleaner.cleanOutputData(invoices);

// 渲染到頁面
console.log(result.company.name + ' ' + result.period);
for (const row of result.rows) {
  console.log(row.label, row.count, row.sales, row.tax, row.zero, row.exempt, row.total);
}
console.log('小計', result.subtotal);
```

進項類似：

```js
const certs = wenzhongRows.map(r => ({
  invoiceNo:  r.發票號碼,
  date:       r.日期,
  amount:     r.金額,
  taxAmount:  r.稅額,
  deductible: r.扣抵否,
  itemName:   r.貸方摘要,  // fallback 品名
}));

// 平台已下載的電子發票號碼集合（用於過濾紙本）
const platformSet = new Set(platformRows.map(r => r.發票號碼));

// 平台明細的品名對應
const itemNames = {};
for (const r of platformDetailRows) {
  if (!itemNames[r.發票號碼]) itemNames[r.發票號碼] = r.品名;
}

const result = InvoiceCleaner.cleanInputData(certs, {
  platformInvoiceNos: platformSet,
  platformItemNames:  itemNames,
  company: { taxId: '12345678', name: 'XX股份有限公司' },
});
```

---

## 授權

MIT。可自由整合進商用／非商用系統。
