# invoice-cleaner-core

進銷項資料清洗的核心邏輯，抽出為**單一 JS 檔、無外部相依**版本，供其他發票系統嵌入使用。

不需要 SheetJS、不操作 DOM；輸入 = 純 JS 陣列／物件，輸出 = 結構化結果物件。

線上完整版（含上傳 xlsx／列印／下載 Excel）：<https://dingcpa.github.io/invoice-cleaner/>

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

把銷項發票陣列依買方統編分為電子二聯／電子三聯，並彙總筆數、銷售額、稅額、合計。

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
| `invoiceDate` | string\|Date? | 發票日期，用於推算申報期別 |
| `sellerTaxId` | string? | 自動推測公司資訊用 |
| `sellerName` | string? | 自動推測公司資訊用 |

#### 選項：`opts`

| 欄位 | 預設 | 說明 |
|---|---|---|
| `skipNonConfirmed` | `true` | 是否略過非「開立已確認」 |
| `company` | – | 直接指定 `{taxId, name}`，覆寫自動推測 |

#### 回傳

```js
{
  rows: [
    { label: '電子二聯', count, sales, tax, total },
    { label: '電子三聯', count, sales, tax, total },
  ],
  subtotal: { count, sales, tax, total },
  period:  '115年3-4月',                  // 自動從 invoiceDate 推算
  company: { taxId: '...', name: '...' },
  skipped: [{ invoiceNo, status }],       // 被略過的非開立確認列
}
```

#### 演算法

- 買方統編去除非數字後若全為 0（含 `''`、`'0'`、`'00000000'`）→ **電子二聯（B2C）**
- 否則 → **電子三聯**
- 電子三聯：直接加總原始 `銷售額` / `稅額` / `合計`
- 電子二聯：先把所有 `總計` 加總，再以 5% 反推
  - 銷售額 = `round(總計合計 / 1.05)`
  - 稅額 = `總計合計 − 銷售額`（保證銷售+稅必等於合計）

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

#### 回傳

```js
{
  deductible:    { rows: [...], subtotal: { amount, tax } },
  nondeductible: { rows: [...], subtotal: { amount, tax } },
  period: '115年3-4月',
  company: { taxId, name },
  paperCount: 5,        // 排除的紙本/非電子發票筆數
  missingNames: 0,      // 仍查無品名的筆數
}
```

每筆 row：`{ seq, date, invoiceNo, itemName, amount, tax, deductible }`，`seq` 在每個區段內從 1 重編號。

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
         total_amount, invoice_date, seller_tax_id, seller_name
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
  invoiceDate: r.invoice_date,
  sellerTaxId: r.seller_tax_id,
  sellerName:  r.seller_name,
}));

const result = InvoiceCleaner.cleanOutputData(invoices);

// 渲染到頁面
console.log(result.company.name + ' ' + result.period);
for (const row of result.rows) {
  console.log(row.label, row.count, row.sales, row.tax, row.total);
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
