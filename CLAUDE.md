# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

---

## Project

SOL Operations — a Google Apps Script project managing customers, events, payments, and invoices through a Google Sheets spreadsheet. Integrates with Sumit accounting software for invoice generation.

---

## Commands

- Deploy: `npm run deploy --msg="commit message"` (clasp push + git commit + push)
- No test framework

---

## Architecture

Generic item pattern: all data entities share a common CRUD layer, with type-specific configuration registered via `addTypeProps()`.

```
src/
  Main.js              Entry point, onOpen/onNewRow handlers
  Menu.js              "SOL" menu definition & handlers
  items/
    Item.js            Generic CRUD (saveItem, getItem, getItems)
    ItemTypesProps.js   Type registry (addTypeProps, getTypeProps)
    ItemUi.js          Dialog/form management
    types/             Per-type config (fields, validation, normalization)
  html/                Server-side templates + client-side form logic
  utils/               Sumit API integration, UI helpers
```

---

## Sheet Structure

- Row 1: Headers (always lowercase)
- Row 2: Template row — new rows copy formatting and formulas from here
- Rows 3+: Data rows

Item types map to sheets: Customers, Events, Payments, Invoices.

---

## Type Definition Pattern

Each type registers a config object:

```javascript
const _typeProps = {
  sheetName: 'SheetName',
  name: 'typeName',
  fields: {
    fieldName: {
      name: 'fieldName',
      columnHeader: 'column header',
      required: true/false,
      showInForm: true/false,
      unique: true/false,
      inputProps: { type, pattern, ... }
    }
  },
  itemDialog: { height, width },
  normalizeFieldValues: fn | undefined,
  validateValues: fn | undefined
};
addTypeProps(_typeProps);
```

---

## Shared Library

Uses `SOLLibrary` (Google Apps Script library, development mode):
- Sheet ops: `getSheet()`, `appendRowsToSheet()`, `sortSheet()`, `getColumnValues()`, `getRowValues()`, `updateRow()`, `addRow()`
- Headers: `getColNumByHeader()`, `getHeaderMap()`
- Data: `generateSheetRows()`, `arrayToObj()`, `toCamelCase()`, `capitalize()`
- UI: `alert()`, `toast()`, `selectRow()`
- HTTP: `post()`
- Logging: `log()`, `logArgs()`

---

## Data Flow

**Save item:**
1. Form submits to `saveItem({typeName, values, rowNum})`
2. Values normalized (type-specific: capitalize names, extract digits)
3. Values validated (type-specific: duplicate checks, required fields)
4. Row appended or updated in sheet
5. Formatting copied from row 2 template

**Create invoices:**
1. Menu → "Create Invoices" (from Payments sheet)
2. Finds unprocessed payment rows (no invoice ID, bank transfer, has bank transaction ID)
3. Calls Sumit API per payment
4. Syncs invoices back from Sumit

---

## External Integration

**Sumit API** — accounting software for invoice/receipt generation.
- Credentials in script properties: `summitCompanyId`, `summitPrivateApiKey`
- Creates InvoiceAndReceipt documents (type 1)
- Currency: THB (764), Language: English (1), Payment: Bank transfer (3)

---

## Conventions

- Private functions: `_functionName()`
- Type config: `_<typeName>Props`
- Constants: `UPPER_CASE` or `_UPPER_CASE`
- Field names in code: camelCase; column headers in sheet: lowercase with spaces
- Validation errors: thrown as JSON in Error message, parsed client-side
- Logging: always include fileName + functionName context

---

## Key Files by Task

| Task | Files |
|------|-------|
| Add a new item type | `src/items/types/NewType.js`, `ItemTypesProps.js` |
| Modify form UI | `src/html/item_form.tmpl.html`, `item_form.js.html` |
| Change menu actions | `src/Menu.js` |
| Sumit API integration | `src/utils/SumitUtils.js` |
| Generic item CRUD | `src/items/Item.js` |
| Event handlers | `src/Main.js` |

---

## Related Projects

- **sol** (`~/dev/src/sol/`) — Flutter mobile app that reads the same Google Sheets as a data source
- **SOLLibrary** — shared GAS utility library (development mode, ID in appsscript.json)
