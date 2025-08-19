# 🌍 Orbitax Trial Balance Generator for Jebsen & Jessen SAP Data

This repository contains Excel Scripts that automate the generation of **Orbitax-ready Trial Balances** for Jebsen & Jessen’s Pillar Two reporting — aligned with GloBE Income and Covered Taxes QNR templates.

---

## ✨ Included Tools

### 1. `Trial Balance from SAP Entity Tabs`
**Creates Trial Balance worksheets** by transforming SAP-exported BS/PL data from each legal entity tab in the workbook.

### 2. `Safe Harbour Consolidated Trial Balance`
**Builds a consolidated trial balance** by extracting Safe Harbour key metrics (e.g., revenue, tax, PBT) from entity-level summary sheets such as `Tax GL`, `PL Info`, and `BS & Others`.

---

## 🧾 Tool 1: Trial Balance from SAP Entity Tabs

### 💡 Features

- 📥 Converts each sheet’s local currency to SGD - Refer to MAS and Wise Exchange Rates .xlsx file for formatting reference
- 🔁 Flips signs for revenue & expense accounts
- 🧠 Maps J&J entities to Orbitax codes using a `General` sheet
- 📊 Creates **2 versions** of the trial balance:
  - `TB with Positive Tax Exp`
  - `TB with Negative Tax Exp` (for GloBE Income QNR)

### 📂 Input Requirements

| Sheet Name | Purpose |
|------------|---------|
| `MAS and Wise Exchange Rates` | Currency codes in Row 10, exchange rates in Row 11 |
| `General` | Entity name mapping to Orbitax |
| One sheet per legal entity | Data starts from Row 9, Column A |

### 📝 Output Format

| Entity Code | Entity Name | Account Code | Account Name | Amount (SGD) |
|-------------|-------------|--------------|--------------|--------------|

---

## 🧾 Tool 2: Safe Harbour Consolidated Trial Balance

### 💡 Features

- 📋 Parses Safe Harbour metrics (Total Revenue, Tax Expense, Net Profit Before Tax, etc.) from 3 summary sheets:
  - `Tax GL`
  - `PL info`
  - `BS & others`
- 🔗 Auto-maps entity codes (e.g., `E99 - JJ Thailand`) to Orbitax codes using `Entity Codes and Names for CbCR` mapping sheet
- ✅ Assigns standardized account codes:
  - `40000` – Total Revenue
  - `72000` – Tax Expense
  - `72001` – Tax Expense (Current)
  - `90000` – Net Profit Before Tax
  - All other metrics: dynamic codes from 1001+

### 📂 Input Requirements

Please attach the `Entity Codes and Names for CbCR` mapping sheet and ensure the sheet is named `Entity Codes and Names for CbCR`.

| Sheet Name | Purpose |
|------------|---------|
| `Tax GL`, `PL info`, `BS & others` | Safe Harbour metric tables |
| `Entity Codes and Names V2` | Maps J&J entity codes to Orbitax |

### 📤 Output Sheet

- `Consolidated Trial Balance` – Clean, Orbitax-ready trial balance for Safe Harbour computations

---

## 🚀 How to Use

1. Open Excel Online with Office Scripts enabled.
2. Click **Automate > New Script**.
3. Paste the relevant script from this repo.
4. Run it on the workbook containing the SAP or Safe Harbour files.
5. Two output sheets will be generated automatically.

---

## 📦 File Naming

- `TB with Positive Tax Exp` → Use for Covered Taxes QNR
- `TB with Negative Tax Exp` → Use for GloBE Income Adjustments QNR
- `Consolidated Trial Balance` → Use for Safe Harbour Testing


