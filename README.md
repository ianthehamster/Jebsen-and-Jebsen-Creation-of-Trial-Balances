# 🌍 Orbitax Trial Balance Generator for Jebsen & Jessen SAP Data

This Excel Script automates the creation of Orbitax-ready Trial Balance files from Jebsen & Jessen’s SAP-exported Balance Sheet and Profit & Loss data. It is designed to streamline **Pillar Two data collection workflows** by transforming raw SAP financial data into structured `.qnr`-compatible trial balances for **GloBE Income Adjustments** and **Covered Taxes** inputs.

---

## 📌 Features

- ✅ **Converts SAP-format BS/PL data to Orbitax trial balance format**
- 💱 **Converts local currency to SGD** using MAS & Wise exchange rates
- 🔁 **Flips P&L signs** for specific account categories (e.g. revenue/expenses)
- 🧾 **Handles missing account codes** by generating synthetic account codes
- 🧠 **Maps J&J Entity Codes to Orbitax Entity Codes** based on master data
- 🧮 **Creates two trial balance sheets**:
  - `TB with Positive Tax Exp`: Tax expenses retained as-is
  - `TB with Negative Tax Exp`: Tax expenses flipped to negative for Pillar Two QNR

---

## 📂 Input Requirements

The workbook must include the following sheets:

| Sheet Name | Purpose |
|------------|---------|
| `MAS and Wise Exchange Rates` | Currency codes in Row 10, exchange rates to SGD in Row 11 |
| `General` | Contains entity mappings between J&J and Orbitax |
| One sheet per legal entity | SAP-exported BS/PL data starting at Row 9 |

---

## 🧠 Logic Summary

### 1. Currency Conversion
Looks up the currency of each entity sheet, and applies the relevant exchange rate to all financial values (converting to SGD).

### 2. Sign Flipping
Flips the signs of accounts beginning with `4`, `5`, `6`, `7` unless the account is a tax expense (e.g. `73000`), in line with Pillar Two income treatment.

### 3. Entity Code Mapping
Maps each entity’s short name (from sheet name) to its Orbitax entity code and legal name using the `General` sheet.

### 4. Tax Expense Handling
- The **Positive** TB is used for standard reporting.
- The **Negative** TB is used to populate **GloBE Income Adjustments**, where tax expenses must be input as negative values.

---

## 📄 Output Sheets

| Sheet Name | Description |
|------------|-------------|
| `TB with Positive Tax Exp` | Orbitax-ready TB with mapped entities and tax expenses intact |
| `TB with Negative Tax Exp` | Identical to above but flips tax expenses to negative |

---

## 📊 Output Format

| Entity Code | Entity Name | Account Code | Account Name | Amount (SGD) |
|-------------|-------------|--------------|--------------|--------------|
| 1050        | JJ Thailand | 40000        | Revenue      | -150000.00   |

---

## 🛠️ How to Run

This script is written in **Office Scripts** for Excel Online.

1. Open your Excel workbook in **Excel for Web** (with Office Scripts enabled).
2. Click **Automate > New Script**.
3. Paste the entire script.
4. Run the script with the workbook containing the required inputs.

---

## 🧾 Notes

- Excludes sheets such as `"Steps"`, `"General"`, `"Orbitax Entity Codes"`, etc.
- Automatically skips rows with no account codes.
- Handles missing account codes by generating random numbers.

---

## 📬 Contact

Created by: **Ian Chow**  
Team: Deloitte Tax Technology Consulting  
Project: Jebsen & Jessen Pillar Two Implementation  
Email: kichow@deloitte.com  
© 2025 Deloitte Singapore  
