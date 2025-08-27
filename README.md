# Excel Bank Reconciliation & Automation Tool

A smart, macro-enabled Excel solution for automating bank reconciliation tasks. It cleans raw transaction data, matches entries, validates balances, highlights discrepancies, and links cheque records with ledger entries — built for accuracy, speed, and ease of use.

---

##  Features

- ✅ **Power Query Integration** for cleaning raw bank data
- ✅ **Automated Transaction Import** with duplicate checks
- ✅ **Balance Validation** using Expected vs Actual SALDO
- ✅ **Cheque Linking** between EDC and Cheques sheets
- ✅ **Search Utilities** for quick transaction lookup
- ✅ **Conditional Formatting** to highlight discrepancies
- ✅ **Cleaned Data Reset** for fresh imports

---

##  Project Structure

````
ExcelBankReconciliation/
│
├── VBA/
│   ├── RefreshCleanData.bas
│   ├── AppendNewTransactions.bas
│   ├── BalanceValidation.bas
│   ├── LinkEDCtoCheque.bas
│   ├── SearchEDC.bas
│   ├── SearchCHEQUES.bas
│   └── ClearCleanData.bas
├── Documentation/
│   └── SaldoValidationGuide.md
├── Screenshots/
│   └── dashboard_preview.png
└── README.md

````
---

##  How It Works

### 1.  Clean Raw Data with Power Query
- Import raw bank statements into the `Raw_Data` sheet.
- Use Power Query to remove noise, format columns, and load cleaned data into `Clean_Data`.
- Refresh the query with a single click using the `RefreshCleanData` macro.

### 2.  Append New Transactions
- The `AppendNewTransactions` macro compares transaction dates and appends only **new, unique entries** to the `EDC` sheet.
- Transactions are automatically sorted from **newest to oldest**.

### 3.  Validate Balance
- The `BalanceValidation` macro calculates **expected balances** based on transaction type (`D` for debit, `C` for credit).
- Discrepancies are flagged using conditional formatting (`True`/`False`).

### 4. 🔗 Link Cheques
- The `LinkEDCtoCheque` macro matches **Cheque Numbers** between:
  - `EDC` sheet → Column **L**
  - `Cheques` sheet → Column **B**
- Transfers related metadata (e.g., Cheque Date) into the `EDC` sheet.

### 5.  Search Transactions
- `SearchEDC` and `SearchCHEQUES` macros highlight matching entries based on user input:
  - `EDC!D4` for EDC search
  - `CHEQUES!D5` for Cheques search

### 6.  Clear Cleaned Data
- The `ClearCleanData` macro resets the `Clean_Data` sheet to prepare for a fresh import cycle.

---

## 🧩 Requirements

- Microsoft Excel 2016 or later  
- Macros must be enabled  
- Developer tab activated in Excel

---

## 📸 Screenshots

> ![Dashboard Preview](<img width="1655" height="972" alt="Screenshot 2025-08-27 130114" src="https://github.com/user-attachments/assets/2da15f16-d55c-4534-a8a1-254698dc75ae" />
)

---

## 📜 License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

## 🙌 Acknowledgements

Built by **Princess Nkemjika .O** in **Lagos, Nigeria** 🇳🇬  
Inspired by real-world financial reconciliation needs and designed for speed, accuracy, and simplicity.




