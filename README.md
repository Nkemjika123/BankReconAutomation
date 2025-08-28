# Excel Bank Reconciliation & Automation Tool

A smart, macro-enabled Excel solution for automating bank reconciliation tasks. It cleans raw transaction data, matches entries, validates balances, highlights discrepancies, and links cheque records with ledger entries — built for accuracy, speed, and ease of use.

---
##  Table of Contents

- [ Data Source](#Data-Source)
- [ Project Overview](#project-overview)
- [ Project Structure](#project-structure)
- [ How It Works](#how-it-works)
- [ Requirements](#requirements)
- [ Download Resources](#download-resources)
- [ Screenshots](#screenshots)
- [ License](#license)
- [ Acknowledgements](#acknowledgements)
- [ Support](#support)

---
## Data Source
This dataset was manually curated for demonstration and testing purposes. It does not contain any real customer or financial institution data.

---
##   Project Overview

- ✅ **Power Query Integration** for cleaning raw bank data
- ✅ **Automated Transaction Import** with duplicate checks
- ✅ **Balance Validation** using Expected vs Actual Balance
- ✅ **Cheque Linking** between EDC and Cheques sheets
- ✅ **Search Utilities** for quick transaction lookup
- ✅ **Conditional Formatting** to highlight discrepancies
- ✅ **Cleaned Data Reset** for fresh imports

---

##  Project Structure

````
BankReconAutomation/
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
│   └── BalanceValidationGuide.md
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
- The `ClearCleanData` macro resets the `Cleaned_Data` sheet to prepare for a fresh import cycle.

---

##  Requirements

- Microsoft Excel 2016 or later  
- Macros must be enabled  
- Developer tab activated in Excel
- Power Query installed and configured

---
## Download Resources

[View VBA Code (PDF)](https://github.com/Nkemjika123/BankReconAutomation/blob/main/Personal%20Financial%20Reconciliation%20VBA%20CODES.pdf)

[Download Manual]([docs/BankReconUserManual.pdf](https://github.com/Nkemjika123/BankReconAutomation/blob/main/%F0%9F%93%98%20User%20Manual_Bank_Reconciliation%20.pdf))

[Download Sample Dataset]([Data/sample_transactions.xlsx](https://github.com/Nkemjika123/BankReconAutomation/blob/main/Personal_Project_Bank_Management.xlsx))

---
##  Screenshots

<img src="https://github.com/user-attachments/assets/2da15f16-d55c-4534-a8a1-254698dc75ae" width="600" alt="Dashboard Preview">


---

##  License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

##  Acknowledgements

Built by **Princess Nkemjika Onwubuche** in **Lagos, Nigeria** 🇳🇬  
Inspired by real-world financial reconciliation needs and designed for speed, accuracy, and simplicity.

---

## 📞Support

For questions or troubleshooting, contact:
Nkemjika 
Email: analystnkem@gmail.com
GitHub: https://github.com/Nkemjika123





