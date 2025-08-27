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

## 📂 Project Structure

```plaintext
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

