# Bank Reconciliation System in Excel (VBA)

A fully automated bank reconciliation system built with Microsoft Excel and VBA.  
This project replicates real-world accounting processes by enabling structured data entry, transaction logging, reconciliation matching, and mismatch detection — all with the click of a button.

---

## Features

- **Entry Form** to input transactions with validation
- **Automated logging** to History Log
- **Posting to Bank or Cash Books** with a single button
- **Reconciliation Sheet** that auto-matches transactions
- **Mismatch detection** with status and amount variance
- **Conditional formatting** to highlight issues
- **Test data included** to demonstrate functionality

---

## Components

| Sheet Name             | Purpose                                                      |
|------------------------|--------------------------------------------------------------|
| `Entry Form`           | User interface for data entry (Bank or Cash side)            |
| `Bank Statement`       | Stores bank-side transactions                                |
| `Cash Book`            | Stores cash-side transactions                                |
| `Reconciliation Sheet` | Matches and compares Bank vs Cash book entries               |
| `History Log`          | Logs all submitted entries for audit trail                   |

---

##  How to Use

1. **Open the `.xlsm` file in Excel**
2. Enable macros when prompted
3. Navigate to the `Entry Form` sheet
4. Fill in the following fields:
    - **Entry Side** → `Bank` or `Cash`
    - **Transaction ID** → Unique ID (e.g., `TXN001`)
    - **Type** → `Credit` or `Debit`
    - **Date** → e.g., `01-Jun-2025`
    - **Description** → e.g., `Payment from ABC Pvt Ltd`
    - **Amount** → e.g., `18500`
5. Click **Submit Entry** — your transaction is:
    - Logged in the History Log
    - Added to the Bank or Cash Book
    - Sent to the Reconcilliation Sheet

---

##  Reconciliation Sheet Layout

| Transaction ID | Type   | Bank Date | Bank Description | Bank Amount (₹)  |Book Date  | Book Description | Book Amount (₹)  |Status           |Amount Mismatch     |
|----------------|--------|-----------|------------------|------------------|-----------|------------------|------------------|-----------------|--------------------|
| TXN001         | Credit | 01-Jun-25 | ABC Pvt Ltd      | ₹18,500          |           |                  |                  | Not in Cash Book| N/A                |
| TXN002         | Debit  |           |                  |                  | 02-Jun-25 | Office Supplies  | ₹3,250           | Not in Bank     | N/A                |
| TXN003         | Debit  | 03-Jun-25 | Vendor X         | ₹12,000          | 03-Jun-25 | Vendor X         | ₹11,000          | Amount Mismatch | Mismatch by ₹1,000 |
| ...            | ...    | ...       | ...              | ...              | ...       | ...              | ...              | ...             | ...                |

---

## Built With

- **Microsoft Excel VBA**
- **Form controls and Macros**
- **Dynamic VLOOKUP, IFERROR, ABS**
- Conditional Formatting for visual tracking

---

## How to Enable Macros

Since this project uses Excel VBA, macros must be enabled to use the form and reconcilliation features.

###  If Macros Are Blocked (Windows Users):

1. **Close the Excel file**
2. Go to the file in **File Explorer**
3. Right-click the file → Click **Properties**
4. At the bottom, check **☑ Unblock**
5. Click **Apply** → **OK**
6. Reopen the file in Excel

### Then Enable Macros in Excel:

1. Open the file
2. You’ll see a **yellow security warning** at the top
3. Click **"Enable Content"**

---

### Optional (For Persistent Use):

You can adjust your macro settings globally:

1. Go to **File > Options > Trust Center > Trust Center Settings**
2. Under **Macro Settings**, choose:
   - ☑ Enable all macros (for testing only)
   - ☑ Trust access to the VBA project object model
3. (Optional) Under **Protected View**, uncheck:
   - ☑ Enable Protected View for files from the internet
4. Click **OK** and restart Excel

---

>  Tip: You can also move the file to a **Trusted Location** under Trust Center settings if you want Excel to always allow macros from that folder.
