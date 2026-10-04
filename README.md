# ADOODS Invoice Manager

ADOODS Invoice Manager is a desktop application for creating GST invoices and quotations. It exports Excel workbooks and PDFs, generates UPI payment QR codes for invoices, and tracks invoice receipts and outstanding balances.

## Features

- Create invoices and quotations from separate document series.
- Store reusable seller, buyer, and bank profile details.
- Calculate GST and generate invoice/quotation workbooks and PDFs.
- Record invoice payments by method, review payment requests, and track unpaid balances.
- Search saved quotations and convert a confirmed quotation into an editable invoice.
- Generate UPI QR codes for invoices when a UPI ID is configured.

## Requirements

- Python 3.11 (used by the GitHub Actions Windows build).
- Windows: Microsoft Excel and Windows PowerShell are required for PDF export.
- Linux: LibreOffice (`libreoffice` or `soffice`) is required for PDF export.
- The workbook templates in `data/` must remain available: `C GST.xlsx` and `I GST.xlsx`.

## Setup

### Windows PowerShell

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

### Linux

```bash
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python main.py
```

Install LibreOffice through your Linux distribution's package manager before exporting PDFs.

## Data and Output

- The local SQLite database and signature image are stored in `data/`. In Windows builds, editable XLSX workbooks are stored under `data/workbooks/` beside the executable.
- Active PDFs are stored under `invoices/`. Windows executable builds use `Invoices_PDF_YYYY/Month`, `Private_PDF_YYYY/Month`, or `Quotation_PDF_YYYY/Month` (for example, `Invoices_PDF_2026/October/`); regular runs keep the existing year-only folders. Saving or recording a payment replaces the active PDF at that path.
- Payment-request snapshots and updated request PDFs are kept in a `Requests/` subfolder beside the active monthly PDF. A duplicate paid-invoice PDF is not created there.
- Generated invoice folders are excluded by `.gitignore`.

**Before publishing this repository on GitHub, review `data/invoices.db` and `data/invoice_signature.png`.** They may contain customer, payment, seller, or signature information. Remove sensitive files from version control and Git history before making the repository public; adding them to `.gitignore` alone does not remove files already tracked by Git.

## Build a Windows Application

The repository includes a GitHub Actions workflow at `.github/workflows/build-windows.yml`. It runs on pushes to `main` and can also be started manually from the Actions tab. The built application is uploaded as the `ADInvoice-Windows` artifact.

To build locally in Windows PowerShell:

```powershell
python -m pip install -r requirements.txt
python -m PyInstaller `
  --noconfirm `
  --clean `
  --windowed `
  --onedir `
  --name AdoodsInvoice `
  --icon app.ico `
  --add-data "data/C GST.xlsx;data" `
  --add-data "data/I GST.xlsx;data" `
  main.py
```

The application bundle is created in `dist/AdoodsInvoice/`.

On Linux, PyInstaller uses `:` instead of `;` to separate the source and destination in each `--add-data` argument.
