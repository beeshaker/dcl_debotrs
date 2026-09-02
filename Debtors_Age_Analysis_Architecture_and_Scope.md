# DCL Debtors Age Analysis — Architecture & Scope

## 1. Purpose

The module converts the monthly consolidated **Debtors List** Excel workbook into a consistent **Debtors Age Analysis** workbook without manual spreadsheet restructuring.

Reference input: `DEBTORS FULL LIST OF ALL SITES 14 AUG 2026.xlsx`  
Reference output: `All_Companies_Master_Age_Analysis_2026-08-14...xlsx`

Platform: **Odoo 19 Community**, hosted on the company VPS.

## 2. V1 Scope

### Included

- Upload one `.xlsx` Debtors List from Odoo.
- Select the reporting date.
- Auto-detect the source worksheet containing the monthly ledger.
- Auto-detect monthly blocks from the spreadsheet headers rather than fixed Excel column letters.
- Identify each debtor using the source Debtor ID; duplicate debtor names are not merged.
- Identify each property/site from the source property column.
- Reconstruct debtor ageing using FIFO/oldest-debt-first allocation.
- Treat Rent/Levy, Recoveries and Adjustments as monthly non-receipt movements.
- Treat negative Receipts as payments/credits and apply them oldest-first.
- Treat negative adjustments as credits and positive adjustments as charges.
- Keep excess credit balances outside positive ageing buckets while preserving the source Current Outstanding.
- Reconcile calculated ageing back to the source Closing Balance/Current Balance.
- Flag differences greater than KES 0.05 as exceptions.
- Generate one Portfolio Summary sheet and one sheet per property/site.
- Generate ageing columns dynamically for every detected month plus Opening.
- Download the generated Excel workbook directly from Odoo.

### Excluded from V1

- Automatic posting of accounting entries.
- Changing `account.move`, invoices or payments.
- Importing debtor balances into Odoo accounting ledgers.
- Automated email distribution.
- Scheduled generation.
- Persistent month-to-month ageing history in Odoo database tables.
- Exact replication of every colour/font/graphic in the supplied manually prepared workbook.

These can be added after the calculation engine is signed off against finance's expected output.

## 3. User Flow

1. Open **Debtors Ageing → Generate Age Analysis**.
2. Upload the monthly consolidated Debtors List `.xlsx`.
3. Enter the reporting date.
4. Optionally enter a worksheet name; otherwise Odoo auto-detects the best ledger sheet.
5. Click **Generate Age Analysis**.
6. Odoo validates the monthly structure and calculates ageing for every debtor.
7. Odoo displays property count, debtor count, exception count and total receivables.
8. Click **Download Excel**.

## 4. Source Workbook Contract

The importer expects:

- Row 1: column labels.
- Row 2: month/date at the start of each monthly block.
- Column A: Debtor ID.
- Column B: Tenant/Debtor Name.
- Column C: Property/Site.
- Each month includes the following labels:
  - `Arrears B/f`
  - `Rent / Levy`
  - `Recoveries`
  - `Adjustments`
  - `Receipts`
  - `Current Bal`
- If the enhanced columns `Opening balance`, `Total Billings`, `Total Receipts`, `Closing Balance` exist, `Closing Balance` is preferred as the reconciliation balance.

The supplied August 2026 source workbook contains an enhanced worksheet with 56 columns and January-August monthly blocks. The importer intentionally detects these headers dynamically so subsequent months can add new blocks without changing Python code.

## 5. Ageing Logic

For each debtor:

1. The earliest detected month's `Arrears B/f` becomes the **Opening** balance.
2. Months are processed chronologically.
3. Positive `Rent / Levy + Recoveries + Adjustments` create debt in that month.
4. Negative non-receipt movements become credits.
5. Negative `Receipts` are treated as payments.
6. Payments and credits clear the oldest positive debt first (FIFO).
7. Any excess/unapplied credit is retained as credit but is not displayed as a negative ageing bucket.
8. The source `Closing Balance` or latest `Current Bal` is retained as Current Outstanding.
9. Calculated net outstanding is compared to the source closing amount. Difference > KES 0.05 creates an exception.

This preserves credit balances such as negative current outstanding without forcing negative numbers into the monthly ageing columns.

## 6. Output Workbook

### Portfolio Summary

- Property / Site Name
- Number of Debtors
- Opening Arrears
- Period Charges
- Period Receipts
- Current Outstanding
- Dynamic monthly ageing buckets, newest to oldest
- Opening
- Exception count

### Property Sheets

- Debtor ID
- Debtor / Tenant
- Billing Frequency (derived indicator)
- Opening Balance
- Period Charges
- Period Receipts
- Current Outstanding
- Dynamic monthly ageing buckets
- Opening
- Reconciliation Exception

Reconciliation exception cells are highlighted so finance can review only the rows that require attention.

## 7. Technical Architecture

```text
Browser / Odoo UI
      |
      v
Transient Wizard: dcl.debtors.age.wizard
      |
      +--> XLSX parser (openpyxl, read-only/data-only)
      |       |
      |       +--> Detect source sheet
      |       +--> Detect month blocks
      |       +--> Read debtor rows
      |
      +--> Pure ageing engine
      |       |
      |       +--> FIFO charge/payment allocation
      |       +--> Credit handling
      |       +--> Closing balance reconciliation
      |
      +--> XLSX renderer (xlsxwriter)
              |
              +--> Portfolio Summary
              +--> One sheet per property
              +--> Downloaded Binary field
```

The ageing engine is isolated from Odoo UI and Excel rendering so it can later be reused by an Odoo list/pivot/dashboard without rewriting the calculation logic.

## 8. Odoo Objects

### Transient model

`dcl.debtors.age.wizard`

Key fields:

- `source_file`
- `source_filename`
- `reporting_date`
- `source_sheet`
- `result_file`
- `result_filename`
- `property_count`
- `debtor_count`
- `exception_count`
- `total_receivables`
- `status_message`

No permanent debtor data is created in V1.

## 9. Security

Access is granted to Odoo's **Accounting / Billing** users via `account.group_account_user`.

The module does not expose a public controller and does not send uploaded financial data outside the VPS.

## 10. Server Dependencies

Python packages required in the same Python environment used by Odoo:

```bash
pip3 install openpyxl xlsxwriter
```

If Odoo runs in a Python virtual environment, activate that environment and install the packages there instead of system Python.

## 11. VPS Deployment

Assuming the custom addons directory is `/opt/odoo/custom_addons`:

```bash
sudo cp -R dcl_debtors_age_analysis /opt/odoo/custom_addons/
sudo chown -R odoo:odoo /opt/odoo/custom_addons/dcl_debtors_age_analysis
```

Confirm the custom addons path exists in `odoo.conf`, for example:

```ini
addons_path = /opt/odoo/odoo/addons,/opt/odoo/custom_addons
```

Install Python dependencies, restart Odoo and update the module list:

```bash
sudo systemctl restart odoo
```

Then in Odoo:

1. Enable Developer Mode.
2. Apps → Update Apps List.
3. Search **DCL Debtors Age Analysis**.
4. Install.

For command-line installation/upgrades, adapt the database name and paths to the VPS:

```bash
/opt/odoo/odoo-bin -c /etc/odoo.conf -d YOUR_DATABASE -i dcl_debtors_age_analysis --stop-after-init
/opt/odoo/odoo-bin -c /etc/odoo.conf -d YOUR_DATABASE -u dcl_debtors_age_analysis --stop-after-init
```

Restart the normal Odoo service after command-line installation.

## 12. Acceptance Tests

Finance should validate at least:

- Portfolio property count equals the source.
- Debtor count by property equals the source.
- Current Outstanding by property equals source closing balances.
- Several straightforward monthly debtors.
- A debtor with an opening balance.
- A debtor with receipts clearing multiple months.
- A debtor with a positive adjustment.
- A debtor with a negative adjustment.
- A debtor with a credit/negative closing balance.
- Duplicate debtor names with different Debtor IDs remain separate.
- All exceptions are explainable and reviewed before sign-off.

## 13. Recommended Phase 2

After Finance signs off V1 calculations, add a persistent `dcl.debtors.age.snapshot` model so users can retain monthly snapshots, compare movements, build Odoo pivot/graph views, track collections by property, and optionally email approved reports.
