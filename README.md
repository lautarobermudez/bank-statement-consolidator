# Bank statement consolidator

A monthly script that reads bank statements from several accounts (two Brubank xlsx, one Supervielle xlsx, Mercado Pago xlsx and PDF), keeps **only incoming transfers** for the month, and produces a consolidated Excel report to prepare invoicing.

I wrote it for my work in accounting back office. Company, account and person names in this public copy are placeholders (`ACME`, `UNIT_B`, `JANE DOE`), so the account list in the script must be edited to run it on real data. No statements are included.

> **Built with AI assistance** (Claude Code). I defined the rules (what counts as income, what is "own account") and checked the results against real statements.

## What it does
1. Parses each account's statement (xlsx; PDF with `pdfplumber`).
2. Keeps only received transfers for the chosen month and drops debits, taxes, cheques, rejected items, QR/card payments and yields.
3. Leaves the original xlsx files showing only those rows (autofilter + hidden rows; **nothing is deleted**). For the PDF it writes a sibling xlsx.
4. Writes `Consolidado ingresos bancarios - <Mes> <Año>.xlsx` with a `Resumen` sheet and a `Detalle` sheet, flagging each row as third party or own account.

## Usage
```bash
pip install -r requirements.txt
python consolidar_ingresos.py                 # simulation, writes nothing
python consolidar_ingresos.py --apply         # writes the files
python consolidar_ingresos.py --mes 2026-09 --apply
```
Options: `--carpeta`, `--propio TEXT` (extra sender treated as own account), `--no-filtrar`, `--forzar`.

## Design decisions
- **Simulation by default**; `--apply` is explicit.
- **Non-destructive**: filters instead of deleting rows.
- Refuses to run if a statement is missing or duplicated unless `--forzar` is used.

## Known limitations
- Spanish code and comments; hard-coded list of accounts and statement layouts.
- No tests; one 600-line file.
