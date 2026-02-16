# Composite Reporter

Automates filling `Composite Report Spreadsheet.xlsx` using QBO Profit & Loss and Balance Sheet exports, with client-specific mapping intelligence and tie-out validation.

## What It Produces
Running the pipeline writes these files to `out/`:
- `coach_filled.xlsx`
- `UNMAPPED_PL_ACCOUNTS.xlsx`
- `UNMAPPED_BS_ACCOUNTS.xlsx`
- `TIEOUT.json`
- `run.log`

## Client Profiles (New Default Flow)
Create client profiles in `clients/<client_id>/profile.json` and connect each client to its own template + mapping files.

Example:

```json
{
  "client_id": "jones-auto",
  "display_name": "Jones Auto Service",
  "template_path": "Composite Report Spreadsheet.xlsx",
  "mapping_pl_path": "mapping_pl.csv",
  "mapping_bs_path": "mapping_bs.csv",
  "confidence_threshold": 0.87,
  "tolerance": 1.0,
  "learned_confidence_threshold": 0.96
}
```

Paths in `profile.json` are relative to that client folder unless absolute.

## Inputs Expected Per Run
- one QBO Profit & Loss `.xlsx`
- one QBO Balance Sheet `.xlsx`
- selected client profile

Mapping CSV columns must be:
- `qbo_account_name`
- `template_label`

## One Command (Hands-Off)
Run everything (tests + parse + map + fill + outputs + tie-outs):

```powershell
python run.py --client jones-auto
```

If tests already passed and you only want pipeline:

```powershell
python run.py --client jones-auto --skip-tests
```

## Web Page Runner (Browser UI)
Launch a local webpage:

```powershell
python run_web.py
```

Open `http://127.0.0.1:8000`, pick the client, upload only P&L + Balance Sheet, and click `Run Composite Reporter`.
The page returns run status and direct download links for all output files.
Template + mappings are pulled from the selected client profile.

## Public Demo Link (One Command)
If you want other people to use the app immediately without deploying hosting, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\start_demo.ps1
```

This starts:
- the web app locally
- a Cloudflare temporary public URL (trycloudflare)

It prints:
- `public_url` (share this link)
- `web_pid` and `tunnel_pid`

To stop the demo:

```powershell
powershell -ExecutionPolicy Bypass -File .\stop_demo.ps1
```

Note: your computer must stay on and connected while others use the demo link.

## Feedback Form (Web)
The home page includes **Submit Mapping Feedback**.
- Saves feedback to a global inbox: `feedback_inbox/feedback_inbox.csv`
- Saves feedback to client log: `clients/<client_id>/feedback_log.csv`
- Optional auto-apply writes mapping rows immediately to:
  - client mapping file (`mapping_pl.csv` or `mapping_bs.csv`)
  - global feedback mapping library in `clients/_generated_reference/`

## Deploy Public URL (Render)
This repo is now configured for Render (`render.yaml` + `Procfile`).

Steps:
1. Push this repo to GitHub.
2. In Render, click `New` -> `Blueprint`.
3. Connect your GitHub repo and deploy.
4. Render builds with `pip install -r requirements.txt` and starts with:
   `uvicorn webapp:app --host 0.0.0.0 --port $PORT`

Notes:
- `TEMPLATE_PATH` is pre-set in `render.yaml` to `clients/sample-auto-repair/Composite Report Spreadsheet.xlsx`.
- Render filesystem is ephemeral, so uploaded run artifacts are temporary unless you add persistent storage.

## New Client Onboarding (COA Upload)
From the same web page, use **Onboard New Client**:
- enter `Client ID` and `Client Display Name`
- upload client COA (`.xlsx` or `.csv`)
- keep template path pointed to your master coaching template

On submit, the app:
- creates `clients/<client_id>/profile.json`
- stores the uploaded COA in that client folder
- seeds mapping files from your proven auto-repair mapping base
- optionally ingests extra mapping CSVs from a training/reference folder (any `*mapping*pl*.csv` and `*mapping*bs*.csv` under that folder)
- if mapping CSVs are not present, it can auto-generate reference mappings from training bundles that include `Profit+and+Loss`, `Balance+Sheet`, and completed `Composite Report Spreadsheet`
- auto-infers additional mappings from the COA

The new client appears in the run dropdown immediately.

Tip for higher first-run accuracy:
- Point **Training Reference Folder** to a directory containing completed client mapping files so onboarding can borrow proven mappings before first run.

## CLI Entrypoint

```powershell
python -m composite_reporter --client jones-auto --pl "P&L.xlsx" --bs "BS.xlsx" --outdir "out"
```

If needed in a fresh environment, install the package first:

```powershell
python -m pip install -e .
```

## Install (Non-Technical)
1. Install Python 3.11+.
2. Open terminal in this project folder.
3. Create a virtual environment:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

4. Install dependencies:

```powershell
pip install -r requirements.txt
```

5. Run:

```powershell
python run.py --client jones-auto
```

## How Mapping Works
- Primary source: each client's `mapping_pl.csv` and `mapping_bs.csv`.
- If no usable mapping exists, suggestions are generated by fuzzy + normalized semantic matching.
- Auto-repair domain rules (parts/labor/sublet/tires/shop supplies/payroll tax/cash/AP/AR/equity, etc.) are applied before generic fuzzy logic.
- Leading account numbers are normalized out so mapping survives COA renumbering better.
- High-confidence fuzzy matches are auto-added to `mapping_pl.learned.csv` and `mapping_bs.learned.csv` per client.
- Confidence threshold default is `0.87`.
- Any match below threshold is **not auto-filled** and goes to `UNMAPPED_*` output.

## Derived Totals Fill
- Totals present in QBO exports (for example `Total Equity`, `Total Assets`, `Total Liabilities`, `Net Income`) are now mapped directly to matching template labels when present.
- This allows key summary rows to populate even when line-level mapping is incomplete.

## How Template Filling Works
- Reads `New Composite Worksheet`.
- Looks for labels in column `C` and column `F`.
- Writes numeric amounts into `D` (for `C`) and `G` (for `F`).
- Duplicate labels are treated as ambiguous and skipped; warning is logged in `TIEOUT.json`.

## Tie-Outs
`TIEOUT.json` includes:
- P&L mapped/unmapped/ignored sums and deltas vs reported totals found in export.
- Balance Sheet mapped/unmapped/ignored sums and deltas.
- Assets = Liabilities + Equity check (within tolerance, default `1.00`).
- Inputs, parameters, warnings, errors, timestamp.
- `status` becomes `FAILED` if critical tie-out fails.

## Review Unmapped + Update Mappings
1. Open `UNMAPPED_PL_ACCOUNTS.xlsx` and `UNMAPPED_BS_ACCOUNTS.xlsx`.
2. Review suggested labels and confidences.
3. Add correct rows to client `mapping_pl.csv` / `mapping_bs.csv`.
4. Re-run `python run.py --client <client_id>`.

## Operator Notes
Typical failure modes:
- Missing one of the required filenames.
- Mapping CSV has wrong column names.
- Template sheet name is not exactly `New Composite Worksheet`.
- Duplicate template labels where destination row is ambiguous.

Adjust threshold/tolerance:
- `--confidence-threshold 0.90` (stricter mapping)
- `--tolerance 0.50` (stricter tie-out)

Add new mappings safely:
- Append rows to mapping CSVs only after reviewing unmapped suggestions.
- Keep `qbo_account_name` exactly as shown in QBO export when possible.

## Assumptions Implemented
- QBO exports may include preamble/header/subtotal lines and indentation; parser auto-detects account + amount columns and ignores total/subtotal lines for filling.
- If multiple amount columns exist, the right-most numeric column is used as the active period.
- Section lines without numeric amounts are ignored for fills but used as context when available.

## Test Notes
- Unit tests cover normalization, numeric parsing, confidence threshold behavior, and template writes.
- Integration test generates synthetic workbooks and validates full output set and tie-out presence.
- In restricted offline environments, a local `pytest` shim is included so `python -m pytest -q` still runs.
