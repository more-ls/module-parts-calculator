# Module Parts Calculator

Calculate how many parts you need when building multiple module types. Each module type is defined by its parts and quantities in an Excel file; the calculator multiplies those by how many modules you want to build, compares against inventory, and produces a console summary plus a PDF report.

## Setup

```powershell
cd D:\Codes\module-parts-calculator
pip install -r requirements.txt
```

Use the same Python interpreter when installing and running (for example, if you run with `python`, install with `python -m pip install -r requirements.txt`).

Optional: create a virtual environment to keep dependencies isolated:

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Quick start

1. Open **`part definitions.xlsx`** in Excel and fill in part quantities per module (if not already done).
2. Optionally update **`inventory.xlsx`** with current stock counts (see below).
3. Run the calculator:

```powershell
python module_calculator.py
```

4. Enter your **company / project name**.
5. Enter how many of each module type you want to build.
6. Review the printed summary and the PDF report. The PDF is saved to **`reports/`** and opens automatically in your default viewer.

To run a non-interactive sample with preset module counts:

```powershell
python run_test.py
```

To recreate or update the Excel template from code defaults:

```powershell
python create_part_definitions.py
```

---

## Files

### `part definitions.xlsx`

The main data file you edit in Excel.

| Column | Description |
|---|---|
| **Part** | Part name |
| **Category** | One of: mechanical part, electronics, 3d printing, cnc metal, pcb |
| **Module columns** | Quantity of that part needed **per module** (leave blank if not used) |

This file is read directly by the calculator — no CSV or other files are created.

If this file is deleted, run `python create_part_definitions.py` to recreate it from the defaults in `create_part_definitions.py`.

---

### `module_calculator.py`

The main script you run to get part totals.

**What it does:**
- Reads `part definitions.xlsx` directly in memory
- Loads stock counts from `inventory.xlsx` (treats missing parts as 0 in stock)
- Asks for a company / project name
- Asks how many of each module type you want to build
- Calculates total parts needed across all modules
- Prints a summary and saves a PDF report

**Report sections:**
1. **Total Parts** — all parts needed, with Needed, In Stock, and Remaining to Order, plus category subtotals
2. **Parts to Order** — only items still needed after stock (omitted if nothing to order)
3. **Modules** — module types and build counts

**Output includes:**
- Console tables for all sections above
- PDF report saved to `reports/{project_name}_{YYYY-MM-DD}.pdf`
- PDF opens automatically when the run completes

Part names, categories, and module types (except `MO:` codes) are formatted consistently in the report (title case).

---

### `reports/` folder

Created automatically when you run the calculator. Contains PDF summaries named after your project and the date, for example:

```
reports/My_Project_2026-07-14.pdf
```

---

### `run_test.py`

Non-interactive sample run using preset module counts. Useful for checking that Excel loading, calculations, and PDF generation work without entering values manually.

Edit `SAMPLE_MODULE_COUNTS` and `PROJECT_NAME` at the top of the file to try different scenarios.

---

### `create_part_definitions.py`

Builds or updates the Excel part definitions file from code defaults.

**What it does:**
- Creates `part definitions.xlsx` with formatted table styling
- Defines all module types, parts, categories, and default quantities
- Preserves your existing Excel data when regenerating (quantities and categories you've already filled in)
- Applies rules such as common parts on every module and frog passive vs frog active PCB assignment

**When to run it:**
- First-time setup (no Excel file yet)
- After adding new module types or parts in the code
- To reset the Excel layout while keeping your filled-in values

Edit the constants at the top of this file to change defaults:
- `MODULE_TYPES` — module type names (column headers)
- `PARTS` — part names and categories (rows)
- `COMMON_MODULE_PARTS` — parts every module gets (frog holder, bottom plate, magnets, etc.)
- `MODULE_FROG_PCB` — which modules use frog passive vs frog active
- `BOM_DEFAULTS` — module-specific part quantities

---

### `inventory.xlsx`

Excel file to track **current stock** for every part.

| Column | Description |
|---|---|
| **Part** | Part name (synced from `part definitions.xlsx`) |
| **Category** | Part type (mechanical part, electronics, 3d printing, cnc metal, pcb) |
| **Stock Count** | How many you currently have in stock |

Create or update it with:

```powershell
python create_inventory_excel.py
```

Fill in the **Stock Count** column in Excel. If you regenerate the file, existing stock counts are preserved.

If `inventory.xlsx` is missing, the calculator still runs and treats all stock counts as 0.

---

### `test_module_calculator.py`

Automated tests for calculations, inventory comparison, report sections, and PDF output.

Run after each edit:

```powershell
python test_module_calculator.py
```

---

### `create_inventory_excel.py`

Builds or updates `inventory.xlsx` from the parts list in `part definitions.xlsx`.

---

### `requirements.txt`

Python dependencies:

| Package | Purpose |
|---|---|
| `openpyxl` | Read and write Excel files |
| `fpdf2` | Generate PDF reports |

---

## Module types

| Module | PCB type |
|---|---|
| tip holder 1000 | frog passive |
| tip holder 100 | frog passive |
| MO:HEAT | frog active |
| tube holder 8T2 | frog passive |
| tube holder 888 | frog passive |
| MO:TILT PASSIVE | frog passive |
| MO:TILT | frog active |
| MO:COOL 2ml | frog active |
| MO:COOL 15m | frog active |
| MO:CROSCOPE | frog active |
| MO:SHAKE | frog active |

Every module also includes: frog holder (1), bottom plate (1), Magnets (4), Spring plungers 6 mm (4), magnet stopper (4).

---

## Part categories

| Category | Examples |
|---|---|
| **mechanical part** | Magnets, spring plungers, thermal paste |
| **electronics** | Fans, Peltier, temperature sensor, voltage step down |
| **3d printing** | Frog holder, magnet stopper, tip/tube/cooler bodies |
| **cnc metal** | Bottom plate, top plate, cooling tubes block |
| **pcb** | frog passive, frog active |
