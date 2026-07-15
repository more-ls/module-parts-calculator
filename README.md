# Module Parts Calculator

Calculate how many parts you need when building multiple module types. Each module type is defined by its parts and quantities in an Excel file; the calculator multiplies those by how many modules you want to build, compares against inventory, and produces a PDF report.

## Setup

### Windows (PowerShell)

```powershell
cd path\to\module-parts-calculator
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

In Cursor/VS Code, select **`.venv\Scripts\python.exe`** as the interpreter (Command Palette → **Python: Select Interpreter**).

### macOS / Linux (bash)

```bash
cd path/to/module-parts-calculator
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

In Cursor/VS Code, select **`.venv/bin/python`** as the interpreter.

Use the same Python interpreter when installing and running on any platform.

## Quick start

1. Activate the virtual environment (see Setup above).
2. Run the calculator:

```bash
python module_calculator.py
```

3. If `part definitions.xlsx` or `inventory.xlsx` are missing, they are **created automatically** from the defaults in `create_part_definitions.py`.
4. Enter your **company / project name** and module build counts at the prompts.
5. The PDF report is saved to **`reports/`** and opens automatically. No summary tables are printed to the terminal.

Non-interactive sample run (also PDF only):

```bash
python run_test.py
```

To manually recreate or update the Excel files from code defaults:

```bash
python create_part_definitions.py
python create_inventory_excel.py
```

Run tests after code changes:

```bash
python test_module_calculator.py
```

---

## Project files

| File | In git | Purpose |
|---|---|---|
| `module_calculator.py` | Yes | Main calculator — prompts, PDF output |
| `create_part_definitions.py` | Yes | Code defaults for parts, module types, and BOM |
| `create_inventory_excel.py` | Yes | Builds `inventory.xlsx` from part definitions |
| `run_test.py` | Yes | Non-interactive sample run |
| `test_module_calculator.py` | Yes | Automated tests |
| `part definitions.xlsx` | Yes | BOM data — parts, categories, qty per module |
| `inventory.xlsx` | No | Local stock counts (gitignored) |
| `reports/` | No | Generated PDF reports (gitignored) |
| `.venv/` | No | Python virtual environment (gitignored) |
| `.vscode/settings.json` | Yes | Points the IDE to the project venv (Windows path; use `.venv/bin/python` on macOS/Linux) |

---

## Data files

### `part definitions.xlsx`

The main BOM file. Edit in Excel to change quantities.

| Column | Description |
|---|---|
| **Part** | Part name (e.g. `Bottom Plate`, `Magnets`) |
| **Category** | Part type code: `MECH`, `ELECT`, `3DPRT`, `CNC`, `PCB` |
| **Module columns** | Quantity of that part needed **per module** (blank = not used) |

Created automatically on first run if missing. To regenerate from code defaults:

```bash
python create_part_definitions.py
```

Regenerating preserves existing quantities and categories where possible.

### `inventory.xlsx`

Local stock tracking (not committed to git).

| Column | Description |
|---|---|
| **Part** | Part name (synced from `part definitions.xlsx`) |
| **Category** | Part type code (`MECH`, `ELECT`, `3DPRT`, `CNC`, `PCB`) |
| **Stock Count** | How many you currently have in stock |

Created automatically on first run if missing. To regenerate:

```bash
python create_inventory_excel.py
```

Existing stock counts are preserved when regenerating. Fill in **Stock Count** in Excel for your current inventory.

---

## Scripts

### `module_calculator.py`

**What it does:**
- Auto-creates `part definitions.xlsx` and `inventory.xlsx` if missing
- Reads BOM and inventory from Excel
- Prompts for project name and module build counts (input prompts only — no report printed)
- Calculates total parts needed and remaining to order
- Saves and opens a PDF report

**PDF report sections:**
1. **Modules Summary** — module build counts, then category summary (Needed / In Stock / To Order)
2. **Total Parts** — full parts table
3. **Parts to Order** — only items still needed after stock

**PDF filename:** `reports/{project_name}_{YYYY-MM-DD}.pdf`

If that file is open in a PDF viewer and cannot be overwritten, the next run saves as `{project_name}_{YYYY-MM-DD}_2.pdf`, `_3.pdf`, and so on.

### `create_part_definitions.py`

Source of truth for code defaults. Edit these constants to add or change parts and modules:

| Constant | Contents |
|---|---|
| `PART_CATEGORIES` | Category codes in sort order: `CNC`, `MECH`, `ELECT`, `PCB`, `3DPRT` |
| `MODULE_TYPES` | Module type names (Excel column headers) |
| `PARTS` | Part names and categories |
| `COMMON_MODULE_PARTS` | Parts every module gets |
| `MODULE_FROG_PCB` | Frog Passive vs Frog Active per module |
| `BOM_DEFAULTS` | Module-specific part quantities |

### `create_inventory_excel.py`

Builds or updates `inventory.xlsx` from the parts in `part definitions.xlsx`.

### `run_test.py`

Non-interactive sample run using preset module counts. Produces a PDF only — no terminal output. Edit `PROJECT_NAME` and `SAMPLE_MODULE_COUNTS` at the top of the file.

### `test_module_calculator.py`

Automated tests for calculations, inventory, report sections, and PDF generation.

---

## Dependencies

| Package | Purpose |
|---|---|
| `openpyxl` | Read and write Excel files |
| `fpdf2` | Generate PDF reports |

---

## Module types

| Module | PCB |
|---|---|
| Tip Holder 1000 | Frog Passive |
| Tip Holder 100 | Frog Passive |
| MO:HEAT | Frog Active |
| Tube Holder 8T2 | Frog Passive |
| Tube Holder 888 | Frog Passive |
| MO:TILT PASSIVE | Frog Passive |
| MO:TILT | Frog Active |
| MO:COOL 2ml | Frog Active |
| MO:COOL 15m | Frog Active |
| MO:CROSCOPE | Frog Active |
| MO:SHAKE | Frog Active |

Every module also includes: Frog Holder (1), Bottom Plate (1), Magnets (4), Spring Plungers 6 mm (4), Magnet Stopper (4).

`MO:` module codes are kept as-is in the data and reports. Other module and part names use title case in the PDF.

---

## Part categories

| Code | Description | Examples |
|---|---|---|
| **MECH** | Mechanical parts | Magnets, spring plungers, thermal paste |
| **ELECT** | Electronics | Fans, Peltier, temperature sensor |
| **3DPRT** | 3D-printed parts | Frog holder, magnet stopper, bodies |
| **CNC** | CNC metal parts | Bottom plate, top plate, holders |
| **PCB** | Circuit boards | Frog Passive, Frog Active |

---

## Gitignored files

These are excluded from version control (see `.gitignore`):

- `reports/` — generated PDF reports
- `inventory.xlsx` — local stock data
- `.venv/` — virtual environment
- `__pycache__/` — Python bytecode cache
