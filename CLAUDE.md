# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Purpose

Desktop tool (Tkinter GUI) that imports building surface data from AutoCAD (.xls) and GeoGex (.xlsx) Excel exports, lets the user categorize surface types via drag-and-drop, then generates formatted Excel surface area reports. Targets French real estate professionals (standards: SUB, SU, SHO, SPD, GLA, TAX).

## How to Run

```bash
python Main.py
```

Dependencies (no requirements.txt yet):
```bash
pip install pandas xlsxwriter openpyxl
```

Or via conda (configured in `.vscode/settings.json`):
```bash
conda install pandas xlsxwriter openpyxl
```

## XLS File Structure (AutoCAD export)

Each `.xls` file contains multiple sections (one per surface type: SUB, SU, GLA…). Data rows are identified by a numeric sequence number in column 0. The surface type is read directly from the `Calque` column (e.g. `"SUB Contours"` → `"SUB"`).

`extract_info` filters rows where col 0 is numeric, assigns column names, then extracts `type_su` from the first word of `Calque`. No header parsing needed.

Etage values are strings: `"-1"` (SS), `"0"` (RdC), `"1"` (R+1).

## Architecture

Three modules with clear separation:

- **`Main.py`** — Tkinter GUI (`SurfaceApp` class, 5 tabs : Importer, Glossaire, Générer, Aide, À propos). Owns all UI state including the drag-and-drop glossary editor and the import log console.
- **`process_surf.py`** — All data processing: parsing Excel files (`extract_info`), aggregating by floor/usage (`tab_cd_type`), auto-categorizing via glossary exact-match (`build_affectation_mapping`), pivoting for output (`TCD2Tab`, `Tab_output`), writing the Excel workbook (`export_tables_to_excel`), and generating the printable HTML report (`export_tables_to_html`).
- **`glossary_surf.py`** — Static reference data: three glossaries (`surfaces_utilisables`, `sho`, `spd`) mapping surface type names to categories, plus `denom_surf` (code → glossary key), `real_su_name` (code → French label), `superficie_names` (default super-categories per type), `predefined_cats` (pre-seeded categories per type), and `nota_surf` (regulatory footnotes per type).

## Data Flow

**Import:** `extract_info()` reads Excel + extracts surface type from the `Calque` column → `tab_cd_type()` groups by floor/usage/occupant → `build_affectation_mapping()` auto-assigns categories via exact lookup against the active glossary → displayed in GUI for manual drag-and-drop adjustment.

**Export:** User fills project metadata form → `TCD2Tab()` merges surfaces with final category assignments → `Tab_output()` builds pivot tables with headers/totals/sub-category spans → `export_tables_to_excel()` writes multi-sheet `.xlsx` AND `export_tables_to_html()` writes an A4-landscape printable HTML (same basename, one section per type with tab navigation on screen, hidden on print).

## Key Concepts

- **Surface type codes**: SUB (Brut), SU (Net), SUBL (Louable), SUN (Non-louable), SHO (Hors-Œuvre), SPD (De Plancher), GLA, TAX, TSB. The active glossary depends on which code is being processed.
- **Category mapping**: stored as a DataFrame (`mapping_df`) with columns `[Affectation, cat]`. The Glossaire tab edits this mapping before export. Unclassified rows get `cat = "autres"`.
- **Super-categories** (`super_cats`): per-type ordered dict `{super_cat_name: [categories]}` grouping columns in the output pivot. Editable in the GUI (rename on double-click, drag across columns, right-click to reassign).
- **Frozen-mode paths**: `process_surf._glossary_path()` and `_img_dir()` handle PyInstaller's `sys.frozen` + `sys._MEIPASS` so `glossary_surf.py` stays writable next to the exe and `img/` assets are found at runtime.

## HTML report

`export_tables_to_html(output_tables, infos, html_path)` produces a self-contained `.html` (logos and stamp embedded as base64) with A4-landscape CSS:
- Header grid: logos left · project infos centered · date/dossier right.
- One `<section class="page">` per surface type, separated by `page-break-after: always`.
- Footer: nota intro (uses `infos["mesurage"]` to fill "en xx 2026") + regulatory nota + stamp on the same row.
- Screen-only `<nav class="tabs">` for navigation between types; hidden via `@media print`.
