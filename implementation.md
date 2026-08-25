# Implementation Details

## Architecture

```
Browser (localhost:5000)
    ↓ GET /api/lobs → list available LOBs
    ↓ GET /api/lobs/<lob>/models → list models
    ↓ POST /api/compare → run comparison
    ↓ GET /api/models/<name>/compiled → runtime-compiled SQL
Flask Server
    ↓ reads from data/ directories
Services Layer
    ├── dbt_compiler.py     → expands macros from data/macros/ at runtime
    ├── excel_parser.py     → reads STM columns (auto-detect headers, filters non-data tabs)
    ├── model_matcher.py    → maps SQL files to STM tabs via alias table
    ├── jinja_preprocessor.py → strips Jinja for parser (keeps metadata)
    ├── sql_parser.py       → extracts columns (sqlglot + regex + JSON + macro)
    ├── rule_engine.py      → compares column by column
    ├── comparator.py       → orchestrates pipeline
    └── report_generator.py → writes .xlsx output
```

## LOB Configuration (server/routes/upload.py)

```python
LOB_CONFIG = {
    "bop": {
        "name": "Business Owners Policy",
        "stm_file": "Businessowners Policy Data Specifications.xlsm",
        "sql_folder": "bop",
    },
    "ca": {
        "name": "Commercial Auto",
        "stm_file": "Commercial Auto Data Specifications.xlsx",
        "sql_folder": "ca",
    },
}
```

## API Endpoints

| Method | Endpoint | Purpose |
|--------|----------|---------|
| GET | `/` | Wizard UI |
| GET | `/api/lobs` | List LOBs with file counts |
| GET | `/api/lobs/<lob>/models` | List models for a LOB |
| GET | `/api/models/<name>/details?lob=x` | STM columns for a model |
| GET | `/api/models/<name>/sql?lob=x` | Raw SQL content |
| GET | `/api/models/<name>/compiled?lob=x` | Runtime-compiled SQL (macros expanded) |
| POST | `/api/compare` | Run comparison `{models: [...], lob: "ca"}` |
| GET | `/api/results/download/<name>` | Download .xlsx |
| GET | `/api/results/download-all` | Download all as .zip |

## DBT Compiler (services/dbt_compiler.py)

Runtime macro expansion — reads macro definitions and substitutes parameters:

| Macro | Expansion |
|-------|-----------|
| `{{ m_cleanse('VARCHAR_NOKEY', 'COL') }}` | `(CASE WHEN COL IS NULL OR TRIM(COL)='' THEN 'NOKEY' ELSE TRIM(COL) END)` |
| `{{ m_cleanse('DATE_LOW', 'COL') }}` | `(CASE WHEN COL IS NULL THEN TO_DATE('1900-01-01') ELSE COL END)` |
| `{{ m_cleanse('NUMERIC_ZERO', 'COL') }}` | `(CASE WHEN COL IS NULL THEN 0 ELSE COL END)` |
| `{{ m_int_term(...) }}` | Full 150+ line SQL with CTEs, joins, OBJECT_AGG, cleansing, SCD dedup |
| `{{ source('s', 't') }}` | `s.t` |
| `{{ ref('model') }}` | `model` |
| `{{ config(...) }}` | removed |
| `{% if %}...{% else %}...{% endif %}` | keeps else branch |

## Column Extraction Pipeline (sql_parser.py)

1. Jinja preprocessing → strip + record m_cleanse metadata
2. sqlglot parse → find SELECT with most columns (skip SELECT *)
3. Regex `_cleaned` CTE → paren-depth-aware FROM detection
4. struct_pack/object_construct → extract KEY := expr pairs
5. m_int_term expansion → 18 standard columns from parameters
6. Dedup → prefer CASE > function > direct ref > NULL

## DataType Inference (rule_engine.py)

Only from m_cleansing metadata and expression syntax:

| Source | Inferred Type |
|--------|---------------|
| m_cleanse('VARCHAR_*', col) | VARCHAR |
| m_cleanse('DATE_*', col) | DATE |
| m_cleanse('NUMERIC_ZERO', col) | NUMBER |
| m_cleanse('TIMESTAMP_NTZ_*', col) | TIMESTAMP_NTZ |
| CAST(... AS type) | type |
| 'literal' \|\| 'concat' | VARCHAR |
| CASE WHEN ... THEN 'Y' | VARCHAR |
| SUM/COUNT/AVG(...) | NUMBER |
| coalesce(..., 0) in struct_pack | NUMBER |
| coalesce(nullif(cast(... as varchar)), ' ') | VARCHAR |

## Column Name Matching

1. Strip (PK), (FK) suffixes
2. Replace spaces with underscores
3. Case-insensitive compare
4. Space vs underscore → MISMATCH (STM typo)
5. Cross-check unmatched by expression → NAME MISMATCH

## Summary Calculation

- matched + mismatched = total STM columns (always adds up)
- extra_in_dbt tracked separately
- match_rate = matched / (matched + mismatched)
- status = MISMATCH if column name OR datatype mismatches

## Output (.xlsx)

Sheet 1 "Detailed Comparison": STM Column, DBT Column, Column Name Compare, STM DataType, DBT DataType, DataType Comparison, Column Logic, Transformation, Source Expression, Suggestion

Sheet 2 "Summary": Totals, match rate, transformation breakdown

## STM Tab Filtering (excel_parser.py)

Non-data tabs are automatically skipped when parsing STM workbooks:
- Domains, Table Of Contents, Instructions, Versions
- Template, Sample, Business Rules, Conformed
- Reference, Audit, Coverable Type Master List
- Sheet1, Backward Compatibility, Issues List, Progress Report, CDC

Only tabs with a recognized "Column Name" / "Target Column" header AND not in the skip list are parsed.
Results are cached for 5 minutes to avoid re-parsing on every navigation.

## Model Matching (model_matcher.py)

SQL filenames are mapped to STM tabs using:
1. Entity name extraction: `int_gwpc_bop_polline` → `polline`
2. Direct match: entity name vs tab name (normalized)
3. Alias lookup: `ENTITY_ALIASES` dictionary maps short names to STM tab patterns
4. Substring match: fallback partial matching

BOP aliases include: bldg→Building, polline→Policy Line, condterm→Condition Term,
covgterm→Coverage Term, exclterm→Exclusion Term, modifier→Modifier,
ratefactor→Modifier Rate Factor, premtxn→Premium Transaction, etc.

CA aliases include: all vehicle types (privpas, publictrans, specialtype, truck, zonerated),
all term variants (covgterm_*, condterm_*, exclterm_* for each entity/vehicle),
SI covgterm variants (*covgsi), entity models (dealer, driver, garagesvc, etc.),
scheduled item sub-models (scheditem_*), and union models (cvrbl, covgterm, condterm, exclterm).

## Union Model Resolution (comparator.py)

Union/ref-only models (`select * from {{ ref(...) }} union all ...`) are detected and resolved
to the first available sub-model for:
- Column extraction (comparison)
- Compiled code display

This means comparing `int_gwpc_bop_covgterm` uses columns from `int_gwpc_bop_covgterm_bldg`.

## BOP SQL Models (data/sql/bop/)

| Model | Type | STM Tab | SCD |
|-------|------|---------|-----|
| int_gwpc_bop_polline | Dimension | BOP Policy Line | SCD1+SCD2 |
| int_gwpc_bop_bldg | Dimension | BOP Building | SCD1+SCD2 |
| int_gwpc_bop_sblocation | Dimension | BOP Location | SCD1+SCD2 |
| int_gwpc_bop_covgterm | Dimension | BOP Coverage Term (1) | SCD1+SCD2 |
| int_gwpc_bop_covgterm_bldg | Dimension | BOP Coverage Term (2) | SCD1+SCD2 |
| int_gwpc_bop_covgterm_polline | Dimension | BOP Coverage Term (3) | SCD1+SCD2 |
| int_gwpc_bop_covgterm_sblocation | Dimension | BOP Coverage Term (2) | SCD1+SCD2 |
| int_gwpc_bop_condterm | Dimension | BOP Condition Term (1) | SCD1+SCD2 |
| int_gwpc_bop_condterm_polline | Dimension | BOP Condition Term (1) | SCD1+SCD2 |
| int_gwpc_bop_exclterm | Dimension | BOP Exclusion Term (1) | SCD1+SCD2 |
| int_gwpc_bop_exclterm_polline | Dimension | BOP Exclusion Term (1) | SCD1+SCD2 |
| int_gwpc_bop_modifier | Dimension | BOP Modifier (1) | SCD1+SCD2 |
| int_gwpc_bop_ratefactor | Dimension | BOP Modifier Rate Factor (1) | SCD1+SCD2 |
| int_gwpc_bop_additionalintrst | Dimension | BOP Additional Interest | SCD1+SCD2 |
| int_gwpc_bop_premtxn | Fact/Transaction | BOP Premium Transaction | None (all N/A) |
| int_gwpc_bop_classification_cvrbl | Dimension | (no tab in new STM) | SCD1+SCD2 |
| int_gwpc_bop_jurisdiction_cvrbl | Dimension | (no tab in new STM) | SCD1+SCD2 |

## CA SQL Models (data/sql/ca/) — 71 total

| Category | Count | Models |
|----------|-------|--------|
| Vehicles | 5 | privpas, publictrans, specialtype, truck, zonerated (full _cvrbl models) |
| Coverage Terms | 19 | 10 entity/vehicle terms + 9 SI terms (all use m_int_term macro) |
| Condition Terms | 9 | dealer, garagesvc, juris, namedind, policyline, privpas, publictrans, specialtype, truck |
| Exclusion Terms | 9 | same 9 entities as condition terms |
| Scheduled Items | 14 | 13 sub-models (line_cond/covg/excl, dealer_covg/excl, garagesvc_covg, juris_covg/excl, pp/pt/st/truck/zr_covg) + 1 union |
| Entity Models | 12 | polline, jurisdiction, dealer, driver, garagesvc, namedind, scheditem, linesicond, modifier, ratefactor, additionalintrst, premtxn |
| Union Models | 4 | covgterm, condterm, exclterm, cvrbl |

## SCD Type Inference (sql_parser.py)

SCD types are detected from two sources:

1. **Explicit Jinja variables** (for full models):
   - `{% set key_cols = [...] %}` → N/A
   - `{% set scd2_cols = [...] %}` → 2
   - `{% set scd1_cols = [...] %}` → 1

2. **m_int_term macro structure** (for term models):
   - Key: `{LOB}_{CLAUSE}_KEY`, `END_EFF_DT` → N/A
   - SCD1: `POL_KEY`, `POL_LINE_KEY`, `{LOB}_CVRBL_KEY` → 1
   - SCD2: all other business attributes → 2

## Columns Intentionally Excluded

These columns exist in the STM but are NOT in the SQL models by design:
- **ROW_PROC_DTS** — removed from all BOP and CA models
- **X_SEQ_NO** — removed from all BOP models

The comparison will correctly show these as "NOT FOUND in DBT" — this is expected behavior.
