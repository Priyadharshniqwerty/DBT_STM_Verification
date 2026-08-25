# STM ↔ DBT Verifier — Claude Context

## How to Run
```bash
cd DBT_STM_Comversion
.venv\Scripts\python.exe server\app.py
# Open http://localhost:5000
```

## Project Layout
- `server/app.py` — Flask entry point (port 5000)
- `server/config.py` — Paths, allowed extensions (.xlsx, .xls, .xlsm)
- `server/templates/base.html` — Base layout (Tailwind + sticky header CSS)
- `server/templates/index.html` — Wizard UI: 3 steps + 4 view tabs (Comparison, STM Data, SQL Code, Compiled Code)
- `server/static/js/app.js` — Frontend (model caching, LOB selection, auto-compare, view tabs)
- `server/routes/upload.py` — LOB listing API (`/api/lobs`, `/api/lobs/<lob>/models`)
- `server/routes/models.py` — Model details + SQL + compiled endpoints (with ref-model resolution)
- `server/routes/compare.py` — Comparison execution + download
- `services/dbt_compiler.py` — Runtime Jinja macro expansion (m_int_term, m_cleansing, m_scd2_scd1)
- `services/excel_parser.py` — STM parser (auto-detects headers, SKIP_TABS filter, 5-min cache)
- `services/model_matcher.py` — SQL filename → STM tab via ENTITY_ALIASES
- `services/jinja_preprocessor.py` — Lightweight Jinja strip for SQL parser (records m_cleanse metadata)
- `services/sql_parser.py` — Column extraction: sqlglot + _cleaned CTE regex + struct_pack JSON + m_int_term + SCD inference
- `services/rule_engine.py` — Comparison: name normalization, datatype from m_cleansing + expression only, source expression
- `services/comparator.py` — Orchestrator (ref-model resolution, matched + mismatched = total STM)
- `services/report_generator.py` — .xlsx with frozen headers, color-coded cells

## Data Organization
```
data/stm/Businessowners Policy Data Specifications.xlsm  # BOP STM (12 data tabs)
data/stm/Commercial Auto Data Specifications.xlsx         # CA STM (54 data tabs)
data/sql/bop/int_gwpc_bop_*.sql                          # 17 BOP SQL models
data/sql/ca/int_gwpc_ca_*.sql                            # 71 CA SQL models
data/macros/*.sql                                         # 41 macro files (used by compiler)
data/results/<timestamp>/                                 # Output per run
```

## LOB Configuration (server/routes/upload.py, compare.py, models.py)
```python
LOB_CONFIG = {
    "bop": {"stm_file": "Businessowners Policy Data Specifications.xlsm", "sql_folder": "bop"},
    "ca": {"stm_file": "Commercial Auto Data Specifications.xlsx", "sql_folder": "ca"},
}
```

## SQL Model Types
- **Full dimension**: Source CTEs + Business Mapping + Cleansing + m_scd2_scd1 (e.g., polline, bldg, modifier)
- **Macro term**: `{{ m_int_term(...) }}` with parameters (e.g., covgterm_truck, condterm_privpas)
- **Union model**: `select * from {{ ref('...') }} union all ...` — resolves to first sub-model for comparison
- **Fact/transaction**: Full model ending with `SELECT *` from _cleaned, no SCD (e.g., premtxn)
- **Scheduled item**: Full model with CVRBL_TYPE_CD = combined CVRBL_TYPE + SI_TYPE

## Key Design Decisions
- DataType inference: ONLY from m_cleansing macro metadata + expression syntax. No column-name guessing. Shows `-` if unknown.
- Column name: space vs underscore = MISMATCH (flags STM typo)
- Source expression: shows full DBT expression as-is
- Missing columns: show `-` for all DBT fields
- Compiled Code tab: uses dbt_compiler.py to expand macros at runtime (not dbt compile)
- Union models: resolve to first sub-model for comparison/compiled view
- Summary: matched + mismatched = total STM columns (always adds up)
- Extra in DBT tracked separately
- Non-data STM tabs filtered via SKIP_TABS in excel_parser.py
- Supported file formats: .xlsx, .xls, .xlsm (config.py ALLOWED_STM_EXTENSIONS)
- Model matching uses ENTITY_ALIASES in model_matcher.py for SQL→STM tab resolution
- SCD types inferred from key_cols/scd1_cols/scd2_cols for full models, and from macro structure for m_int_term models
- ROW_PROC_DTS and X_SEQ_NO intentionally excluded from all SQL models
- Server always starts at Step 1 (no localStorage session restore)
- STM parsing cached for 5 minutes to speed up "Back to Models" navigation
- LOB_CONFIG duplicated in 3 route files (must keep in sync when adding new LOBs)
