# STM ↔ DBT Verifier

Automated comparison tool that verifies STM Excel workbooks against DBT/SQL code and generates structured gap analysis reports.

## Quick Start

```bash
cd DBT_STM_Comversion
.venv\Scripts\activate
python server\app.py
```

Open `http://localhost:5000`

## How It Works

1. **Select LOB** — BOP or CA (auto-detects STM files from `data/stm/` folder)
2. **Select Model** — click a model → comparison runs automatically (non-data tabs like Template, Sample, Instructions are filtered out)
3. **View Results** — 4 tabs:
   - **Comparison** — side-by-side table with summary, filters, frozen headers
   - **STM Data** — raw STM columns with source mapping and business rules
   - **SQL Code** — original raw SQL (with Jinja `{{ }}` if present)
   - **Compiled Code** — fully expanded SQL (macros resolved at runtime)

## Key Features

- **Runtime macro compilation** — `m_int_term`, `m_cleansing`, `m_scd2_scd1` expanded to full SQL
- **Union model resolution** — ref-only models (e.g., `select * from ref(...)`) resolve to sub-model for comparison
- **SCD type inference** — detects SCD1/SCD2/Key from `key_cols`/`scd1_cols`/`scd2_cols` AND from `m_int_term` macro structure
- **JSON column extraction** — `struct_pack`/`object_construct` keys parsed as individual columns
- **Frozen headers** — table headers stick when scrolling
- **Smart matching** — normalizes (PK)/(FK) suffixes, spaces vs underscores
- **Cleansing rule detection** — identifies VARCHAR_NOKEY, NUMERIC_ZERO, etc. from expressions
- **5-min STM cache** — parsed Excel results cached for fast "Back to Models" navigation
- **Non-data tab filtering** — Template, Sample, Instructions, Business Rules, etc. automatically excluded

## Comparison Output

Each column shows:
| Field | Description |
|-------|-------------|
| STM Column | Target column from STM |
| DBT Column | Matching column in DBT (or `-`) |
| Column Compare | MATCH / MISMATCH (space vs underscore) / NOT FOUND |
| STM DataType | Type from STM |
| DBT DataType | Inferred from m_cleansing + expression syntax (or `-` if unknown) |
| DataType Compare | MATCH or MISMATCH |
| STM SCD Type | SCD type from STM (N/A, 1, 2) |
| DBT SCD Type | SCD type from SQL (inferred from key_cols/scd1/scd2 or m_int_term) |
| Column Logic | Mapped as per STM / Intentional NULL / Missing |
| Transformation | Direct Pass-through / COALESCE / CASE Logic / Type Cast |
| Source Expression | Full DBT mapping expression |
| Suggestion | No action required / Review needed |

## Project Structure

```
DBT_STM_Comversion/
├── server/                  # Flask web app
│   ├── app.py               # Entry point (port 5000)
│   ├── config.py            # Paths, allowed extensions (.xlsx, .xls, .xlsm)
│   ├── routes/              # API endpoints
│   │   ├── upload.py        # /api/lobs, /api/lobs/<lob>/models
│   │   ├── models.py        # /details, /sql, /compiled endpoints
│   │   └── compare.py       # /api/compare, /download
│   ├── templates/           # HTML (Jinja2 + Tailwind)
│   └── static/js/app.js    # Frontend logic
├── services/                # Core engine
│   ├── dbt_compiler.py      # Runtime macro expansion (m_int_term, m_cleansing, m_scd2_scd1)
│   ├── excel_parser.py      # STM parser (auto-detect headers, skips non-data tabs, 5-min cache)
│   ├── model_matcher.py     # SQL filename → STM tab via alias table
│   ├── jinja_preprocessor.py # Lightweight Jinja strip for SQL parser
│   ├── sql_parser.py        # Column extraction (sqlglot + regex + JSON + macro + SCD inference)
│   ├── rule_engine.py       # Comparison rules, datatype inference, name normalization
│   ├── comparator.py        # Orchestrator + ref-model resolution for union models
│   └── report_generator.py  # .xlsx generation with frozen headers, color cells
├── data/
│   ├── stm/                 # STM workbooks (.xlsx, .xlsm)
│   ├── sql/bop/             # 17 BOP SQL models
│   ├── sql/ca/              # 71 CA SQL models
│   ├── macros/              # 41 DBT macro .sql files
│   └── results/             # Output per run (timestamped folders)
├── git-workflow.html        # Team Git workflow guide
├── requirements.txt
└── .env.example
```

## Supported LOBs

| LOB | STM File | SQL Models | Coverage |
|-----|----------|------------|----------|
| BOP | `Businessowners Policy Data Specifications.xlsm` | 17 models | Policy Line, Building, Location, Coverage Terms (x4), Condition Terms (x2), Exclusion Terms (x2), Modifier, Rate Factor, Premium Transaction, Additional Interest, Classification, Jurisdiction |
| CA | `Commercial Auto Data Specifications.xlsx` | 71 models | 5 Vehicles, 19 Coverage Terms (incl. 9 SI), 9 Condition Terms, 9 Exclusion Terms, 13 Scheduled Items, Policy Line, Jurisdiction, Dealer, Driver, Garage Service, Named Individual, Modifier, Rate Factor, Additional Interest, Premium Transaction, Line SI Condition, plus union models |

## SQL Model Types

| Type | Pattern | SCD | Example |
|------|---------|-----|---------|
| Full dimension | Source CTEs + Business Mapping + Cleansing + m_scd2_scd1 | SCD1 + SCD2 | `int_gwpc_bop_polline.sql` |
| Macro term | `{{ m_int_term(...) }}` call with parameters | Built into macro | `int_gwpc_ca_covgterm_truck.sql` |
| Union model | `select * from {{ ref('...') }} union all ...` | Resolves to sub-model | `int_gwpc_bop_covgterm.sql` |
| Fact/transaction | Full model ending with `SELECT *` (no SCD) | None (all N/A) | `int_gwpc_bop_premtxn.sql` |

## Technology Stack

- Python 3.12 + Flask
- openpyxl (Excel read/write, supports .xlsx/.xls/.xlsm)
- sqlglot (SQL parsing, Snowflake dialect)
- Tailwind CSS via CDN (dark theme)

## Git Workflow

See `git-workflow.html` for the full team branching workflow, or the quick version:

```bash
git checkout main && git pull origin main     # Get latest
git checkout -b feature/your-feature          # Create branch
# ... do work ...
git add -A && git commit -m "message"         # Commit
git push origin feature/your-feature          # Push
# Create Pull Request on GitHub → Review → Merge
```
