# File Locations Reference

## HTML / UI Files
| File | Path | Purpose |
|------|------|---------|
| Base template | `server/templates/base.html` | Layout, Tailwind config, sticky header CSS |
| Main wizard | `server/templates/index.html` | 3 steps + 4 view tabs |
| JavaScript | `server/static/js/app.js` | Wizard state, model caching, auto-compare, view switching |

## Server / API Files
| File | Path | Purpose |
|------|------|---------|
| App entry | `server/app.py` | Flask app factory |
| Config | `server/config.py` | Paths, limits, allowed extensions (.xlsx, .xls, .xlsm) |
| Upload/LOB API | `server/routes/upload.py` | /api/lobs, /api/lobs/<lob>/models |
| Models API | `server/routes/models.py` | /details, /sql, /compiled endpoints + ref-model resolution |
| Compare API | `server/routes/compare.py` | /api/compare, /download |
| Pages | `server/routes/pages.py` | GET / serves UI |

## Services (Core Logic)
| File | Path | Purpose |
|------|------|---------|
| DBT Compiler | `services/dbt_compiler.py` | Runtime macro expansion (m_int_term, m_cleansing, m_scd2_scd1) |
| Excel Parser | `services/excel_parser.py` | STM reader (auto-detect headers, SKIP_TABS filter, 5-min cache) |
| Model Matcher | `services/model_matcher.py` | SQL filename → STM tab via ENTITY_ALIASES |
| Jinja Preprocessor | `services/jinja_preprocessor.py` | Strips Jinja for parser, records metadata |
| SQL Parser | `services/sql_parser.py` | Column extraction (sqlglot + regex + JSON + macro + SCD inference) |
| Rule Engine | `services/rule_engine.py` | Comparison, datatype inference, name normalization |
| Comparator | `services/comparator.py` | Orchestrator, ref-model resolution, summary builder |
| Report Generator | `services/report_generator.py` | .xlsx with frozen headers, color cells |

## Data Directories
| Directory | Contents |
|-----------|----------|
| `data/stm/` | Businessowners Policy Data Specifications.xlsm (BOP), Commercial Auto Data Specifications.xlsx (CA) |
| `data/sql/bop/` | 17 BOP SQL models (full models + union models) |
| `data/sql/ca/` | 71 CA SQL models (vehicles, terms, entities, scheduled items, union models) |
| `data/macros/` | 41 macro .sql files (m_cleansing, m_int_term, m_scd2_scd1, etc.) |
| `data/results/` | Output: timestamped folders with .xlsx + run_log.txt |

## Documentation
| File | Purpose |
|------|---------|
| `README.md` | Quick start + feature overview + project structure |
| `run_steps.txt` | Step-by-step instructions for running and extending |
| `implementation.md` | Technical details, API, compiler, inference rules |
| `CLAUDE.md` | Context for Claude Code sessions |
| `file_locations.md` | This file |
| `git-workflow.html` | Team Git workflow guide (open in browser) |
| `requirements.txt` | Python dependencies |
| `.env.example` | Environment variable template |

## LOB Configuration
LOB_CONFIG is defined in 3 files (must be kept in sync):
- `server/routes/upload.py`
- `server/routes/compare.py`
- `server/routes/models.py`

```python
LOB_CONFIG = {
    "bop": {"stm_file": "Businessowners Policy Data Specifications.xlsm", "sql_folder": "bop"},
    "ca": {"stm_file": "Commercial Auto Data Specifications.xlsx", "sql_folder": "ca"},
}
```
