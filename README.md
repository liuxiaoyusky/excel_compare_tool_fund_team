# NAV Validation Platform

Automated NAV reconciliation between inbound provider files and HSBC authoritative data. Upload the excel workbooks through Streamlit or call the CLI to trigger the zero-tolerance validator. All logic now follows a domain-driven design (DDD) layout with type-safe entities and application services backed by testable domain services.

👉 New contributor? Read [`AGENTS.md`](AGENTS.md) first.

## Quick Start

```bash
# 1. Set up environment
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 2a. Run Streamlit UI
streamlit run app.py

# 2b. Or run the CLI
python -m nav_checker.cli spectra.xls "HSBC Position Appraisal Report (EXCEL).xlsx"
```

The CLI prints a summary; the Streamlit app offers mapping management, tables, and downloadable diff files.

## Tests (TDD)

Unit tests live under `tests/` and focus on the domain validator and mapping storage. Run them with:

```bash
python -m pytest
```

(`pytest` must be installed inside the virtualenv.)

## Mapping Management

The Streamlit “Security Mapping Editor” manipulates `mapping_override.json`. Keys/values are normalized to uppercase and merged with the seeded defaults from `seg_mapping_config.py`. You can also edit the JSON directly for bulk changes.

## New Architecture at a Glance

```
nav_checker/
  domain/                 # Entities, repositories, validation service
  application/            # Use cases orchestrating repositories + validator
  infrastructure/
    parsing/              # Excel parsers for Spectra and HSBC
    repositories/         # Excel-backed repository implementations
    storage/              # Mapping persistence utilities
  presentation/           # CSV/HTML diff renderers
  cli.py                  # CLI entrypoint for workflows & n8n hooks
app.py                    # Streamlit UI built on nav_checker APIs
compare.py                # Backwards-compatible facade (legacy callers)
```

Key characteristics:
- **Domain-Driven Design**: `NavIdentity`, `NavRecord`, and `NavValidator` isolate business rules.
- **Future orchestration**: n8n can call the CLI or import the package for Stage 2+ automation.
- **File lineage**: every `NavRecord` carries the file hash and parsing lineage for auditing.

## Roadmap Alignment

The staged plan from the task statement maps directly onto the new modules:

1. **Schema foundation** – captured via `nav_checker.domain.models`.
2. **Ingestion** – Excel repositories ready for n8n file or DB triggers.
3. **Normalization** – parsing utilities enforce currency/date typing with Decimal math.
4. **Authoritative fetchers** – HSBC repository implemented; add DB repository next.
5. **Validation & diffs** – `NavValidator` + presentation renderers.
6. **Orchestration & alerts** – CLI output is n8n-ready today.
7. **Human review** – Streamlit downloads CSV/HTML for manual sign-off.

Future work: add DB repositories, audit-store writers, and Slack/email alert adapters under `infrastructure/`.
