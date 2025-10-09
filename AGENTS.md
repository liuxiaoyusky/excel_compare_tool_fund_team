# Repository Guidelines

## Project Structure & Module Organization
The Streamlit front end lives in `app.py`, orchestrating uploads and the Security mapping editor. Core reconciliation logic is in `compare.py` (DataFrame transforms, diff assembly), while `extract.py` houses Excel loaders shared by both code paths. Configuration toggles such as `TOLERANCE_ABS`, `ENABLE_HISTORY`, and file locations sit in `config.py`. Manual mappings are stored in `mapping_override.json`; do not edit `seg_mapping_config.py` except to seed defaults. Generated exports and input snapshots are archived under `history/<timestamp>/`.

## Build, Test, and Development Commands
```
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
python compare.py  # CLI smoke-test using packaged fixtures
```
Activate the virtual environment before running Streamlit. Use the CLI entry point to trace merges and inspect stdout during debugging—both rely on the sample spreadsheets in the repo root.

## Coding Style & Naming Conventions
Stick to Python 3.9+, four-space indentation, and `snake_case` for functions and variables; constants remain UPPER_CASE as in `config.py`. Type hints are encouraged on new helpers (see `compare.py` exports). Format patches with the default `black` profile and keep imports grouped by standard library, third-party, then local modules.

## Testing Guidelines
There is no automated suite yet; prioritize lightweight `pytest` cases around `compare.run_compare_from_sources`, mocking Excel inputs with `pandas` DataFrames. Name files `tests/test_<feature>.py` and mirror fixture names after the Excel sheet being simulated. In addition to unit coverage, exercise the Streamlit flow with the bundled `Spectra.xls` and HSBC workbook before merging.

## Commit & Pull Request Guidelines
Git history mixes uppercase hotfixes and concise imperatives (`add fallback compare`), so keep subject lines short, present tense, and focused on one change. For pull requests, include: purpose summary, notable config adjustments, data samples or screenshots of the `diffs` sheet when UI-affecting, and mention any required updates to `mapping_override.json` or history archives. Confirm you ran Streamlit locally and attached test evidence.
