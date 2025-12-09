# Repository structure review

This repository collects independent utilities for test data processing, Excel exports, and analysis scripts. The layout has been normalized into topic folders with snake/kebab-case names to simplify navigation and dependency management.

## Current layout (high level)
- `automation/`: Auto-VPO automation, PDK reporting helpers, and unified test automation scripts.【F:automation/auto-vpo/test_spark.py†L1-L40】【F:automation/pdk/PDK_weekly_report.py†L1-L40】【F:automation/test-data-auto-analysis/使用说明.md†L1-L24】
- `converters/`: Excel/JSON/pin conversion utilities grouped under consistent names.【F:converters/pin-to-excel/readme.txt†L1-L2】【F:converters/json-to-excel/readme.txt†L1-L16】【F:converters/pin-map-for-leakage/match_pin_groups.py†L1-L36】
- `analysis/`: Test time analysis, PUP model analysis, shop limits extraction, and leakage instance helpers.【F:analysis/test-time-analysis/Test_time_analysis.py†L1-L32】【F:analysis/pup-file-analysis/pup_file_model_rate_analysis.py†L1-L38】【F:analysis/shops-analysis/extract_shops_limits.py†L1-L20】
- `data/`: Input datasets required by specific tools (e.g., leakage conjunction tables).【F:data/leakage-conjunction/Leakage_LIMIT_COLD.xlsx†L1-L1】

## What changed
1. **Normalized naming:** Removed spaces and mixed casing by renaming folders to snake/kebab-case (e.g., `auto-vpo`, `pin-to-excel`).
2. **Topical grouping:** Tools now live under `automation/`, `converters/`, and `analysis/` to reduce clutter at the root and make related utilities discoverable.
3. **Code/data separation:** Required datasets were moved under `data/` (e.g., `data/leakage-conjunction` for merge-table inputs). Add `outputs/` per tool for generated artifacts.

## Next steps
- Add a shared `requirements.txt` for common pandas/Excel dependencies and optional per-tool extras (e.g., `requirements-auto-vpo.txt`).
- Introduce a lightweight `common/` package for repeated helpers (CSV/Excel IO, MTPL parsing) and update scripts to import from it.
- Extend the top-level README with per-tool quick-start links or badges when CI is available.
- Keep generated outputs out of git via `.gitignore` entries (see top-level `.gitignore`).

## Suggested reorganization
- **Normalize names** to kebab- or snake-case without spaces (e.g., `auto-vpo`, `pin-to-excel`). Add `.gitignore` entries to keep generated CSV/XLSX outputs out of version control.
- **Group similar tools** under topical directories, e.g.:
  - `automation/` for `auto-vpo`, `test-data-auto-analysis`, `pdk` helpers.
  - `converters/` for Excel/JSON/pin mapping utilities.
  - `analysis/` for test time, PUP analysis, and leakage pin studies.
  - `data/` for input templates and example datasets (with subfolders per tool).
  - `reports/` or `outputs/` for generated CSV/XLSX results.
- **Add a top-level README** that links to each tool with a one-sentence purpose and quick start (Python version, `pip install -r requirements.txt`, etc.).
- **Extract shared helpers** (e.g., Excel export helpers, MTPL parsing) into a small common module (e.g., `common/`) to avoid duplicating logic across scripts.
- **Document entry points** by adding short `README.md` files inside grouped folders, clarifying required configs (like `spark_automation_config.yaml`) and sample commands.

Adopting the above will make the repository easier to navigate and onboard while keeping data and generated artifacts separate from code.

## Should all tools stay in one repository?
Keeping the utilities together can be effective if they share a language (mostly Python here), have overlapping users, and benefit from a single onboarding guide. A monorepo is helpful for:
- **Shared workflows:** common CI, linting, packaging, and dependency management.
- **Discoverability:** one place to find test/analysis/export helpers.
- **Shared helpers:** easier to factor repeated Excel/CSV/MTPL handling into `common/`.

However, consider splitting a tool into its own repository if it meets most of these criteria:
- Needs an independent release cadence, versioning, or distribution channel.
- Has unique dependencies or runtime environments (e.g., specific Spark/GUI stacks) that bloat the shared environment.
- Has a distinct audience or access policy (e.g., internal-only vs. partner-facing).

If you keep the tools together, organize them as a **structured toolkit**:
1. Create top-level domains (`automation/`, `converters/`, `analysis/`, `data/`, `outputs/`).
2. Standardize names to snake/kebab case and add `README.md` per tool describing inputs/outputs and sample commands.
3. Add a shared `requirements.txt` (plus optional per-tool `requirements-<tool>.txt` for extras) and a `Makefile`/`tox` target to run lint/tests across tools.
4. Extract reusable helpers into `common/` and import them instead of copy/pasting scripts.
5. Gate large data and generated files via `.gitignore`, storing only small sample fixtures under `data/<tool>/`.

This approach preserves the convenience of a single toolkit while keeping code paths clean and reducing coupling between unrelated utilities.
