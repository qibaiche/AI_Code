# Test and Analysis Toolkit

This repository collects Python utilities for automation, data conversion, and analysis workflows used across test projects. The codebase is now organized as a structured toolkit with consistent naming and clearer separation between code and data.

## Directory layout
- `automation/`: Automation flows and helpers (e.g., auto-vpo, PDK reporting, unified automation scripts).
- `converters/`: Utilities that transform device, JSON, or pin data into Excel or other formats.
- `analysis/`: Analysis and reporting scripts (test time, leakage studies, PUP rate analysis, etc.).
- `data/`: Sample or input datasets required by specific tools. Generated outputs should be kept out of version control.

Each tool keeps its own README (where available) describing inputs/outputs and usage. Shared patterns (e.g., pandas-based CSV/Excel handling) are consistent across tools.

## Quick start
1. Use Python 3.9+.
2. Install dependencies per tool (for example, `pip install -r requirements.txt` if provided) or shared requirements if you add them later.
3. Run a tool from its folder, e.g.:
   ```bash
   python analysis/leakage-conjunction/merge_tables.py
   ```
4. Place required input data in the matching `data/<tool>/` folder (for the merge tables script, see `data/leakage-conjunction`).

## Data and outputs
- Keep input samples under `data/<tool>/`.
- Store generated results under `outputs/<tool>/` (add new folders as needed) and keep large artifacts out of git.

This structure should make it easier to navigate the toolkit, share helpers, and onboard new contributors.
