# Leakage conjunction merge helper

This helper merges `SIO_BSCAN_PCD_4JMP.xlsx` with `Leakage_LIMIT_COLD.xlsx` using the shared `Test_Type` and `Configuration` columns.

## Inputs
- Place both Excel files under `data/leakage-conjunction/`.

## Usage
```bash
python analysis/leakage-conjunction/merge_tables.py
```

The script prints column matches and writes the merged table to `conjunction_merge_<timestamp>.xlsx` in the current working directory.
