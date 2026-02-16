# ACC4510 Homework 2B — Deterministic Rank-Order Pipeline

## Research question
**“Rank order the programs or courses based on student ratings or preferences for that year.”**

## What this project does
This repository runs a deterministic pipeline that:
1. Reads **one year** of anonymized UVU MAcc exit survey data from `data/`.
2. Detects ranking/rating fields (best-effort with explicit error messages if detection fails).
3. Computes a course/program rank order using **Net Approval Score (NAS)**.
4. Saves outputs to:
   - `outputs/rank_order.csv`
   - `outputs/rank_order.png`

## NAS method
Primary (forced-rank 1–8) binning logic:
- `1–3` → **Most Beneficial**
- `4–5` → **Neutral**
- `6–8` → **Least Beneficial**

Then for each course/program:
- `pct_most = n_most / n_total * 100`
- `pct_least = n_least / n_total * 100`
- `nas = pct_most - pct_least`

Fallback if scale differs:
- For `1–5` Likert:
  - `4–5` → Most
  - `3` → Neutral
  - `1–2` → Least
  - NAS computed the same way.

If data is already bucketed as Most/Neutral/Least, NAS is computed directly from those categories.

### Exclusions from denominator
Blank values, `NaN`, and values like “did not take” are excluded from `n_total` for that course/program.

## Data file behavior
- Place one or more `.xlsx` files in `data/`.
- Use `DATA_FILE` env var to force a specific file, e.g.:
  - `DATA_FILE="Grad Program Exit Survey Data 2024.xlsx" python src/rank_order.py`
- If `DATA_FILE` is not set, the script uses the first `.xlsx` file in sorted order.

## Local run instructions
```bash
python -m pip install -r requirements.txt
python src/rank_order.py
```

## GitHub Actions automation
Workflow file: `.github/workflows/run.yml`
- Triggers on:
  - `push`
  - `workflow_dispatch`
- Steps:
  1. Set up Python 3.11
  2. Install dependencies
  3. Run `python src/rank_order.py`
  4. Commit updated output files in `outputs/` back to the branch

## Output files
- `outputs/rank_order.csv`: ranking table with `rank`, `course`, counts, percentages, and `nas`.
- `outputs/rank_order.png`: horizontal bar chart sorted by NAS with a zero reference line.
- `outputs/reflection.md`: placeholder template for your reflection write-up.
