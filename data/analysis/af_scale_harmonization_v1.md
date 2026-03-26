# AF Scale Harmonization v1

Date: `2026-03-23`

## Goal

Create an exploratory `T90 -> incident atrial fibrillation` sensitivity meta-analysis by harmonizing the French `per 1%` T90 row onto the Cleveland cohort's reported `per 10-unit` scale.

## Rows Used

- Cleveland reported row:
  - `Heinzinger_2023_Cleveland_incAF_T90_per10`
  - reported directly as `HR 1.06 (1.04-1.08)` per 10-percentage-point increase in `T90`
- Briancon-Marjollet reported row:
  - `BrianconMarjollet_2021_multicenterOSA_incAF_T90_per1pct`
  - reported as `HR 1.01 (1.00-1.02)` per 1% increase in `T90`

## Harmonization Rule

- primary exploratory method:
  - assume log-linear hazards across `T90`
  - convert `HR_1%` to `HR_10%` via `HR_10 = HR_1^10`
  - propagate reported confidence limits the same way on the log scale
- precision-check method:
  - keep the same `HR_10 = HR_1^10`
  - derive `SE` from the reported `p` value for the continuous `T90` coefficient
  - this avoids over-widening caused by the rounded lower bound of `1.00`

## Exploratory AF Sensitivity Results

- primary harmonized AF sensitivity:
  - pooled random-effects `HR 1.0615 (1.0420-1.0814)`
  - `I²=0.00%`
- precision-check AF sensitivity:
  - pooled random-effects `HR 1.0637 (1.0401-1.0878)`
  - `I²=5.85%`

## Interpretation

- both harmonization variants point in the same direction
- the pooled AF signal is numerically stable because the transformed Briancon-Marjollet row is directionally concordant with the Cleveland row
- this remains an exploratory sensitivity analysis rather than a default primary pooled cell because one contributing study was mathematically rescaled from rounded published summary statistics

## Non-Pooled Categorical Support

- `BrianconMarjollet_2021_multicenterOSA_incAF_T90_q4`:
  - `HR 2.20 (1.23-3.91)`
  - highest versus lowest `T90` quartile
  - retained as supportive narrative evidence only
