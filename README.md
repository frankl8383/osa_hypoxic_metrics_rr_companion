# Beyond-AHI OSA Hypoxic Metrics Companion Repository

This repository is a public transparency package for the manuscript:

`Beyond-AHI hypoxic metrics and hard cardiovascular and mortality outcomes in adult OSA-related cohorts: a systematic review and meta-analysis`

It is designed for a `GitHub + Zenodo` release path and is deliberately limited to reproducibility materials that can be shared without redistributing copyrighted full-text PDFs.

Public repository URL:

- `https://github.com/frankl8383/osa_hypoxic_metrics_rr_companion`
- Zenodo DOI: `https://doi.org/10.5281/zenodo.19228372`

## Release Mode

- Preferred: `GitHub` for visibility and version control, then archive a tagged release to `Zenodo` to mint a citable DOI.
- Acceptable: `Zenodo`-only deposit if a public code repository is not desired.
- Optional extra transparency layer: deposit the executed search strings to `searchRxiv` using the ready-to-upload packet in `searchrxiv/`.

## What This Repository Contains

- final protocol and search appendix
- title/abstract screening and full-text review logs
- final extraction master and synthesis backbones
- final risk-of-bias source table
- AF harmonization and cell-freeze support files
- final export and figure-generation script
- final public supplementary CSV files
- final main figure PDFs
- Zenodo and data-availability metadata files
- a submission-ready `searchRxiv` packet for the executed search strategies

## What Is Intentionally Excluded

- copyrighted publisher full-text PDFs
- internal pilot logs and superseded package drafts
- rendered manuscript and cover-letter files
- local environment caches and temporary render outputs

## Final Dataset State Documented Here

- historical executed package: `24` unique articles, `45` cohort-level rows, `33` primary rows, `12` sensitivity/comparator rows
- updated final submission dataset: `31` unique articles, `54` cohort-level rows, `38` primary rows, `16` sensitivity/comparator rows
- primary pooled structure: unchanged `4` cells
- final anchor-centered citation-chasing completeness pass: no further protocol-concordant retained study beyond the `31 / 54` dataset

## Repository Layout

- `code/export_submission_artifacts.py`: final export and figure-generation script
- `data/analysis/`: screening, full-text, extraction, harmonization, and cell-freeze source files
- `data/results/`: final TSV backbones used to build study tables and synthesis outputs
- `release/supplement/`: final public-facing supplementary CSV/MD files
- `release/figures/`: final upload-ready figure PDFs
- `searchrxiv/`: manual-upload packet for the search-strategy deposit
- `metadata/`: Zenodo/GitHub deposit notes, environment notes, and statement options

## Reproduction Notes

This repository is primarily a transparency deposit, not a full standalone workspace clone. The included script is the final export script used in the v37 submission package, but exact regeneration of all submission assets also depends on the broader `osa_meta_20260323` workspace layout and optional desktop tools such as `LibreOffice`, `pandoc`, and `Ghostscript`.

For most public-repository purposes, the key reusable materials are:

- the executed search appendix
- the final screening and full-text review logs
- the final extraction worksheet and source TSV backbones
- the final analysis/export script
- the final upload-ready figures

## Public Release Checklist

- review `metadata/deposit_checklist.md`
- Zenodo DOI minted: `10.5281/zenodo.19228372`
- if desired, submit the search-strategy packet in `searchrxiv/` through the searchRxiv web interface

## Citation Practice

After deposit, cite both:

- the manuscript
- the repository DOI

Repository citation metadata is provided in `CITATION.cff` and `.zenodo.json`.
