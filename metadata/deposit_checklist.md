# Deposit Checklist

1. Decide release route:
   - `GitHub -> Zenodo` is recommended if you want visible version history.
   - `Zenodo only` is acceptable if you want a simple DOI-bearing archive.
2. Review `README.md`, `CITATION.cff`, and `.zenodo.json`.
3. Confirm that `CC BY 4.0` is acceptable for this public release, or replace `LICENSE` deliberately before release.
4. Mint a persistent DOI.
5. Replace DOI placeholders in `metadata/data_and_code_availability_statement_draft.md`.
6. If using GitHub, create a release tag such as `v1.0.0` before archiving to Zenodo.
7. Keep the public deposit limited to shareable materials; do not upload copyrighted journal PDFs.
8. If using searchRxiv, upload the packet in `searchrxiv/` and record the minted DOI.
9. After the DOI is minted, update the manuscript data/code-availability statement if you decide to submit the repository link with the paper.
