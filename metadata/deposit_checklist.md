# Deposit Checklist

1. Decide release route:
   - `GitHub -> Zenodo` is recommended if you want visible version history.
   - `Zenodo only` is acceptable if you want a simple DOI-bearing archive.
2. Review `README.md`, `CITATION.cff`, and `.zenodo.json`.
3. Confirm that `CC BY 4.0` is acceptable for this public release, or replace `LICENSE` deliberately before release.
4. Confirm the minted Zenodo DOI (`10.5281/zenodo.19228372`) and keep all public-facing links synchronized to it.
5. Use the appropriate finalized statement in `metadata/data_and_code_availability_statement_draft.md`; only the searchRxiv variant should remain conditional on a future searchRxiv DOI.
6. If using GitHub for a later refresh, create a new release tag deliberately before archiving a new version to Zenodo.
7. Keep the public deposit limited to shareable materials; do not upload copyrighted journal PDFs.
8. If using searchRxiv, upload the packet in `searchrxiv/` and record the DOI or persistent URL only after it is actually minted.
9. If a future searchRxiv DOI is minted, decide whether to append it to the manuscript data/code-availability statement; the current GitHub + Zenodo statement is already submission-ready.
