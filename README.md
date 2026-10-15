# OOS Mobile Auditor V15

V15 connects the Week-to-Date report to the private GitHub data repository using these Streamlit secrets:

- `GITHUB_TOKEN`
- `GITHUB_DATA_REPO`
- `GITHUB_DATA_BRANCH`

Daily audits are stored in `oos_wtd_history.csv` in the private data repository. Re-running the same route date replaces that date instead of double-counting it. The WTD button at the top loads the shared stored Monday-Friday data without requiring pick-sheet uploads.
