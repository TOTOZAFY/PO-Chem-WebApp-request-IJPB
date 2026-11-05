# PO-Chem WebApp – FAIR Data Treatment

Version 2.4 adds the FAIR Data treatment entry workflow while preserving the existing request, Sample, Analysis, user-profile, language and Nextcloud features.

## Data treatment
- Web form organized by the sections already present in `FAIR_template.xlsx`.
- Full editable POS/NEG table for reviewing and editing all parameter values.
- Separate save actions for the web form and full table.
- Values are written back to the matching cells of the existing `Data treatment` worksheet; the workbook structure and labels are preserved.
- Requester access uses the request reference and requester email. Administrator access uses the existing `ADMIN_PASSWORD` secret.
- Data treatment interface preferences use the active FAIR user profile and are stored separately from FAIR workbook values.
- Existing example values are retained as editable suggestions; blank values remain allowed.

## Deployment
Replace the files in the GitHub repository with the contents of this ZIP and allow Streamlit Cloud to redeploy.
