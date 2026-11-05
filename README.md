# PO-Chem WebApp – Analysis Requests & FAIR Metadata

Application Streamlit for submitting analytical requests to the OV-Chimie–Métabolisme platform and completing the associated FAIR metadata workbook.

The application is bilingual (French/English), stores request files and FAIR workbooks in Nextcloud, and provides separate requester and administrator workflows.

## Main features

### 1. Analysis request

- Multi-step request form for submitting an analytical project.
- French/English interface.
- Analysis types can be managed by the administrator, including English translations for custom types.
- Submitted requests are stored in Nextcloud.
- A Word request document is generated from `template_E-GES-01.docx`.
- The administrator dashboard lists requests and provides access to request details and associated files.
- Request filenames use the following structure:

  `REQ-<IDENTIFIER>-<YYYYMMDD>-<HHMMSS>-<REQUESTER>-<PROJECT>-<LABORATORY>`

  Associated files share the same base name:
  - `...docx` – request document
  - `..._FAIR.xlsx` – FAIR workbook

- The request identifier is retained as the unique reference. Existing files created with older naming conventions are not automatically renamed.

### 2. FAIR workbook

A FAIR workbook is created for each request from `FAIR_template.xlsx` and stored in the same Nextcloud request folder as the Word document.

The workbook contains four sheets:

- `Sample`
- `Analysis`
- `Data treatment`
- `Pipeline Step by Step`

The application edits the corresponding sheets while preserving the workbook structure, labels and formatting as far as supported by the template.

#### Sample

- Metadata are arranged in rows and samples in columns, following the supplied FAIR template.
- Requester workflows include a web form, an interactive table and spreadsheet import/export.
- Sample matching uses `ID tube analyse`.
- The administrator can manage the Q-ToF file identifiers (`File ID Q-Tof ESI+` and `File ID Q-Tof ESI-`).
- Imports are checked against sample identifiers to avoid accidental mismatches.
- Requester and administrator information is merged into the same workbook without overwriting the other role's fields.

#### Analysis

- Web form for entering analytical method information.
- Extraction protocol presets are available, including the base Routaboul et al. (2006) protocol, Proto Massi, Proto Jean Marc and Proto Gilles, as well as a custom option.
- Selected protocol text remains editable before saving.
- The complete protocol description is stored in the `Extraction protocol` field of the FAIR `Analysis` sheet.
- Other analytical information can be completed by either the requester or the administrator.

#### Data treatment

Two mutually exclusive modes are available:

1. **MZmine** – guided web form and full editable POS/NEG parameter table.
2. **Custom** – free-text description of the data-processing workflow.

- The MZmine table supports review and editing of the parameter values.
- The custom mode allows the user to describe the software, processing steps, parameters, filtering, normalization and statistical methods.
- Switching between modes preserves the information entered for the other mode.
- The selected mode and its information are saved in the existing FAIR `Data treatment` worksheet.

#### Pipeline Step by Step

Two modes are available:

1. **Interactive table** – based on the workflow steps in the FAIR template.
2. **None** – leaves the `Pipeline Step by Step` worksheet empty.

The interactive table allows users to record the status of each step and add comments. Statuses can be assigned individually, to all rows at once, or only to selected rows. The predefined workflow can be restored when switching back from `None`.

### 3. FAIR user profiles and preferences

- User profiles are available inside the FAIR area.
- Profiles can be created, added, renamed, selected and deleted.
- Adding a user copies the preferences of the last active profile.
- Preferences are stored separately from request data and FAIR workbook metadata.
- At the bottom of each FAIR sheet (`Analysis`, `Data treatment` and `Pipeline Step by Step`), the user can select **“Keep as preference in my profile”** to save the current interface choices to the active profile.
- Preferences are stored in Nextcloud (`MyDrive/user_profiles.json`) and can be reused during later sessions.
- Saving preferences does not save or modify the experimental answers in a request's FAIR workbook.

### 4. Administrator area

- Authenticated administrator access.
- Request listing and request management.
- Access to the FAIR workbook associated with each request.
- Platform resources are displayed in the sidebar for all users:
  - OV-Chimie–Métabolisme and institutional/partner logos.
  - [Green Chemistry Platform website](https://green-chemistry-platform.versailles-saclay.hub.inrae.fr/).
  - Download of the platform charter (`Charte E-GES-00-ver4.pdf`).

## Repository files

Keep the following files at the root of the GitHub repository:

| File | Purpose |
|---|---|
| `app.py` | Streamlit application |
| `requirements.txt` | Python dependencies |
| `FAIR_template.xlsx` | FAIR workbook template |
| `template_E-GES-01.docx` | Word request document template |
| `Charte E-GES-00-ver4.pdf` | Platform charter available in the sidebar |
| `ov_chimie_logo.png` | Platform logo asset |
| `ov_chimie_logo_user.png` | OV-Chimie logo asset |
| `README.md` | Project documentation |
| `README_DEPLOIEMENT.txt` | Deployment and configuration notes |

## Deployment

1. Update the files in the GitHub repository with the files from the application package.
2. Keep the repository structure and filenames unchanged; the application expects the templates and image assets at the repository root.
3. Check that the existing Streamlit Cloud application points to the correct repository, branch and `app.py` entry point.
4. Allow Streamlit Cloud to redeploy, then test the requester, FAIR and administrator workflows.

Do not remove or replace the existing Streamlit secrets when updating the repository. The application relies on its configured Nextcloud, administrator and SMTP settings.

## Data storage

- Request documents and FAIR workbooks are stored in Nextcloud.
- The FAIR workbook is specific to each request.
- User interface preferences are stored separately in `MyDrive/user_profiles.json`.
- Deleting a request permanently is intended to delete its associated FAIR workbook as well.
- Replacing the application code does not automatically migrate or rename existing request files in Nextcloud.

## Notes

- The workbook layout and field names are based on the supplied `FAIR_template.xlsx`. If the template is changed, the corresponding application mapping may also need to be updated.
- Always verify experimental metadata and protocol descriptions against the actual work performed.
- Test changes on a non-critical request before deploying them for routine use.
