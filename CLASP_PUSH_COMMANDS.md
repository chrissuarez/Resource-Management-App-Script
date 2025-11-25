# Clasp push commands

From the project root, use these to push to the correct Apps Script project:

- Staging (PowerShell): `$env:CLASP_CONFIG_PATH=".clasp.staging.json"; clasp push`
- Production (PowerShell): `$env:CLASP_CONFIG_PATH=".clasp.prod.json"; clasp push`

If you prefer Command Prompt (cmd.exe):

- Staging: `set CLASP_CONFIG_PATH=.clasp.staging.json && clasp push`
- Production: `set CLASP_CONFIG_PATH=.clasp.prod.json && clasp push`
