# Clasp push commands

From the project root, use these to push to the correct Apps Script project:

- Staging (PowerShell): `Copy-Item .clasp.staging.json .clasp.json -Force; clasp push`
- Production (PowerShell): `Copy-Item .clasp.prod.json .clasp.json -Force; clasp push`

If you prefer Command Prompt (cmd.exe):

- Staging: `copy /Y .clasp.staging.json .clasp.json && clasp push`
- Production: `copy /Y .clasp.prod.json .clasp.json && clasp push`
