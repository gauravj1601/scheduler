# Weekly Schedule Extractor — Cloud Version

This is the GitHub-ready Streamlit version of the weekly schedule checker.

## What it does
- Lets team members upload a weekly schedule workbook.
- Lets them select batch-wise tabs.
- Extracts `Date | Day | Batch | Start Time | End Time | Subject | Teacher Name`.
- Detects teacher timing clashes.
- Downloads a macro-enabled `.xlsm` file.
- Reuses the tested VBA project from the approved schedule output.
- The downloaded XLSM creates:
  - Teacherwise Schedule
  - Batchwise Schedule
  - `<original Excel name>_teacher schedule` folder with one PDF per teacher
  - `<original Excel name>_Batchwise schedule` folder with one PDF per batch

## App files
```text
app.py
schedule_core.py
requirements.txt
assets/schedule_macro_template.xlsm
.streamlit/config.toml
```

## Upload to GitHub

1. Create a new GitHub repository, for example:
   `weekly-schedule-checker`
2. Keep the repository **Private** if you do not want the code to be publicly visible.
3. Extract this folder and upload all files exactly as they are.
4. Ensure `app.py` and `requirements.txt` are in the repository root.

## Deploy on Streamlit Community Cloud

1. Open `share.streamlit.io`.
2. Sign in with GitHub.
3. Click **Create app**.
4. Select:
   - Repository: your new repository
   - Branch: `main`
   - Main file path: `app.py`
5. Click **Deploy**.
6. Copy the generated app URL and share it with the team.

Streamlit Community Cloud installs Python libraries from `requirements.txt`. The app is cloud-compatible because it does not use Windows Excel or pywin32.

## Recommended access protection

### Option A — Private Streamlit app
Deploy it as a private app and add team members as viewers through Streamlit Cloud.

### Option B — Simple password
In Streamlit Community Cloud, open app settings and add this secret:

```toml
APP_PASSWORD = "Choose-a-strong-team-password"
```

When this secret is present, the app asks every visitor for the password. Do not put the password in GitHub.

## Team usage after download

1. Download the `.xlsm` file.
2. Open it in desktop Microsoft Excel.
3. Click **Enable Content**.
4. Press `Alt + F8`.
5. Run `Create_All_Schedules`.

The separate teacher and batch PDF folders are created on the team member's own computer, beside the downloaded XLSM file.

## Important

- The output macros are embedded in `assets/schedule_macro_template.xlsm`.
- Do not rename the `Schedule` sheet in the downloaded XLSM, because the macro reads it as the source sheet.
- Select only batch-wise timetable tabs. Do not select Faculty/Teacher/Team tabs unless their structure matches the batch-wise tables.
- For very old `.xls` source files, save as `.xlsx` first if Excel data cannot be read correctly.
