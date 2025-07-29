# BayesTraitsAssist

IMPORTANT USAGE INFORMATION: (Setup instructions below)
- This program creates a trees file using saved labels and pulls values from columns
- It runs BayesTraits twice per comparison and uploads values to Google Sheets (Results sheet)
- Results sheet is cleared of all data if it already exists, save data by relocating it or renaming the Results sheet
- There are two modes: Run all possible comparisons, or Run specific column comparison. For running specific columns, specify the letter of the columns when prompted.

Requirements:
- All N/A data points must be replaced with question marks (?)
- The data must start on column E. Change line 300 in RunProgram.py if otherwise.
- There must not be gaps between columns.
- Sheet must be shared with Google Cloud email
- Sheet name must be entered correctly


FIRST TIME SETUP:

1. Install BayesTraits V5

2. Generate "credentials.json" and connect Sheet to Google Cloud
- Visit the Google Cloud website and login:
- Create a project
- Enable Google Sheets API
- Enable Google Drive API
- Go to APIs & Services > Credentials
- Click “+ CREATE CREDENTIALS” → “Service Account”
- Fill in the form and click “Done”
- Go to the new service account → “Keys” tab
- Click “Add Key” → “Create new key” → Choose JSON*
- Download file and rename it to "credentials.json" if necessary
- Under credentials, copy email.
- Share the Google Sheet you plan to use Bayes Traits with with that email(as editor)

2. Add to BayesTraits folder:
- RunProgram.py
- RunProgram.bat (optional)
- credentials.json

3. Run Program
- Install Python (if needed)
- pip install gspread oauth2client (if needed)
- Launch RunProgram.bat (or other method)
- Enter spreadsheet name

4. (Usually not needed) Fix if Google Cloud stops working
- Create a billing account
- Click "Budgets & alerts"
- Click “Create Budget”
- Set budget amount to $0


