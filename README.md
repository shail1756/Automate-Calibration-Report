\
    Calibration Report Automation - Streamlit App
    --------------------------------------------
    This package contains the Streamlit app that reads calibration data from a Google Sheet,
    generates PDF reports (per-instrument) and a ZIP of reports, and provides a dashboard.

    Instructions:
    1. Create a GitHub repo and upload these files.
    2. In Streamlit Cloud, create a new app pointing to this repo (branch: main, file: app.py).
    3. Add your Google service account JSON into Streamlit Secrets under the key `service_account`.
    4. Share your Google Sheet with the service account email (client_email).
    5. Update SHEET_ID and optional LOGO_PATH inside app.py, then redeploy.
