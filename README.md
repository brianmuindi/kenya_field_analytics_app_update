# Kenya Field Analytics

Kenya Field Analytics is a Streamlit-based analytics app for Kenya OSA reporting. It lets teams:

- review OSA by outlet category, region, and account
- inspect root-cause reasons behind gaps
- generate a fully formatted Excel workbook for distribution and review
- use saved MHSKU reference data for MBQ-backed scoring

## Project structure

- `kenya_app/Home.py` — landing page and navigation
- `kenya_app/pages/1_OSA_Analytics.py` — analytics dashboard
- `kenya_app/pages/2_Report_Generator.py` — Excel report generation
- `kenya_app/pages/3_MHSKU_Reference.py` — shared MHSKU reference management
- `kenya_app/engines.py` — data processing and workbook generation
- `kenya_app/auth.py` — login and session handling

## Run locally

```powershell
cd kenya_app
pip install -r requirements.txt
python -m streamlit run Home.py
```

## Host / share

This project runs as a normal Streamlit app. For a public host, deploy the repository to Streamlit Community Cloud or another Python host that supports Streamlit.

### Required deployment secrets

Create your deployed app secrets in the host environment using the same user names and passwords you want to allow.

```toml
[users]
admin = "your_password_here"
analyst = "your_password_here"
```

## Notes

- The app is designed for Kenya-specific OSA reporting and filters out Uganda/Kampala leakage.
- The report generator uses the saved MHSKU workbook logic and produces formatted Excel output.
- Do not commit live credentials or private data into the repository.
