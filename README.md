# HCRE Wellness Center – Cash Flow App

Interactive dashboard + actuals entry tool for the HCRE Wellness Center project budget and RMH Mezzanine facility cash flow.

## Features

- **Dashboard** – KPIs, charts, monthly cash & debt snapshot  
- **Enter Actuals** – override any cost cell; dashboard updates live  
- **Line Items** – full detail view  
- **Export** – CSV downloads of current state  

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Place `Cash Flow Forecast.xlsx` in the **same folder** as `app.py`.

## Deploy to Streamlit Community Cloud

### Recommended repo layout (flat)

```
your-repo/
├── app.py
├── requirements.txt
├── README.md
└── Cash Flow Forecast.xlsx
```

### Steps

1. Create a **public** GitHub repository.
2. Upload the four files above (all at the root of the repo).
3. Go to [https://share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
4. Click **New app**:
   - Repository → your repo  
   - Branch → `main`  
   - Main file path → `app.py`
5. Click **Deploy**.

Your app will be live at a URL like `https://your-app-name.streamlit.app`.

### Notes for Cloud

- The Excel file is read-only. Overrides exist only for the current browser session.
- Use the **Export** tab to download CSVs if you need to keep changes.
- Free Streamlit Cloud requires a **public** repository.

## Facility

RMH Mezzanine capital facility limit: **R 10,000,000**.
