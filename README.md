# What If Stress Test Web App

This app turns `WhatIf_StressTest_v4_Fixed.xlsx` into a live web app that:

- pulls real annual statement data from Yahoo Finance for a ticker
- maps that data into the workbook's historical financial input structure
- calculates all workbook scenarios and severities
- produces scenario ratings, dashboard outputs, and a full scenario matrix
- builds a financial ratio scorecard with yearly history and 0-5 star ratings using only Yahoo-reported data

## Run

1. Create a virtual environment if needed:

```bash
python3 -m venv .venv
```

2. Install dependencies:

```bash
./.venv/bin/pip install -r requirements.txt
```

3. Start the app:

```bash
./.venv/bin/streamlit run app.py
```

## Deploy

### Streamlit Community Cloud

Streamlit Community Cloud deploys from a GitHub repository. The current official getting-started flow is:

1. Create or sign in to your Streamlit Community Cloud account.
2. Connect your GitHub account.
3. Push this project to a GitHub repository.
4. In Streamlit Community Cloud, choose that repo and deploy `app.py`.

Official docs:

- [Get started with Streamlit Community Cloud](https://docs.streamlit.io/deploy/streamlit-community-cloud/get-started)

Recommended settings for this project:

- Repository: your GitHub repo containing this folder
- Branch: usually `main`
- Main file path: `app.py`
- Python version: `3.11`

This project is already set up for Community Cloud with:

- `requirements.txt`
- `.streamlit/config.toml`
- `.gitignore`

Before you deploy:

1. Make sure the repo includes:
   - `app.py`
   - `stress_backend.py`
   - `financial_ratios.py`
   - `requirements.txt`
   - `WhatIf_StressTest_v4_Fixed.xlsx`
   - `.streamlit/config.toml`
2. Do not upload local temp files like `~$*.xlsx`.
3. Make sure the app can access the internet, because Yahoo Finance data is fetched live.

After deployment:

- open the app URL from Streamlit Community Cloud
- test a few tickers such as `AAPL`, `NKE`, and `KO`
- confirm the workbook file was included, because the scenario engine depends on it

### Docker / Other Hosts

The app is also packaged for container deployment with `Dockerfile`.

### Option 1: Deploy anywhere with Docker

This is the most portable path. It works for Render, Railway, Fly.io, Google Cloud Run, Azure Container Apps, a VPS with Docker, or your own server.

Build the image:

```bash
docker build -t what-if-stress-test .
```

Run it locally:

```bash
docker run --rm -p 8501:8501 -e PORT=8501 what-if-stress-test
```

Then open `http://localhost:8501`.

### Option 2: Deploy to a container host

Use these settings on the hosting platform:

- Runtime: `Docker`
- Container port: `8501`
- Start command: use the Dockerfile default command
- Health check: `/` is usually enough for Streamlit platforms that support HTTP checks

No fake data, API keys, or external databases are required. The app only needs outbound internet access so it can fetch real Yahoo Finance data.

### Files required in production

Make sure these files are included in the deployment:

- `app.py`
- `stress_backend.py`
- `financial_ratios.py`
- `requirements.txt`
- `WhatIf_StressTest_v4_Fixed.xlsx`
- `.streamlit/config.toml`

### Notes before you deploy

- The server must allow outbound requests to Yahoo Finance through `yahooquery`.
- The app is stateless, so no persistent disk or database is required.
- Some providers put the app to sleep when idle. The first load after sleep may be slower because the app has to start and fetch fresh data.
- If a provider supports environment variables, setting `PORT` is enough. The container already reads it.

## Notes

- All model scenarios are loaded from `WhatIf_StressTest_v4_Fixed.xlsx`.
- Historical data comes from Yahoo Finance via `yahooquery`.
- Outputs are shown in millions of the reported currency.
- The model is intended for operating companies. Banks, insurers, and some other financial businesses may not map cleanly to this statement structure.
- The financial ratio scorecard uses transparent rule-based stars and leaves ratios as `n/a` when Yahoo does not provide the required fields.
