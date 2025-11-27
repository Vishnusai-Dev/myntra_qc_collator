# Myntra QC Collator (Streamlit)

Single-file Streamlit app to collate Failure Summary + Failure Report into core buckets,
auto-loads mapping.xlsx bundled in the repo so users don't need to upload mapping each run.

## How to run locally
1. Create a virtual environment and activate it.
2. Install requirements: `pip install -r requirements.txt`
3. Run `streamlit run app.py`
4. Upload your output Excel on the app UI and click run.


## Auto-deploy / CI

- Streamlit Cloud natively supports connecting a GitHub repo and will auto-deploy on push. After you push this repo to GitHub, go to https://share.streamlit.io and connect this repo to enable automatic deployments on push.
- The repo also contains a GitHub Actions workflow (`.github/workflows/ci_notify.yml`) that installs dependencies and runs a basic import test on push. You can extend it to call a Streamlit Cloud deploy API if you have the appropriate credentials.
