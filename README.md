# sap-synthetic-dataset-generator

Streamlit web app for generating synthetic SAP-like inventory datasets for safety stock SaaS demos, testing, and inventory optimization experiments.

## Features

- Generates SAP-like tables: `T001`, `T001W`, `T006`, `T134`, `MARA`, `MARC`, `MBEW`, `MATDOC`
- Produces different random datasets on every run using a time-based seed
- Simulates stable, volatile, seasonal, intermittent, and zero-demand materials
- Calculates safety stock into `MARC.EISBE` with ABC-based service levels and lead times
- Exports a ready-to-use Excel workbook named `synthetic_sap_dataset.xlsx`
- Runs directly on Streamlit Cloud

## Local Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud Deployment

1. Push this repository to GitHub.
2. In Streamlit Cloud, click **New app**.
3. Connect your GitHub repository.
4. Select `app.py` as the main file.
5. Click **Deploy**.

## Project Structure

- `app.py`: Streamlit UI and download workflow
- `generator.py`: synthetic SAP dataset generation logic
- `requirements.txt`: Python dependencies
- `README.md`: deployment and usage instructions
