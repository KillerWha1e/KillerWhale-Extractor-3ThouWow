# KillerWhale Extractor — Streamlit Web App

Files:
- `streamlit_app.py` — website interface
- `killerwhale_backend.py` — RE/CE/OATS extraction and Excel logic
- `requirements.txt` — Python packages for deployment

## Run locally

Put the three files in the same folder, then run:

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

A browser window will open with the app.

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `streamlit_app.py`, `killerwhale_backend.py`, and `requirements.txt` to the repository.
3. Sign in to Streamlit Community Cloud and create a new app from that repository.
4. Choose `streamlit_app.py` as the main file.
5. Deploy.

The site will then be reachable from other computers through its Streamlit URL.

## CE filename rule

CE PDF and Excel filenames must match exactly except for the extension, for example:

- `KC Class A AC Line.pdf`
- `KC Class A AC Line.xlsx`

## Privacy

If EMC reports contain customer-confidential or company-confidential data, confirm that external cloud hosting is approved before uploading those files. An internal server deployment can use the same Streamlit files.
