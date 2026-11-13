# tawk.to chat export (Streamlit on hcode)

Exports all chats in a date range to Excel: name, email, date, content, duration (+ chat ID).

## Deploy on hcode
Follow the Streamlit KB article. Settings that matter:
- Start command: `streamlit run app.py --server.port=$PORT --server.address=0.0.0.0 --server.headless=true`
- Resources: **1 GB memory** recommended (512 MB is tight when the Excel file is built).
- Optional environment variables: `TAWK_API_KEY`, `TAWK_PROPERTY_ID`, `APP_PASSWORD` (protects the public URL), `DATA_DIR` (default `data`).

## Use
1. Open the app, enter the API key and Property ID (or set them as env vars).
2. Press **Test 1 chat** and check name, email, date, duration and content look right.
3. Press **Start / Resume**. You can close the tab; the job keeps running on the server.
4. Press **Build Excel**, then **Download Excel**.

Progress is saved in `data/export.db`. If the app restarts, press Start again and it continues
where it stopped (only if hcode keeps the disk between restarts; otherwise download the Excel before redeploying).
