# Creative Audit Report Generator — Web App

A Streamlit web app that connects to Smartsheet, lets your team filter and preview creatives, and generates a branded PPTX report for direct download.

---

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit front-end (filters, preview, download) |
| `report_engine.py` | All report generation logic (slides, charts, AI) |
| `requirements.txt` | Python dependencies |
| `title_template.pdf` | Your branded title slide (upload to repo root) |

---

## Deployment (Streamlit Cloud — free)

### 1. Create a GitHub repo
Create a new **private** GitHub repository and push these four files to the root:
```
app.py
report_engine.py
requirements.txt
title_template.pdf   ← export this from PowerPoint as PDF
```

### 2. Deploy on Streamlit Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.
2. Click **New app**.
3. Select your repo, branch (`main`), and set **Main file path** to `app.py`.
4. Click **Deploy**.

### 3. Add your API keys (Secrets)
In the Streamlit Cloud dashboard, go to your app → **Settings → Secrets** and paste:

```toml
SMARTSHEET_API_KEY = "your_smartsheet_api_key_here"
GEMINI_API_KEY     = "your_gemini_api_key_here"
```

> ⚠️ Never put API keys directly in the code or commit them to GitHub.

### 4. Share the URL
Once deployed, Streamlit gives you a public URL (e.g. `https://your-app-name.streamlit.app`).  
Share it with your team — no login required.

---

## Title slide

Export your branded PowerPoint title slide as a PDF:
1. Open your title template in PowerPoint.
2. **File → Export → Create PDF/XPS** — export only page 1.
3. Name it `title_template.pdf` and place it in the repo root.

The app will automatically overlay the report scope text and date on top of it.

---

## Local development

```bash
pip install -r requirements.txt
# Set env vars (or create a .streamlit/secrets.toml file)
export SMARTSHEET_API_KEY="your_key"
export GEMINI_API_KEY="your_key"
streamlit run app.py
```

`.streamlit/secrets.toml` format (for local dev only, never commit):
```toml
SMARTSHEET_API_KEY = "your_key"
GEMINI_API_KEY     = "your_key"
```
