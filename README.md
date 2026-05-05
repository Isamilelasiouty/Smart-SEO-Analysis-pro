# 📊 Smart SEO Analysis Tool

A professional SEO analysis dashboard built with Streamlit. Supports Google Search Console exports, Looker Studio data, and direct Google Sheets connection.

## ✨ Features

- **3 Data Input Methods:**
  - 📁 Upload CSV / Excel directly
  - 🌐 Google Sheets (Public URL — no auth needed)
  - 🔐 Google Sheets (Service Account — private sheets)

- **13 Analysis Reports:**
  - Overview KPIs
  - Top Pages & Keywords
  - Smart Keyword Clustering
  - Brand vs Non-brand Split
  - Search Intent Analysis
  - Opportunity Score (CTR gap model)
  - CTR Problems detection
  - Keyword Cannibalization
  - Position Distribution
  - Declining Keywords & Pages (requires comparison tab)

- **Visual Dashboard:** 7 interactive Plotly charts
- **Export:** Excel (all tabs) + PDF summary
- **Arabic & English** column detection
- **Dark professional UI**

## 🚀 Deploy on Streamlit Cloud

1. Fork this repo
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repo
4. Set main file: `app.py`
5. Deploy!

## 💻 Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 📋 Supported Data Formats

| Column | Alternative Names |
|--------|------------------|
| Query string | keyword, search query, search term |
| Landing page | page, url, link |
| Clicks | clicks, نقر |
| Impressions | impressions, ظهور |
| CTR | click through rate, نسبة النقر |
| Position | avg position, rank, ترتيب |

## 📉 Comparison Tab (for decline analysis)

Add a second tab to your Excel/Sheets with columns:
- `Clicks Current` + `Clicks Previous`
- `Impressions Current` + `Impressions Previous`
- `Position Current` + `Position Previous`

Or use GSC's "Compare" export which automatically includes both periods.

## 🔐 Service Account Setup

1. Google Cloud Console → Enable Sheets API + Drive API
2. Create Service Account → Download JSON key
3. Share your Sheet with the service account email as Viewer
4. Paste JSON content in the app

## 📦 Requirements

See `requirements.txt` — all standard packages, no proprietary dependencies.
