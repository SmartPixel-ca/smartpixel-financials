# Active Média — Financial Statements Tool v2

Upload QuickBooks PDFs → auto-extract → generate a complete audited-style financial package (Word + Excel) with two-year comparison.

---

## What it generates

- Cover page (bilingual FR/EN)
- Income Statement (P&L) — current year vs prior year
- Balance Sheet — current year vs prior year
- Cash Flow Statement
- Annexes (COGS detail, OpEx detail, Financial expenses, Other income)
- Notes to financial statements

---

## Running locally

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

Add your API key in the sidebar.

---

## Deploying to Streamlit Cloud (share with your team)

### Step 1 — GitHub
1. Create a free account at github.com
2. Click "New repository" → name it `smartpixel-financials` → Public → Create
3. Upload these 3 files: `app.py`, `requirements.txt`, `README.md`
   (drag and drop them into the GitHub page)

### Step 2 — Streamlit Cloud
1. Go to share.streamlit.io → sign in with GitHub
2. Click "New app"
3. Select your repository `smartpixel-financials`
4. Main file: `app.py`
5. Click "Advanced settings" → Secrets → paste this:

```toml
ANTHROPIC_API_KEY = "sk-ant-your-actual-key-here"
```

6. Click Deploy

Your app will be live at a URL like:
`https://smartpixel-financials.streamlit.app`

Share that link with anyone — no install, no API key needed on their end.

---

## How to use

1. Open the app link
2. (Optionally) adjust company name, fiscal year, period end in the sidebar
3. Upload current year + prior year QuickBooks PDFs together
4. Click **Extract & Build Full Report**
5. Download Word or Excel

---

## Cost

~$0.05–0.10 per report (depends on PDF length). $10 in API credits = ~100–200 reports.
