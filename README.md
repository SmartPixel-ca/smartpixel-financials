# SmartPixel Financial Statement Automation Tool

Turns raw QuickBooks exports (Profit & Loss + Balance Sheet) into audited-style, bilingual (FR/EN) financial statements — Word and PDF — automatically.

**How it works:** Claude (the AI) reads the uploaded QuickBooks files and transcribes account codes and dollar amounts — it never decides what anything means. Python code then sorts every account onto the correct financial statement line using a fixed mapping table (`ACCOUNT_MAP` in `app.py`). This split matters: the math never "guesses," so if a number is ever wrong, it's either an account missing from the mapping table, or a genuine judgment call flagged for review — never AI confusion.

Reconciled line-by-line against KPMG-audited statements for FY24-25 — every P&L line and every balance sheet line matches exactly, aside from a handful of figures that come from the accountant's own schedules rather than QuickBooks (see below).

---

## Using the app

The app is hosted on Streamlit Cloud — no installation needed. Open the app URL, fill in the sidebar (company, fiscal year, period end, etc.), upload the P&L and Balance Sheet exports, and click Generate.

**To run it locally instead:**
```bash
git clone https://github.com/SmartPixel-ca/smartpixel-financials.git
cd smartpixel-financials
pip install -r requirements.txt
streamlit run app.py
```

---

## ⚠️ Update these four fields every period

In the sidebar, under **"🧾 Ajustements de l'auditeur,"** four figures can't be calculated from QuickBooks — they come from the accountant's separate schedules. They default to last year's values, so **always update them before generating a new period's statements**:

| Field | Where it comes from |
|---|---|
| Travaux en cours (WIP) | Accountant's percentage-of-completion / revenue recognition schedule |
| Produits reportés (deferred revenue) | Same WIP schedule — ask for both together |
| Tranche à court terme de la dette à long terme | Loan amortization schedules / lender statements — portion due within 12 months |
| Amortissement de l'avantage incitatif lié au bail | Fixed annual lease-incentive amortization, from the lease agreement |

---

## Adding a new QuickBooks account

When QuickBooks gets a new account code, the app flags it in a "needs review" / "unmapped accounts" panel instead of guessing. To fix it:

1. Open [claude.ai](https://claude.ai), upload the current `app.py`.
2. Tell Claude the new account code, its description, and (if known) which statement line it belongs on.
3. Claude edits `ACCOUNT_MAP` and returns the updated file.
4. Publish it to GitHub (see below).

**Prompt template:**
```
Here's the current app.py [attach file]. QuickBooks has a new account,
code [XXXXXXX], named "[account description]". Based on similar existing
accounts already in ACCOUNT_MAP, where should this go? Add it and give
me back the updated file.
```

---

## Making bigger code changes

- **Claude.ai** (easiest) — upload `app.py`, describe the change in plain language, download the updated file.
- **Claude Code** (for larger changes) — install from [claude.com/claude-code](https://claude.com/claude-code), run `claude` in the project folder; it can read the whole codebase, edit, and test changes directly.

Before publishing any change that touches calculations, ask Claude to confirm the balance sheet still balances and the P&L totals are unaffected.

---

## Publishing changes to GitHub

Streamlit Cloud auto-redeploys whenever `app.py` changes on the `main` branch.

1. Go to `github.com/SmartPixel-ca/smartpixel-financials`
2. Click `app.py` → the pencil (✏️) icon to edit
3. Select all, delete, paste in the new version
4. Scroll down, add a short commit message, click **Commit changes**
5. Streamlit Cloud picks it up automatically within 1–2 minutes

> ⚠️ **Don't use "Add file → Upload files"** to replace `app.py` — it sometimes creates a duplicate like `app (1).py` instead of overwriting, which confuses Streamlit about which file to run. Always edit `app.py` directly (steps above).

If a code change needs a new Python package, add it to `requirements.txt` in the same commit, or the app will fail to start.

---

## Troubleshooting

**"Balance sheet doesn't balance" warning** — expected in specific known cases (accounts `2163` and `1090` are presented differently in audited statements than in QuickBooks; the true fix needs an adjusting journal entry not visible in a flat trial balance). Not a bug.

**A number doesn't match what the accountant expects** —
1. Check the "needs review" / "unmapped accounts" panels first.
2. Check the four Auditor Adjustments fields are set to the *current* period, not last year's.
3. Otherwise, ask Claude to help trace it — same process used to verify every line originally.

**App won't start on Streamlit Cloud** — check logs under "Manage app" on [share.streamlit.io](https://share.streamlit.io). Usually a syntax error from a recent edit, or a missing package in `requirements.txt`.

**Two `app.py`-looking files in the repo** — delete the duplicate, confirm Streamlit Cloud points at the correct one under "Manage app" → Settings.

---

## Key files & concepts

| | |
|---|---|
| `app.py` | The entire app — extraction, mapping, statement generation, UI. Normally the only file you edit. |
| `requirements.txt` | Python packages the app needs. |
| `ACCOUNT_MAP` | Dictionary mapping every QuickBooks account code to a statement line — the source of truth (see "Adding a new account"). |
| `SIGN_NORMALIZE` | Declares what sign (+/-) certain accounts should carry in the final statement, independent of how QuickBooks files them. |
| Auditor Adjustments | The four manual-entry figures described above. |
| Streamlit Cloud | Free hosting; auto-redeploys from GitHub `main`. |

For the full walkthrough with screenshots and more detail, see the companion document: *SmartPixel Tool — User and Maintenance Guide.docx*.
