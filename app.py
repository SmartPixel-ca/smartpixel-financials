import streamlit as st
import anthropic
import base64
import json
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

st.set_page_config(page_title="Active Média — Financial Statements", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
html,body,[class*="css"]{font-family:'Inter',sans-serif;}
.main-header{background:linear-gradient(135deg,#1A2E4A 0%,#2E75B6 100%);padding:2rem 2.5rem;border-radius:12px;margin-bottom:2rem;color:white;}
.main-header h1{margin:0;font-size:1.6rem;font-weight:600;}
.main-header p{margin:.3rem 0 0;opacity:.75;font-size:.88rem;}
.kpi-card{background:white;border:1px solid #E8ECF0;border-radius:10px;padding:1.2rem 1.4rem;}
.kpi-label{font-size:.7rem;font-weight:600;text-transform:uppercase;letter-spacing:.06em;color:#7A8799;margin-bottom:4px;}
.kpi-value{font-size:1.5rem;font-weight:600;color:#1A2E4A;font-variant-numeric:tabular-nums;}
.kpi-sub{font-size:.73rem;color:#7A8799;margin-top:2px;}
.kpi-pos{color:#1A7A4A;} .kpi-neg{color:#C0392B;}
.info-box{background:#EBF4FF;border-left:3px solid #2E75B6;padding:.9rem 1.1rem;border-radius:0 8px 8px 0;margin:1rem 0;font-size:.88rem;color:#1A2E4A;}
</style>
""", unsafe_allow_html=True)

def get_api_key():
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY","")
        if key and not key.startswith("sk-ant-your"): return key
    except: pass
    return st.session_state.get("api_key","")

st.markdown("""<div class="main-header"><h1>📊 Financial Statements</h1>
<p>Active Média inc. — SmartPixel &nbsp;·&nbsp; Upload QuickBooks PDFs or Excel to generate a complete audited-style report</p></div>""",unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### ⚙️ Settings")
    if not get_api_key():
        mk = st.text_input("Anthropic API key",type="password",placeholder="sk-ant-...")
        if mk: st.session_state["api_key"]=mk
    else:
        st.success("API key configured ✓")
    st.markdown("---")
    st.markdown("**Company info**")
    company_name = st.text_input("Company name",value="Active Média inc.")
    fiscal_year  = st.text_input("Current fiscal year",value="2026")
    prior_year   = st.text_input("Prior fiscal year",value="2025")
    period_end   = st.text_input("Period end date",value="January 31, 2026")
    preparer     = st.text_input("Prepared by",value="Management")
    currency     = st.selectbox("Currency",["CAD","USD","EUR"],index=0)
    st.markdown("---")
    st.markdown("**Output format**")
    out_format = st.radio("Format",["PDF (audited style)","Word (.docx)","Both"],index=0)
    st.markdown("---")
    st.caption("SmartPixel Financial Tool · v3.0")

settings = dict(company_name=company_name,fiscal_year=fiscal_year,prior_year=prior_year,
                period_end=period_end,preparer=preparer,currency=currency)

# ── Account grouping map ──────────────────────────────────────────────────────
# Maps QuickBooks account code prefixes/names to consolidated line items
GROUPING = {
    # ANNEXE 1 — Coût des ventes
    "ann1_achats": {
        "label_fr": "Achats", "label_en": "Purchases",
        "codes": ["5500301","5501101","5999","Inventory"]
    },
    "ann1_salaires": {
        "label_fr": "Salaires et avantages sociaux", "label_en": "Salaries and benefits",
        "codes": ["5000101","5000301","5000401","5000901","5015301","5015401","5015901",
                  "5020301","5020401","5020901","5025301"]
    },
    "ann1_soustraitance": {
        "label_fr": "Sous-traitance", "label_en": "Subcontracting",
        "codes": ["5050101","5050401","5050501","5050901","5075101"]
    },
    "ann1_livraison": {
        "label_fr": "Frais de livraison", "label_en": "Delivery costs",
        "codes": ["6016401"]
    },
    "ann1_logiciel": {
        "label_fr": "Logiciel", "label_en": "Software",
        "codes": ["7500301","7500311","7500401","7500901","7501401","8300701","5100201"]
    },
    # ANNEXE 2 — Frais d'exploitation
    "ann2_salaires_rd": {
        "label_fr": "Salaires et avantages sociaux — recherche et développement",
        "label_en": "Salaries and benefits — research and development",
        "codes": ["5001501","5001701","5001801","5001901","5056501",
                  "5015101","5015201","5015501","5015601","5015701","5015801",
                  "5020101","5020201","5020601","5020701","5020801"]
    },
    "ann2_credits_rd": {
        "label_fr": "Crédits d'impôt pour la recherche et le développement",
        "label_en": "R&D tax credits",
        "codes": ["5100301","5100501","5100601","5100701"]
    },
    "ann2_salaires_admin": {
        "label_fr": "Salaires et avantages sociaux — ventes et administration",
        "label_en": "Salaries and benefits — sales and administration",
        "codes": ["5000201","5000501","5000601","5000701","5000801",
                  "5010101","5010102","5010301","5011101","5011401","5012101","6100201"]
    },
    "ann2_deplacement": {
        "label_fr": "Frais de déplacement", "label_en": "Travel expenses",
        "codes": ["6015601","6016101","6016104","6016201","6016301","6016601",
                  "6020201","6020301","6020401","6020601",
                  "6010101","6010201","6010301","6010401",
                  "6005101","6005102","6005201","6005301","6005401","6005601",
                  "8010601"]
    },
    "ann2_publicite": {
        "label_fr": "Publicité et promotion", "label_en": "Advertising and promotion",
        "codes": ["7000201","7005201","7010201","7010202","7025601"]
    },
    "ann2_honoraires": {
        "label_fr": "Honoraires", "label_en": "Professional fees",
        "codes": ["8050501","8050601","8050701","8050801","8050901","8055601","8060601","8065101","8065301","8065601"]
    },
    "ann2_loyer": {
        "label_fr": "Loyer", "label_en": "Rent",
        "codes": ["8000601"]
    },
    "ann2_telecom": {
        "label_fr": "Télécommunications", "label_en": "Telecommunications",
        "codes": ["8400601","8405101","8405201","8405301","8405401","8405601",
                  "7515601","7520601","7525501","7525601"]
    },
    "ann2_representation": {
        "label_fr": "Frais de représentation", "label_en": "Entertainment",
        "codes": ["6000101","6000201","6005101","6005201","6005301","6005401","6005601",
                  "6100101","6100201"]
    },
    "ann2_bureau": {
        "label_fr": "Frais de bureau", "label_en": "Office expenses",
        "codes": ["8150601","8152601","8155601","8200601","5505301","8110601","9150601"]
    },
    "ann2_cotisation": {
        "label_fr": "Cotisation et abonnement", "label_en": "Dues and subscriptions",
        "codes": ["8350601"]
    },
    "ann2_courrier": {
        "label_fr": "Courrier et frais postaux", "label_en": "Postage and courier",
        "codes": ["8100601","8100605"]
    },
    "ann2_assurance": {
        "label_fr": "Assurances", "label_en": "Insurance",
        "codes": ["8305601"]
    },
    "ann2_taxes": {
        "label_fr": "Taxes et permis", "label_en": "Taxes and permits",
        "codes": ["8300601"]
    },
    "ann2_licence": {
        "label_fr": "Licence", "label_en": "Licences",
        "codes": ["7505601","7510101","7510201","7510601","7500701","7500801","7501801"]
    },
    "ann2_entretien": {
        "label_fr": "Entretien et réparations", "label_en": "Maintenance and repairs",
        "codes": ["8200601"]
    },
    "ann2_paie": {
        "label_fr": "Frais de gestion de paie", "label_en": "Payroll admin fees",
        "codes": ["8105601","8105605"]
    },
    "ann2_fx": {
        "label_fr": "(Gain) Perte de change", "label_en": "Foreign exchange (gain) loss",
        "codes": ["8450601","5830"]
    },
    "ann2_interet": {
        "label_fr": "Intérêts et pénalités", "label_en": "Interest and penalties",
        "codes": ["8460601"]
    },
    "ann2_representant": {
        "label_fr": "Représentant externe", "label_en": "External representatives",
        "codes": ["5030101","5030104","5030201"]
    },
    # ANNEXE 3 — Frais financiers
    "ann3_interet_lt": {
        "label_fr": "Intérêts sur la dette à long terme", "label_en": "Interest on long-term debt",
        "codes": ["8520601","8520801","8510701"]
    },
    "ann3_frais_bancaires": {
        "label_fr": "Intérêts et frais bancaires", "label_en": "Bank charges and interest",
        "codes": ["8500601","8500605","8500811","8515601"]
    },
    # ANNEXE 4 — Autres revenus
    "ann4_aide_gouv": {
        "label_fr": "Aide gouvernementale", "label_en": "Government assistance",
        "codes": ["4408801","4405601","4407801"]
    },
    "ann4_autres": {
        "label_fr": "Autres revenus", "label_en": "Other revenue",
        "codes": ["4400601","4400901"]
    },
}

def fmt_num(n):
    if n is None or n == 0: return "—"
    try: n = float(n)
    except: return "—"
    s = f"{abs(n):,.0f}".replace(",", " ")
    return f"({s})" if n < 0 else s

def fmt_with_dollar(n):
    v = fmt_num(n)
    if v == "—": return "— $"
    return f"{v} $"

def pct(a,b):
    try: return f"{float(a)/float(b)*100:.1f}%"
    except: return "—"

# ── Extraction prompt ─────────────────────────────────────────────────────────
def build_prompt():
    return """CRITICAL INSTRUCTION: You must respond with ONLY a valid JSON object. No explanation, no steps, no text before or after the JSON. Start your response with { and end with }. Do not write any words, headers, bullet points or analysis — ONLY the JSON object.

IMPORTANT: Only include accounts with a non-zero balance in your calculations. Ignore all accounts showing $0.00. Do not add deleted accounts (marked as 'deleted') to any group.

Extract ALL financial data from these QuickBooks PDFs/Excel for Active Média inc. / SmartPixel.

CRITICAL GROUPING RULES — group accounts exactly as follows:

INCOME (Ventes): All 4000xxx, 4025xxx, 4050xxx, 4075xxx revenue accounts → single total "Ventes"
Include discounts (rabais 4200xxx, 4225xxx, 4250xxx, 4275xxx) as reductions within revenue.

ANNEXE 1 — Coût des ventes:
- Achats: 5500xxx, 5501xxx, 5999, Inventory Shrinkage
- Salaires et avantages sociaux: ALL 5000x01/301/401/901, 5015x01/301/401/901, 5020x01/301/401/901, 5025xxx in COGS
- Sous-traitance: ALL 5050xxx, 5075xxx
- Frais de livraison: 6016401
- Logiciel: ALL 7500xxx, 7501xxx, 8300xxx, 5100201 in COGS

ANNEXE 2 — Frais d'exploitation:
- Salaires et avantages sociaux — R&D: 5001501, 5001701, 5001801, 5001901, 5056501 + benefits 5015101/201/501/601/701/801 + 5020101/201/601/701/801
- Crédits d'impôt R&D: 5100301, 5100501, 5100601, 5100701 (all NEGATIVE)
- Salaires et avantages sociaux — ventes et admin: 5000201,501,601,701,801 + commissions 5010xxx + 5011xxx + 5012xxx + 6100201 (comité social)
- Frais de déplacement: ALL 6015xxx, 6016xxx (excl 6016401), 6020xxx, 6010xxx, 6005xxx (repas), 8010601 (stationnement)
- Publicité et promotion: ALL 7000xxx, 7005xxx, 7010xxx, 7025xxx
- Honoraires: ALL 8050xxx, 8055xxx, 8060xxx, 8065xxx
- Loyer: 8000601 only (8010601 stationnement goes in déplacement)
- Télécommunications: ALL 8400xxx, 8405xxx, 7515xxx, 7520xxx, 7525xxx
- Frais de représentation: 6000xxx, 6100101 only (NOT 6100201 — that goes in salaires ventes et admin)
- Frais de bureau: 8150xxx, 8152xxx, 8155xxx, 8200xxx, 5505xxx, 8110xxx, 9150xxx
- Cotisation et abonnement: 8350xxx
- Assurances: 8305xxx
- Taxes et permis: 8300601
- Licence: 7505xxx, 7510xxx, 7500701, 7500801, 7501801
- Frais de gestion de paie: 8105xxx
- (Gain) Perte de change: 8450xxx, 5830 (Gain/Perte sur change)
- Intérêts et pénalités: 8460xxx
- Représentant externe: ALL 5030xxx
- Courrier: 8100xxx

AMORTISSEMENT: 9000xxx → separate line on P&L

ANNEXE 3 — Frais financiers:
- Intérêts sur la dette à long terme: 8520xxx, 8510701 (intérêt marge de crédit)
- Intérêts et frais bancaires: 8500xxx, 8515601 (intérêt carte crédit)

ANNEXE 4 — Autres revenus:
- Aide gouvernementale: 4408xxx, 4405601, 4407801
- Autres revenus: 4400xxx only

BALANCE SHEET groupings — follow the mapping EXACTLY:

ACTIF À COURT TERME:
- Encaisse: 1055, 1056, 1057, 1058, 1059, 1060, 1075, 1076, 1080, 1090, 1499, 1700, 2641
  (1065 BMO Autres is deleted — exclude it. 2641 Marge de crédit DesJardins nets against cash as per QuickBooks structure)
- Comptes clients et autres créances (note 2): 1200, 1201, 1202, 1203, 1205
- Stocks: 1999
- Travaux en cours (note 15): 1250, 1251 (1251 Produits perçus d'avance A.R.Tx is netted here)
- Crédits d'impôt à recevoir: 1210, 1220, 1225, 1226, 1230, 1231, 1232
- Charges payées d'avance: 1300, 1320

ACTIF À LONG TERME (listed individually on balance sheet, NOT grouped):
- Dépôts à long terme: 1710 only
- Frais payés d'avance (LT): 1711 only (show as "-" if zero)
- Immobilisations corporelles (note 3) — NET value: 1810+1811, 1815+1816, 1820+1821, 1825+1826, 1835+1836, 1840+1841, 1845+1846, 1855+1856, 1891
- Actifs incorporels (note 4): 1880, 1890
- Impôts futurs (note 14): 2163
- Avances à la filiale: 1901, 1903
- Avances aux actionnaires: 1750
- Placement dans la filiale (note 5): 1900

PASSIF À COURT TERME:
- Emprunt bancaire (note 6): 2630, 2635 (show as "-" if zero)
- Comptes fournisseurs et charges à payer (note 7): 2100, 2101, 2102, 2155, 2136 (all A/P + accruals + credit cards)
- Impôts sur le bénéfice à payer: 2160, 2161
- Produits reportés (note 15): 2170, 2180, 2190, 2200, 2205, 2210, 2212, 2234, 2300, 2305,
  2310, 2315, 2340, 2345, 2350, 2355, 2367, 2455, 2684, 2685, 2686, 2687, 2688,
  BC Ministry of Finance Payable, BC Ministry of Finance Suspense, PST BC Payable
- Tranche à court terme de la dette à long terme (note 8): compute from LT loans — current portion

PASSIF À LONG TERME:
- Avantages incitatifs liés aux baux: 2400
- Dette à long terme (note 8): 2620, 2621, 2623, 2624, 2625, 2626, 2627, 2628, 2629, 2631, 2680, 2681, 2682, 2683

AVOIR DES ACTIONNAIRES:
- Capital-actions (note 9): 3350, 3351, 3551, 3570, 3580
- Déficit: 3560 + Profit for the year (current year net income)

Return ONLY valid JSON:
{
  "company": "string",
  "fiscal_year": "string",
  "prior_year": "string",
  "period_end": "string",
  "currency": "CAD",
  "pl": {
    "ventes": {"current": number, "prior": number},
    "cout_des_ventes": {"current": number, "prior": number},
    "benefice_brut": {"current": number, "prior": number},
    "frais_exploitation": {"current": number, "prior": number},
    "frais_financiers": {"current": number, "prior": number},
    "amortissement": {"current": number, "prior": number},
    "total_charges": {"current": number, "prior": number},
    "benefice_avant_autres": {"current": number, "prior": number},
    "autres_revenus": {"current": number, "prior": number},
    "benefice_net": {"current": number, "prior": number}
  },
  "ann1": {
    "achats": {"current": number, "prior": number},
    "salaires": {"current": number, "prior": number},
    "soustraitance": {"current": number, "prior": number},
    "livraison": {"current": number, "prior": number},
    "logiciel": {"current": number, "prior": number},
    "total": {"current": number, "prior": number}
  },
  "ann2": {
    "salaires_rd": {"current": number, "prior": number},
    "credits_rd": {"current": number, "prior": number},
    "salaires_admin": {"current": number, "prior": number},
    "deplacement": {"current": number, "prior": number},
    "location_equip": {"current": number, "prior": number},
    "publicite": {"current": number, "prior": number},
    "honoraires": {"current": number, "prior": number},
    "loyer": {"current": number, "prior": number},
    "telecom": {"current": number, "prior": number},
    "representation": {"current": number, "prior": number},
    "bureau": {"current": number, "prior": number},
    "cotisation": {"current": number, "prior": number},
    "assurance": {"current": number, "prior": number},
    "taxes": {"current": number, "prior": number},
    "entretien": {"current": number, "prior": number},
    "courrier": {"current": number, "prior": number},
    "lease_incentive": {"current": number, "prior": number},
    "licence": {"current": number, "prior": number},
    "paie": {"current": number, "prior": number},
    "formation": {"current": number, "prior": number},
    "fx": {"current": number, "prior": number},
    "interet_penalites": {"current": number, "prior": number},
    "representant": {"current": number, "prior": number},
    "total": {"current": number, "prior": number}
  },
  "ann3": {
    "interet_lt": {"current": number, "prior": number},
    "frais_bancaires": {"current": number, "prior": number},
    "total": {"current": number, "prior": number}
  },
  "ann4": {
    "aide_gouv": {"current": number, "prior": number},
    "autres": {"current": number, "prior": number},
    "total": {"current": number, "prior": number}
  },
  "bs": {
    "encaisse": {"current": number, "prior": number},
    "comptes_clients": {"current": number, "prior": number},
    "stocks": {"current": number, "prior": number},
    "travaux_en_cours": {"current": number, "prior": number},
    "credits_impot": {"current": number, "prior": number},
    "charges_payees_avance": {"current": number, "prior": number},
    "total_actif_ct": {"current": number, "prior": number},
    "depots_lt": {"current": number, "prior": number},
    "frais_payes_avance_lt": {"current": number, "prior": number},
    "immobilisations": {"current": number, "prior": number},
    "actifs_incorporels": {"current": number, "prior": number},
    "impots_futurs": {"current": number, "prior": number},
    "avances_filiale": {"current": number, "prior": number},
    "avances_actionnaires": {"current": number, "prior": number},
    "placement_filiale": {"current": number, "prior": number},
    "total_actif": {"current": number, "prior": number},
    "emprunt_bancaire": {"current": number, "prior": number},
    "comptes_fournisseurs": {"current": number, "prior": number},
    "impots_benefice": {"current": number, "prior": number},
    "produits_reportes": {"current": number, "prior": number},
    "tranche_ct_lt": {"current": number, "prior": number},
    "total_passif_ct": {"current": number, "prior": number},
    "avantages_baux": {"current": number, "prior": number},
    "dette_lt": {"current": number, "prior": number},
    "total_passif": {"current": number, "prior": number},
    "capital_actions": {"current": number, "prior": number},
    "deficit": {"current": number, "prior": number},
    "total_avoir": {"current": number, "prior": number},
    "total_passif_avoir": {"current": number, "prior": number}
  }
}
Use 0 for missing values. prior year = 0 if only one year available."""

def extract(files):
    client = anthropic.Anthropic(api_key=get_api_key())
    content = []
    for f in files:
        data_bytes = f.read()
        b64 = base64.b64encode(data_bytes).decode()
        # Always send as PDF document type (Claude supports both)
        content.append({
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
            "title": f.name
        })
    content.append({"type": "text", "text": build_prompt()})
    
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=16000,
        system="You are a financial data extraction API. You ONLY output valid JSON. Never write explanations, steps, headers, or any text outside the JSON object. Your entire response must be parseable by json.loads().",
        messages=[{"role": "user", "content": content}]
    )
    
    if not msg.content:
        raise ValueError("Claude returned an empty response. Try uploading one PDF at a time.")
    
    raw = msg.content[0].text.strip()
    if not raw:
        raise ValueError("Claude returned empty text. The PDFs may be too large — try one at a time.")
    
    # Clean markdown fences
    raw = raw.lstrip("```json").lstrip("```").rstrip("```").strip()
    
    try:
        return json.loads(raw)
    except json.JSONDecodeError as e:
        # Try to repair truncated JSON
        last = raw.rfind("}")
        if last > 0:
            fixed = raw[:last+1]
            open_cnt = fixed.count("{") - fixed.count("}")
            try:
                return json.loads(fixed + "}" * open_cnt)
            except:
                pass
        raise ValueError(f"Could not parse Claude response as JSON: {e}\n\nRaw response (first 500 chars):\n{raw[:500]}")

# ── PDF Builder ──────────────────────────────────────────────────────────────
def build_pdf(data, s):
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, HRFlowable, KeepTogether
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER
    from reportlab.lib import colors
    from reportlab.lib.units import inch
    from reportlab.lib.pagesizes import letter

    BLACK = colors.black
    GREY  = colors.HexColor("#666666")

    cy = data.get("fiscal_year", s["fiscal_year"])
    py = data.get("prior_year",  s["prior_year"])
    period_str = f"Exercice clos le {s['period_end']}, avec informations comparatives de {py}"

    def ps(name, size=10, bold=False, color=BLACK, align=TA_LEFT):
        return ParagraphStyle(name + str(id(name)),
            fontName="Helvetica-Bold" if bold else "Helvetica",
            fontSize=size, textColor=color, alignment=align,
            leading=size * 1.4, spaceAfter=0, spaceBefore=0)

    # Column widths — label takes most space, two equal value columns
    LW  = 4.2 * inch
    VW  = 1.35 * inch
    CW  = [LW, VW, VW]
    TW  = LW + VW + VW  # total table width

    def hdr_para(text, size=10, bold=False, color=BLACK, align=TA_LEFT):
        return Paragraph(text, ps("h", size, bold, color, align))

    def val_para(text, bold=False, color=BLACK):
        return Paragraph(f'<para alignment="right">{text}</para>',
            ps("v", 9.5, bold, color, TA_RIGHT))

    def lbl_para(text, bold=False, indent=0, color=BLACK):
        style = ps("l", 9.5, bold, color, TA_LEFT)
        style.leftIndent = indent
        return Paragraph(text, style)

    # Table style base
    BASE_STYLE = [
        ('TOPPADDING',    (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING',   (0,0), (-1,-1), 0),
        ('RIGHTPADDING',  (0,0), (-1,-1), 0),
        ('VALIGN',        (0,0), (-1,-1), 'MIDDLE'),
        ('FONTNAME',      (0,0), (-1,-1), 'Helvetica'),
        ('FONTSIZE',      (0,0), (-1,-1), 9.5),
    ]

    def build_table(rows):
        """
        rows: list of (label_cell, v1_cell, v2_cell, style_type)
        style_type: 'normal'|'indent'|'section'|'subtotal'|'grand'|'blank'
        """
        tdata = []
        cmds  = list(BASE_STYLE)

        # Column header row
        tdata.append([
            hdr_para("", 9, False, GREY),
            val_para(cy, bold=True, color=BLACK),
            val_para(py, bold=True, color=BLACK),
        ])
        cmds += [
            ('LINEBELOW', (0,0), (-1,0), 0.75, BLACK),
            ('TOPPADDING', (0,0), (-1,0), 4),
            ('BOTTOMPADDING', (0,0), (-1,0), 4),
        ]

        for i, (lbl, v1, v2, stype) in enumerate(rows):
            rn = i + 1
            if stype == 'blank':
                tdata.append([Paragraph("", ps("b",3)), Paragraph("",ps("b2",3)), Paragraph("",ps("b3",3))])
                cmds.append(('FONTSIZE',(0,rn),(-1,rn),3))
                cmds.append(('TOPPADDING',(0,rn),(-1,rn),1))
                cmds.append(('BOTTOMPADDING',(0,rn),(-1,rn),1))
                continue

            if stype == 'section':
                tdata.append([lbl_para(lbl, bold=False, color=BLACK), val_para(""), val_para("")])
                continue

            bold_row = stype in ('grand', 'subtotal')
            tdata.append([lbl, val_para(v1, bold=bold_row), val_para(v2, bold=bold_row)])

            if stype == 'grand':
                cmds += [
                    ('LINEABOVE', (0,rn), (-1,rn), 0.75, BLACK),
                    ('LINEBELOW', (0,rn), (-1,rn), 1.5,  BLACK),
                    ('FONTNAME',  (0,rn), (-1,rn), 'Helvetica-Bold'),
                ]
            elif stype == 'subtotal':
                cmds += [
                    ('LINEABOVE', (0,rn), (-1,rn), 0.5, BLACK),
                    ('FONTNAME',  (0,rn), (-1,rn), 'Helvetica-Bold'),
                ]

        t = Table(tdata, colWidths=CW)
        t.setStyle(TableStyle(cmds))
        return t

    def v(obj, key):
        d = (obj or {}).get(key, {})
        return (d.get("current", 0) or 0) if isinstance(d, dict) else (d or 0)

    def vp(obj, key):
        d = (obj or {}).get(key, {})
        return (d.get("prior", 0) or 0) if isinstance(d, dict) else (d or 0)

    pl = data.get("pl", {}); a1 = data.get("ann1", {})
    a2 = data.get("ann2", {}); a3 = data.get("ann3", {}); a4 = data.get("ann4", {})

    story = []

    def page_hdr(title, period):
        story.append(Paragraph(s["company_name"].upper(), ps("co", 16, bold=True)))
        story.append(Spacer(1, 2))
        story.append(Paragraph(title, ps("ti", 10)))
        story.append(Spacer(1, 2))
        story.append(HRFlowable(width=TW, thickness=0.75, color=BLACK, spaceAfter=3, spaceBefore=0))
        story.append(Paragraph(period, ps("pe", 9, color=GREY)))
        story.append(Spacer(1, 10))

    # ── P&L ───────────────────────────────────────────────────────────────────
    page_hdr("État non consolidé des résultats", period_str)

    pl_rows = [
        (lbl_para("Ventes (notes 11 et 12)"), fmt_with_dollar(v(pl,"ventes")), fmt_with_dollar(vp(pl,"ventes")), "normal"),
        (lbl_para(""), "", "", "blank"),
        (lbl_para("Coût des ventes (annexe 1)", indent=12), fmt_num(v(pl,"cout_des_ventes")), fmt_num(vp(pl,"cout_des_ventes")), "normal"),
        (lbl_para(""), fmt_num(v(pl,"benefice_brut")), fmt_num(vp(pl,"benefice_brut")), "subtotal"),
        (lbl_para(""), "", "", "blank"),
        ("Charges", "", "", "section"),
        (lbl_para("    Frais d'exploitation (annexe 2)", indent=20), fmt_num(v(pl,"frais_exploitation")), fmt_num(vp(pl,"frais_exploitation")), "normal"),
        (lbl_para("    Frais financiers (annexe 3)", indent=20), fmt_num(v(pl,"frais_financiers")), fmt_num(vp(pl,"frais_financiers")), "normal"),
        (lbl_para("    Amortissement des immobilisations corporelles et actifs incorporels", indent=20), fmt_num(v(pl,"amortissement")), fmt_num(vp(pl,"amortissement")), "normal"),
        (lbl_para(""), fmt_num(v(pl,"total_charges")), fmt_num(vp(pl,"total_charges")), "subtotal"),
        (lbl_para(""), "", "", "blank"),
        (lbl_para("Bénéfice (perte) avant les autres revenus"), fmt_num(v(pl,"benefice_avant_autres")), fmt_num(vp(pl,"benefice_avant_autres")), "normal"),
        (lbl_para(""), "", "", "blank"),
        (lbl_para("Autres revenus (annexe 4)"), fmt_num(v(pl,"autres_revenus")), fmt_num(vp(pl,"autres_revenus")), "normal"),
        (lbl_para(""), "", "", "blank"),
        (lbl_para("Bénéfice net (perte nette)", bold=True), fmt_with_dollar(v(pl,"benefice_net")), fmt_with_dollar(vp(pl,"benefice_net")), "grand"),
    ]
    story.append(build_table(pl_rows))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Se reporter aux notes afférentes aux états financiers.", ps("note", 8, color=GREY)))
    story.append(PageBreak())

    # ── ANNEXES 1 & 2 ─────────────────────────────────────────────────────────
    page_hdr("Annexes", period_str)

    story.append(Paragraph("<b>Annexe 1 - Coût des ventes</b>", ps("a1", 10, bold=True)))
    story.append(Spacer(1, 5))
    ann1_rows = [
        (lbl_para("Achats"), fmt_with_dollar(v(a1,"achats")), fmt_with_dollar(vp(a1,"achats")), "normal"),
        (lbl_para("Salaires et avantages sociaux"), fmt_num(v(a1,"salaires")), fmt_num(vp(a1,"salaires")), "normal"),
        (lbl_para("Sous-traitance"), fmt_num(v(a1,"soustraitance")), fmt_num(vp(a1,"soustraitance")), "normal"),
        (lbl_para("Frais de livraison"), fmt_num(v(a1,"livraison")), fmt_num(vp(a1,"livraison")), "normal"),
        (lbl_para("Logiciel"), fmt_num(v(a1,"logiciel")), fmt_num(vp(a1,"logiciel")), "normal"),
        (lbl_para("", bold=True), fmt_with_dollar(v(a1,"total")), fmt_with_dollar(vp(a1,"total")), "grand"),
    ]
    story.append(build_table(ann1_rows))
    story.append(Spacer(1, 16))

    story.append(Paragraph("<b>Annexe 2 - Frais d'exploitation</b>", ps("a2", 10, bold=True)))
    story.append(Spacer(1, 5))

    def a2row(label, key):
        return (lbl_para(label), fmt_num(v(a2,key)), fmt_num(vp(a2,key)), "normal")

    ann2_rows = [
        (lbl_para("Salaires et avantages sociaux - recherche et développement"), fmt_with_dollar(v(a2,"salaires_rd")), fmt_with_dollar(vp(a2,"salaires_rd")), "normal"),
        a2row("Crédits d'impôt pour la recherche et le développement", "credits_rd"),
        a2row("Salaires et avantages sociaux - ventes et administration", "salaires_admin"),
        a2row("Frais de déplacement", "deplacement"),
        a2row("Location d'équipement", "location_equip"),
        a2row("Publicité et promotion", "publicite"),
        a2row("Honoraires", "honoraires"),
        a2row("Loyer", "loyer"),
        a2row("Télécommunications", "telecom"),
        a2row("Frais de représentation", "representation"),
        a2row("Frais de bureau", "bureau"),
        a2row("Cotisation et abonnement", "cotisation"),
        a2row("Assurances", "assurance"),
        a2row("Taxes et permis", "taxes"),
        a2row("Entretien et réparations", "entretien"),
        a2row("Courrier et frais postaux", "courrier"),
        a2row("Licence", "licence"),
        a2row("Frais de gestion de paie", "paie"),
        a2row("Frais de formation", "formation"),
        a2row("(Gain) Perte de change", "fx"),
        a2row("Intérêts et pénalités", "interet_penalites"),
        a2row("Représentant externe", "representant"),
        (lbl_para("", bold=True), fmt_with_dollar(v(a2,"total")), fmt_with_dollar(vp(a2,"total")), "grand"),
    ]
    story.append(build_table(ann2_rows))
    story.append(PageBreak())

    # ── ANNEXES 3 & 4 ─────────────────────────────────────────────────────────
    page_hdr("Annexes (suite)", period_str)

    story.append(Paragraph("<b>Annexe 3 - Frais financiers</b>", ps("a3", 10, bold=True)))
    story.append(Spacer(1, 5))
    ann3_rows = [
        (lbl_para("Intérêts sur la dette à long terme"), fmt_with_dollar(v(a3,"interet_lt")), fmt_with_dollar(vp(a3,"interet_lt")), "normal"),
        (lbl_para("Intérêts et frais bancaires"), fmt_num(v(a3,"frais_bancaires")), fmt_num(vp(a3,"frais_bancaires")), "normal"),
        (lbl_para("", bold=True), fmt_with_dollar(v(a3,"total")), fmt_with_dollar(vp(a3,"total")), "grand"),
    ]
    story.append(build_table(ann3_rows))
    story.append(Spacer(1, 20))

    story.append(Paragraph("<b>Annexe 4 - Autres revenus</b>", ps("a4", 10, bold=True)))
    story.append(Spacer(1, 5))
    ann4_rows = [
        (lbl_para("Aide gouvernementale (note 15 b))"), fmt_with_dollar(v(a4,"aide_gouv")), fmt_with_dollar(vp(a4,"aide_gouv")), "normal"),
        (lbl_para("Autres revenus"), fmt_num(v(a4,"autres")), fmt_num(vp(a4,"autres")), "normal"),
        (lbl_para("", bold=True), fmt_with_dollar(v(a4,"total")), fmt_with_dollar(vp(a4,"total")), "grand"),
    ]
    story.append(build_table(ann4_rows))

    # ── BALANCE SHEET ─────────────────────────────────────────────────────────
    bs = data.get("bs", {})
    if bs:
        def bv(key): return (bs.get(key,{}) or {}).get("current",0) or 0
        def bvp(key): return (bs.get(key,{}) or {}).get("prior",0) or 0
        page_hdr("Bilan non consolidé", str(s.get("period_end","")) + ", avec informations comparatives de " + str(data.get("prior_year", s.get("prior_year",""))))
        story.append(PageBreak())
        page_hdr(story, "Bilan non consolidé", str(s.get("period_end","")) + ", avec informations comparatives de " + str(data.get("prior_year", s.get("prior_year",""))))

        bs_rows = [
            # ACTIF
            (lbl_para("Actif"), "", "", "section"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("Actif à court terme", bold=False), "", "", "section"),
            (lbl_para("    Encaisse", indent=20),               fmt_with_dollar(bv("encaisse")),           fmt_with_dollar(bvp("encaisse")),           "normal"),
            (lbl_para("    Comptes clients et autres créances (note 2)", indent=20), fmt_num(bv("comptes_clients")), fmt_num(bvp("comptes_clients")), "normal"),
            (lbl_para("    Stocks", indent=20),                 fmt_num(bv("stocks")),                     fmt_num(bvp("stocks")),                     "normal"),
            (lbl_para("    Travaux en cours (note 15)", indent=20), fmt_num(bv("travaux_en_cours")),       fmt_num(bvp("travaux_en_cours")),           "normal"),
            (lbl_para("    Crédits d'impôt à recevoir", indent=20), fmt_num(bv("credits_impot")),        fmt_num(bvp("credits_impot")),              "normal"),
            (lbl_para("    Charges payées d'avance", indent=20), fmt_num(bv("charges_payees_avance")),   fmt_num(bvp("charges_payees_avance")),      "normal"),
            (lbl_para(""),                                       fmt_num(bv("total_actif_ct")),             fmt_num(bvp("total_actif_ct")),             "subtotal"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("Dépôts à long terme"),                   fmt_num(bv("depots_lt")),                  fmt_num(bvp("depots_lt")),                  "normal"),
            (lbl_para("Frais payés d'avance"),                 fmt_num(bv("frais_payes_avance_lt")),      fmt_num(bvp("frais_payes_avance_lt")),      "normal"),
            (lbl_para("Immobilisations corporelles (note 3)"),  fmt_num(bv("immobilisations")),            fmt_num(bvp("immobilisations")),            "normal"),
            (lbl_para("Actifs incorporels (note 4)"),           fmt_num(bv("actifs_incorporels")),         fmt_num(bvp("actifs_incorporels")),         "normal"),
            (lbl_para("Impôts futurs (note 14)"),               fmt_num(bv("impots_futurs")),              fmt_num(bvp("impots_futurs")),              "normal"),
            (lbl_para("Avances à la filiale, sans intérêt ni modalités de d'encaissement"), fmt_num(bv("avances_filiale")), fmt_num(bvp("avances_filiale")), "normal"),
            (lbl_para("Avances aux actionnaires, sans intérêt\nni modalités de d'encaissement"), fmt_num(bv("avances_actionnaires")), fmt_num(bvp("avances_actionnaires")), "normal"),
            (lbl_para("Placement dans la filiale (note 5)"),    fmt_num(bv("placement_filiale")),          fmt_num(bvp("placement_filiale")),          "normal"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("", bold=True),                           fmt_with_dollar(bv("total_actif")),        fmt_with_dollar(bvp("total_actif")),        "grand"),
            (lbl_para(""), "", "", "blank"),
            # PASSIF
            (lbl_para("Passif et avoir des actionnaires"), "", "", "section"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("Passif à court terme", bold=False), "", "", "section"),
            (lbl_para("    Emprunt bancaire (note 6)", indent=20), fmt_with_dollar(bv("emprunt_bancaire")), fmt_with_dollar(bvp("emprunt_bancaire")), "normal"),
            (lbl_para("    Comptes fournisseurs et charges à payer (note 7)", indent=20), fmt_num(bv("comptes_fournisseurs")), fmt_num(bvp("comptes_fournisseurs")), "normal"),
            (lbl_para("    Impôts sur le bénéfice à payer", indent=20), fmt_num(bv("impots_benefice")),   fmt_num(bvp("impots_benefice")),            "normal"),
            (lbl_para("    Produits reportés (note 15)", indent=20), fmt_num(bv("produits_reportes")),     fmt_num(bvp("produits_reportes")),          "normal"),
            (lbl_para("    Tranche à court terme de la dette à long terme (note 8)", indent=20), fmt_num(bv("tranche_ct_lt")), fmt_num(bvp("tranche_ct_lt")), "normal"),
            (lbl_para(""),                                       fmt_num(bv("total_passif_ct")),            fmt_num(bvp("total_passif_ct")),            "subtotal"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("Avantages incitatifs liés aux baux"),    fmt_num(bv("avantages_baux")),             fmt_num(bvp("avantages_baux")),             "normal"),
            (lbl_para("Dette à long terme (note 8)"),           fmt_num(bv("dette_lt")),                   fmt_num(bvp("dette_lt")),                   "normal"),
            (lbl_para(""),                                       fmt_num(bv("total_passif")),               fmt_num(bvp("total_passif")),               "subtotal"),
            (lbl_para(""), "", "", "blank"),
            # EQUITY
            (lbl_para("Avoir des actionnaires"), "", "", "section"),
            (lbl_para("    Capital-actions (note 9)", indent=20), fmt_num(bv("capital_actions")),          fmt_num(bvp("capital_actions")),            "normal"),
            (lbl_para("    Déficit", indent=20),                fmt_num(bv("deficit")),                    fmt_num(bvp("deficit")),                    "normal"),
            (lbl_para(""),                                       fmt_num(bv("total_avoir")),                fmt_num(bvp("total_avoir")),                "subtotal"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("Engagements (note 13)"), "", "", "normal"),
            (lbl_para(""), "", "", "blank"),
            (lbl_para("", bold=True),                           fmt_with_dollar(bv("total_passif_avoir")), fmt_with_dollar(bvp("total_passif_avoir")), "grand"),
        ]
        story.append(build_table(bs_rows))
        story.append(Spacer(1,10))
        story.append(Paragraph("Se reporter aux notes afférentes aux états financiers.", ps("note",8,color=GREY)))
        story.append(Spacer(1,8))
        story.append(Paragraph("Au nom du conseil,", ps("conseil",9,color=GREY)))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.9*inch, rightMargin=0.9*inch,
        topMargin=0.85*inch, bottomMargin=0.85*inch)
    doc.build(story)
    buf.seek(0)
    return buf

# ── Word builder ──────────────────────────────────────────────────────────────
def build_word(data, s):
    doc = Document()
    for sec in doc.sections:
        sec.page_width=Inches(8.5); sec.page_height=Inches(11)
        sec.left_margin=sec.right_margin=Inches(0.9)
        sec.top_margin=sec.bottom_margin=Inches(0.85)

    NAVY=RGBColor(0x1A,0x2E,0x4A); BLACK=RGBColor(0,0,0); GREY=RGBColor(0x55,0x55,0x55)
    cy=data.get("fiscal_year",s["fiscal_year"]); py=data.get("prior_year",s["prior_year"])
    COL1=Twips(5400); COL2=Twips(1800); COL3=Twips(1800)

    def set_bg(cell,hex_color):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex_color)
        tcPr.append(shd)

    def no_bdr(cell):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        for ex in tcPr.findall(qn('w:tcBorders')): tcPr.remove(ex)
        tcB=OxmlElement('w:tcBorders')
        for side in ['top','left','bottom','right']:
            b=OxmlElement(f'w:{side}'); b.set(qn('w:val'),'none'); tcB.append(b)
        tcPr.append(tcB)

    def underline_cell(cell, thick=False):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        for ex in tcPr.findall(qn('w:tcBorders')): tcPr.remove(ex)
        tcB=OxmlElement('w:tcBorders')
        for side in ['top','left','right']:
            b=OxmlElement(f'w:{side}'); b.set(qn('w:val'),'none'); tcB.append(b)
        b=OxmlElement('w:bottom'); b.set(qn('w:val'),'single')
        b.set(qn('w:sz'),'12' if thick else '6'); b.set(qn('w:color'),'000000')
        tcB.append(b); tcPr.append(tcB)

    def run(para,text,bold=False,size=10,color=BLACK):
        r=para.add_run(text); r.bold=bold; r.font.name="Arial"; r.font.size=Pt(size)
        r.font.color.rgb=color

    def ns(p): p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)

    def page_hdr(title, period):
        p=doc.add_paragraph(); ns(p); run(p,s["company_name"].upper(),bold=True,size=16)
        p2=doc.add_paragraph(); ns(p2); run(p2,title,size=10)
        p3=doc.add_paragraph()
        pPr=p3._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single')
        bot.set(qn('w:sz'),'6'); bot.set(qn('w:color'),'000000')
        pBdr.append(bot); pPr.append(pBdr); p3.paragraph_format.space_after=Pt(8)
        run(p3,period,size=9)

    def make_table(rows):
        tbl=doc.add_table(rows=0,cols=3); tbl.alignment=WD_TABLE_ALIGNMENT.LEFT
        # Header row
        tr=tbl.add_row(); cells=tr.cells
        cells[0].width=COL1; cells[1].width=COL2; cells[2].width=COL3
        for c in cells: set_bg(c,"FFFFFF"); no_bdr(c)
        for c,val,al in [(cells[0],"",WD_ALIGN_PARAGRAPH.LEFT),
                          (cells[1],cy,WD_ALIGN_PARAGRAPH.RIGHT),
                          (cells[2],py,WD_ALIGN_PARAGRAPH.RIGHT)]:
            p=c.paragraphs[0]; p.alignment=al; ns(p)
            run(p,val,bold=True,size=9)
            underline_cell(c)

        for lbl,v1,v2,stype in rows:
            tr=tbl.add_row(); cells=tr.cells
            cells[0].width=COL1; cells[1].width=COL2; cells[2].width=COL3
            for c in cells: set_bg(c,"FFFFFF"); no_bdr(c)

            if stype=="blank":
                for c in cells:
                    p=c.paragraphs[0]; ns(p); p.paragraph_format.space_before=Pt(3)
                continue

            bold_row=stype in ("grand","total","subtotal")
            for c,val,al in [(cells[0],lbl,WD_ALIGN_PARAGRAPH.LEFT),
                              (cells[1],v1,WD_ALIGN_PARAGRAPH.RIGHT),
                              (cells[2],v2,WD_ALIGN_PARAGRAPH.RIGHT)]:
                p=c.paragraphs[0]; p.alignment=al; ns(p)
                p.paragraph_format.space_before=Pt(1); p.paragraph_format.space_after=Pt(1)
                run(p,val,bold=bold_row,size=9.5)

            if stype=="grand":
                for c in cells: underline_cell(c,thick=True)
            elif stype in ("total","subtotal"):
                for c in cells: underline_cell(c)

        doc.add_paragraph().paragraph_format.space_after=Pt(6)

    def v(obj,key):
        d=obj.get(key,{}) if obj else {}
        return (d.get("current",0) or 0) if isinstance(d,dict) else (d or 0)
    def vp(obj,key):
        d=obj.get(key,{}) if obj else {}
        return (d.get("prior",0) or 0) if isinstance(d,dict) else (d or 0)

    pl=data.get("pl",{}); a1=data.get("ann1",{}); a2=data.get("ann2",{})
    a3=data.get("ann3",{}); a4=data.get("ann4",{})
    period_str=f"Exercice clos le {s['period_end']}, avec informations comparatives de {py}"

    page_hdr("État non consolidé des résultats", period_str)
    make_table([
        ("Ventes (notes 11 et 12)", fmt_with_dollar(v(pl,"ventes")), fmt_with_dollar(vp(pl,"ventes")), "normal"),
        ("","","","blank"),
        ("Coût des ventes (annexe 1)", fmt_num(v(pl,"cout_des_ventes")), fmt_num(vp(pl,"cout_des_ventes")), "indent"),
        ("", fmt_num(v(pl,"benefice_brut")), fmt_num(vp(pl,"benefice_brut")), "subtotal"),
        ("","","","blank"),
        ("Charges","","","section"),
        ("    Frais d'exploitation (annexe 2)", fmt_num(v(pl,"frais_exploitation")), fmt_num(vp(pl,"frais_exploitation")), "indent"),
        ("    Frais financiers (annexe 3)", fmt_num(v(pl,"frais_financiers")), fmt_num(vp(pl,"frais_financiers")), "indent"),
        ("    Amortissement des immobilisations corporelles et actifs incorporels",
         fmt_num(v(pl,"amortissement")), fmt_num(vp(pl,"amortissement")), "indent"),
        ("", fmt_num(v(pl,"total_charges")), fmt_num(vp(pl,"total_charges")), "subtotal"),
        ("","","","blank"),
        ("Bénéfice (perte) avant les autres revenus", fmt_num(v(pl,"benefice_avant_autres")), fmt_num(vp(pl,"benefice_avant_autres")), "normal"),
        ("","","","blank"),
        ("Autres revenus (annexe 4)", fmt_num(v(pl,"autres_revenus")), fmt_num(vp(pl,"autres_revenus")), "normal"),
        ("","","","blank"),
        ("Bénéfice net (perte nette)", fmt_with_dollar(v(pl,"benefice_net")), fmt_with_dollar(vp(pl,"benefice_net")), "grand"),
    ])
    p=doc.add_paragraph(); run(p,"Se reporter aux notes afférentes aux états financiers.",size=8,color=GREY)
    doc.add_page_break()

    page_hdr("Annexes", period_str)
    p=doc.add_paragraph(); run(p,"Annexe 1 - Coût des ventes",bold=True,size=10); p.paragraph_format.space_after=Pt(4)
    make_table([
        ("Achats", fmt_with_dollar(v(a1,"achats")), fmt_with_dollar(vp(a1,"achats")), "normal"),
        ("Salaires et avantages sociaux", fmt_num(v(a1,"salaires")), fmt_num(vp(a1,"salaires")), "normal"),
        ("Sous-traitance", fmt_num(v(a1,"soustraitance")), fmt_num(vp(a1,"soustraitance")), "normal"),
        ("Frais de livraison", fmt_num(v(a1,"livraison")), fmt_num(vp(a1,"livraison")), "normal"),
        ("Logiciel", fmt_num(v(a1,"logiciel")), fmt_num(vp(a1,"logiciel")), "normal"),
        ("", fmt_with_dollar(v(a1,"total")), fmt_with_dollar(vp(a1,"total")), "grand"),
    ])
    p=doc.add_paragraph(); run(p,"Annexe 2 - Frais d'exploitation",bold=True,size=10); p.paragraph_format.space_after=Pt(4)
    make_table([
        ("Salaires et avantages sociaux - recherche et développement", fmt_with_dollar(v(a2,"salaires_rd")), fmt_with_dollar(vp(a2,"salaires_rd")), "normal"),
        ("Crédits d'impôt pour la recherche et le développement", fmt_num(v(a2,"credits_rd")), fmt_num(vp(a2,"credits_rd")), "normal"),
        ("Salaires et avantages sociaux - ventes et administration", fmt_num(v(a2,"salaires_admin")), fmt_num(vp(a2,"salaires_admin")), "normal"),
        ("Frais de déplacement", fmt_num(v(a2,"deplacement")), fmt_num(vp(a2,"deplacement")), "normal"),
        ("Location d'équipement", fmt_num(v(a2,"location_equip")), fmt_num(vp(a2,"location_equip")), "normal"),
        ("Publicité et promotion", fmt_num(v(a2,"publicite")), fmt_num(vp(a2,"publicite")), "normal"),
        ("Honoraires", fmt_num(v(a2,"honoraires")), fmt_num(vp(a2,"honoraires")), "normal"),
        ("Loyer", fmt_num(v(a2,"loyer")), fmt_num(vp(a2,"loyer")), "normal"),
        ("Télécommunications", fmt_num(v(a2,"telecom")), fmt_num(vp(a2,"telecom")), "normal"),
        ("Frais de représentation", fmt_num(v(a2,"representation")), fmt_num(vp(a2,"representation")), "normal"),
        ("Frais de bureau", fmt_num(v(a2,"bureau")), fmt_num(vp(a2,"bureau")), "normal"),
        ("Cotisation et abonnement", fmt_num(v(a2,"cotisation")), fmt_num(vp(a2,"cotisation")), "normal"),
        ("Assurances", fmt_num(v(a2,"assurance")), fmt_num(vp(a2,"assurance")), "normal"),
        ("Taxes et permis", fmt_num(v(a2,"taxes")), fmt_num(vp(a2,"taxes")), "normal"),
        ("Entretien et réparations", fmt_num(v(a2,"entretien")), fmt_num(vp(a2,"entretien")), "normal"),
        ("Courrier et frais postaux", fmt_num(v(a2,"courrier")), fmt_num(vp(a2,"courrier")), "normal"),
        ("Dépense de l'avantage incitatifs liés aux baux", fmt_num(v(a2,"lease_incentive")), fmt_num(vp(a2,"lease_incentive")), "normal"),
        ("Licence", fmt_num(v(a2,"licence")), fmt_num(vp(a2,"licence")), "normal"),
        ("Frais de gestion de paie", fmt_num(v(a2,"paie")), fmt_num(vp(a2,"paie")), "normal"),
        ("Frais de formation", fmt_num(v(a2,"formation")), fmt_num(vp(a2,"formation")), "normal"),
        ("(Gain) Perte de change", fmt_num(v(a2,"fx")), fmt_num(vp(a2,"fx")), "normal"),
        ("Intérêts et pénalités", fmt_num(v(a2,"interet_penalites")), fmt_num(vp(a2,"interet_penalites")), "normal"),
        ("Représentant externe", fmt_num(v(a2,"representant")), fmt_num(vp(a2,"representant")), "normal"),
        ("", fmt_with_dollar(v(a2,"total")), fmt_with_dollar(vp(a2,"total")), "grand"),
    ])
    doc.add_page_break()

    page_hdr("Annexes (suite)", period_str)
    p=doc.add_paragraph(); run(p,"Annexe 3 - Frais financiers",bold=True,size=10); p.paragraph_format.space_after=Pt(4)
    make_table([
        ("Intérêts sur la dette à long terme", fmt_with_dollar(v(a3,"interet_lt")), fmt_with_dollar(vp(a3,"interet_lt")), "normal"),
        ("Intérêts et frais bancaires", fmt_num(v(a3,"frais_bancaires")), fmt_num(vp(a3,"frais_bancaires")), "normal"),
        ("", fmt_with_dollar(v(a3,"total")), fmt_with_dollar(vp(a3,"total")), "grand"),
    ])
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(14)
    run(p,"Annexe 4 - Autres revenus",bold=True,size=10); p.paragraph_format.space_after=Pt(4)
    make_table([
        ("Aide gouvernementale (note 15 b))", fmt_with_dollar(v(a4,"aide_gouv")), fmt_with_dollar(vp(a4,"aide_gouv")), "normal"),
        ("Autres revenus", fmt_num(v(a4,"autres")), fmt_num(vp(a4,"autres")), "normal"),
        ("", fmt_with_dollar(v(a4,"total")), fmt_with_dollar(vp(a4,"total")), "grand"),
    ])

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# ── Main UI ───────────────────────────────────────────────────────────────────
st.markdown('<div class="info-box">Upload your QuickBooks P&L PDF or Excel export. Include prior year PDF alongside current year for the two-column comparison.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader("Drop files here", type=["pdf","xlsx"], accept_multiple_files=True, label_visibility="collapsed")
if uploaded:
    st.markdown(f"**{len(uploaded)} file(s) ready:** " + ", ".join(f.name for f in uploaded))

if not get_api_key():
    st.markdown('<div class="info-box">👈 Enter your Anthropic API key in the sidebar.</div>', unsafe_allow_html=True)

if st.button("⚡ Extract & Build Report", disabled=not(uploaded and get_api_key()), type="primary", use_container_width=True):
    with st.spinner("Reading files and grouping accounts... ~30 seconds"):
        try:
            data = extract(uploaded)
            st.session_state["data_v3"] = data
        except Exception as e:
            st.error(f"Extraction failed: {e}")

if "data_v3" in st.session_state:
    data = st.session_state["data_v3"]
    pl = data.get("pl",{})
    cy = data.get("fiscal_year", fiscal_year)
    py_val = data.get("prior_year", prior_year)

    st.success(f"✅ {data.get('company','')} · FY{cy} vs FY{py_val}")
    st.markdown("---")

    # KPIs
    def gv(key): return (pl.get(key,{}) or {}).get("current",0) or 0
    rev=gv("ventes"); gp=gv("benefice_brut"); np_=gv("benefice_net")
    c1,c2,c3,c4=st.columns(4)
    with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-label">Revenue</div><div class="kpi-value">{fmt_num(rev)}</div></div>',unsafe_allow_html=True)
    with c2:
        cls="kpi-pos" if gp>=0 else "kpi-neg"
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Gross profit</div><div class="kpi-value {cls}">{fmt_num(gp)}</div><div class="kpi-sub">Margin {pct(gp,rev)}</div></div>',unsafe_allow_html=True)
    with c3:
        cls="kpi-pos" if np_>=0 else "kpi-neg"
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Net profit</div><div class="kpi-value {cls}">{fmt_num(np_)}</div><div class="kpi-sub">Margin {pct(np_,rev)}</div></div>',unsafe_allow_html=True)
    with c4:
        bs = data.get("bs", {})
        ta = (bs.get("total_actif",{}) or {}).get("current",0) if bs else 0
        st.markdown(f'<div class="kpi-card"><div class="kpi-label">Total actif</div><div class="kpi-value">{fmt_num(ta)}</div></div>',unsafe_allow_html=True)

    st.markdown("---")
    col1,col2=st.columns(2)
    slug=f"FY{cy}"

    with col1:
        pdf_buf=build_pdf(data,settings)
        st.download_button("📄 Download PDF (audited style)",data=pdf_buf,
            file_name=f"ActiveMedia_FinancialStatements_{slug}.pdf",
            mime="application/pdf",use_container_width=True,type="primary")
    with col2:
        word_buf=build_word(data,settings)
        st.download_button("📝 Download Word (.docx)",data=word_buf,
            file_name=f"ActiveMedia_FinancialStatements_{slug}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True)
