import streamlit as st
import anthropic
import base64
import json
import io
from datetime import datetime
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
<p>Active Média inc. — SmartPixel &nbsp;·&nbsp; Upload QuickBooks PDFs to generate a complete audited-style report</p></div>""",unsafe_allow_html=True)

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
    fiscal_year  = st.text_input("Current fiscal year",value="2025")
    prior_year   = st.text_input("Prior fiscal year",value="2024")
    period_end   = st.text_input("Period end date",value="January 31, 2025")
    preparer     = st.text_input("Prepared by",value="Management")
    currency     = st.selectbox("Currency",["CAD","USD","EUR"],index=0)
    st.markdown("---")
    st.markdown("**Report style**")
    mode = st.radio("View",["Consolidated (clean, like audited report)","Detailed (all accounts)"],index=0)
    consolidated = mode.startswith("Consolidated")
    st.markdown("---")
    st.markdown("**Include**")
    inc_cover   = st.checkbox("Cover page",value=True)
    inc_pl      = st.checkbox("Income statement",value=True)
    inc_bs      = st.checkbox("Balance sheet",value=True)
    inc_cf      = st.checkbox("Cash flow",value=True)
    inc_annexes = st.checkbox("Annexes",value=True)
    inc_notes   = st.checkbox("Notes",value=True)
    st.caption("SmartPixel Financial Tool · v2.1")

settings = dict(company_name=company_name,fiscal_year=fiscal_year,prior_year=prior_year,
                period_end=period_end,preparer=preparer,currency=currency,
                consolidated=consolidated,inc_cover=inc_cover,inc_pl=inc_pl,
                inc_bs=inc_bs,inc_cf=inc_cf,inc_annexes=inc_annexes,inc_notes=inc_notes)

def fmt(n,short=False):
    if n is None: return "—"
    try: n=float(n)
    except: return "—"
    if short:
        if abs(n)>=1_000_000: return f"{'(' if n<0 else ''}${abs(n)/1_000_000:.1f}M{')' if n<0 else ''}"
        if abs(n)>=1_000: return f"{'(' if n<0 else ''}${abs(n)/1_000:.0f}K{')' if n<0 else ''}"
    s=f"{abs(n):,.0f}"
    return f"({s})" if n<0 else s

def pct(a,b):
    try: return f"{float(a)/float(b)*100:.1f}%"
    except: return "—"

PROMPT_CONSOLIDATED = """Extract financial data from these PDFs for Active Média inc. / SmartPixel.

CONSOLIDATED MODE: Group into clean high-level line items only (5-8 lines per section max), like a real audited annual report. Do NOT list individual QuickBooks accounts. Group all salary accounts together, all software together, etc.

Return ONLY valid JSON:
{
  "company":"string","fiscal_year":"string","prior_year":"string","period_end":"string","currency":"CAD",
  "pl":{"available":true,"period_label":"string",
    "revenue":[{"label":"string","current":number,"prior":number}],
    "total_revenue":{"current":number,"prior":number},
    "cogs":[{"label":"string","current":number,"prior":number}],
    "total_cogs":{"current":number,"prior":number},
    "gross_profit":{"current":number,"prior":number},
    "operating_expenses":[{"label":"string","current":number,"prior":number}],
    "total_opex":{"current":number,"prior":number},
    "financial_expenses":[{"label":"string","current":number,"prior":number}],
    "total_financial":{"current":number,"prior":number},
    "depreciation":{"current":number,"prior":number},
    "profit_before_other":{"current":number,"prior":number},
    "other_income":[{"label":"string","current":number,"prior":number}],
    "total_other":{"current":number,"prior":number},
    "net_profit":{"current":number,"prior":number},
    "annexes":{"cogs_detail":[{"label":"string","current":number,"prior":number}],
               "opex_detail":[{"label":"string","current":number,"prior":number}],
               "financial_detail":[{"label":"string","current":number,"prior":number}],
               "other_detail":[{"label":"string","current":number,"prior":number}]}},
  "bs":{"available":true,"period_label":"string",
    "current_assets":[{"label":"string","current":number,"prior":number}],
    "total_current_assets":{"current":number,"prior":number},
    "longterm_assets":[{"label":"string","current":number,"prior":number}],
    "total_longterm_assets":{"current":number,"prior":number},
    "total_assets":{"current":number,"prior":number},
    "current_liabilities":[{"label":"string","current":number,"prior":number}],
    "total_current_liabilities":{"current":number,"prior":number},
    "longterm_liabilities":[{"label":"string","current":number,"prior":number}],
    "total_longterm_liabilities":{"current":number,"prior":number},
    "total_liabilities":{"current":number,"prior":number},
    "equity":[{"label":"string","current":number,"prior":number}],
    "total_equity":{"current":number,"prior":number},
    "total_liabilities_equity":{"current":number,"prior":number}},
  "cf":{"available":false}
}"""

PROMPT_DETAILED = PROMPT_CONSOLIDATED.replace(
    "CONSOLIDATED MODE: Group into clean high-level line items only (5-8 lines per section max), like a real audited annual report. Do NOT list individual QuickBooks accounts. Group all salary accounts together, all software together, etc.",
    "DETAILED MODE: Include all individual account lines with their account codes exactly as they appear."
)

def extract(pdf_files, consolidated):
    client = anthropic.Anthropic(api_key=get_api_key())
    content = []
    for f in pdf_files:
        b64 = base64.b64encode(f.read()).decode()
        content.append({"type":"document","source":{"type":"base64","media_type":"application/pdf","data":b64},"title":f.name})
    content.append({"type":"text","text": PROMPT_CONSOLIDATED if consolidated else PROMPT_DETAILED})
    msg = client.messages.create(model="claude-sonnet-4-6",max_tokens=16000,messages=[{"role":"user","content":content}])
    raw = msg.content[0].text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        last = raw.rfind("}")
        if last>0:
            fixed=raw[:last+1]
            open_count=fixed.count("{")-fixed.count("}")
            return json.loads(fixed+"}"*open_count)
        raise

# ── Word document builder ─────────────────────────────────────────────────────
def build_word(data, s):
    doc = Document()
    # Page setup — Letter, narrow margins to maximize table width
    for sec in doc.sections:
        sec.page_width  = Inches(8.5)
        sec.page_height = Inches(11)
        sec.left_margin = sec.right_margin = Inches(0.9)
        sec.top_margin  = sec.bottom_margin = Inches(0.85)

    NAVY  = RGBColor(0x1A,0x2E,0x4A)
    BLUE  = RGBColor(0x2E,0x75,0xB6)
    WHITE = RGBColor(0xFF,0xFF,0xFF)
    GREY  = RGBColor(0x55,0x55,0x55)
    cy = data.get("fiscal_year", s["fiscal_year"])
    py = data.get("prior_year",  s["prior_year"])
    cur = s["currency"]

    # content width in twips: (8.5 - 1.8) * 1440 = 9648
    CONTENT_W = 9648
    COL1 = 6048  # label column
    COL2 = 1800  # current year
    COL3 = 1800  # prior year

    def set_bg(cell, hex_color):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex_color)
        tcPr.append(shd)

    def set_cell_borders(cell, top=None, bottom=None, color="000000", size="6"):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        for ex in tcPr.findall(qn('w:tcBorders')): tcPr.remove(ex)
        tcB=OxmlElement('w:tcBorders')
        for side in ['top','left','bottom','right']:
            b=OxmlElement(f'w:{side}')
            val = top if side=='top' else bottom if side=='bottom' else None
            if val:
                b.set(qn('w:val'),val); b.set(qn('w:sz'),size); b.set(qn('w:color'),color)
            else:
                b.set(qn('w:val'),'none')
            tcB.append(b)
        tcPr.append(tcB)

    def no_space(para):
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after  = Pt(0)

    def run(para, text, bold=False, italic=False, size=10, color=None):
        r = para.add_run(text)
        r.bold=bold; r.italic=italic; r.font.name="Arial"; r.font.size=Pt(size)
        if color: r.font.color.rgb=color

    def page_break(): doc.add_page_break()

    def company_header(subtitle, period):
        # Bold company name — large, matches reference
        p = doc.add_paragraph(); no_space(p); p.paragraph_format.space_before=Pt(0)
        run(p, s["company_name"].upper(), bold=True, size=18, color=RGBColor(0,0,0))
        # Subtitle in normal weight
        p2 = doc.add_paragraph(); no_space(p2); p2.paragraph_format.space_after=Pt(0)
        run(p2, subtitle, size=10, color=RGBColor(0,0,0))
        # Period line with thick underline beneath
        p3 = doc.add_paragraph()
        pPr=p3._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single')
        bot.set(qn('w:sz'),'8'); bot.set(qn('w:color'),'000000')
        pBdr.append(bot); pPr.append(pBdr)
        p3.paragraph_format.space_after=Pt(10)
        run(p3, period, size=9, color=RGBColor(0,0,0))

    def make_fin_table(rows_data):
        tbl = doc.add_table(rows=0, cols=3)
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.style = 'Table Normal'

        zebra = False
        for rd in rows_data:
            sty = rd.get("s","data"); lbl = rd.get("l","")
            v1  = rd.get("v1","");   v2  = rd.get("v2","")

            tr = tbl.add_row()
            cells = tr.cells

            # Set column widths
            cells[0].width = Twips(COL1)
            cells[1].width = Twips(COL2)
            cells[2].width = Twips(COL3)

            if sty == "blank":
                for c in cells:
                    set_bg(c,"FFFFFF"); set_cell_borders(c)
                    no_space(c.paragraphs[0])
                    c.paragraphs[0].paragraph_format.space_before=Pt(2)
                    c.paragraphs[0].paragraph_format.space_after=Pt(2)
                continue

            # Background colors
            if sty == "section":
                bg="FFFFFF"; fc=NAVY_hex="1A2E4A"; text_color=NAVY; bold=True
            elif sty == "total":
                bg="FFFFFF"; text_color=BLACK=RGBColor(0,0,0); bold=True
            elif sty == "grand":
                bg="FFFFFF"; text_color=NAVY; bold=True
            elif sty == "pbt":  # profit before tax line
                bg="FFFFFF"; text_color=BLACK=RGBColor(0,0,0); bold=False
            else:
                bg="FFFFFF"; text_color=RGBColor(0,0,0); bold=False

            for c in cells: set_bg(c, bg)

            # Cell 0 — label
            p = cells[0].paragraphs[0]; no_space(p)
            p.paragraph_format.space_before=Pt(1); p.paragraph_format.space_after=Pt(1)
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            indent = "    " if sty=="data" else ""
            run(p, indent+lbl, bold=bold, size=9.5,
                color=NAVY if sty in ("section","grand") else RGBColor(0,0,0))

            # Cell 1 & 2 — values
            for c, val in [(cells[1],v1),(cells[2],v2)]:
                p2=c.paragraphs[0]; no_space(p2)
                p2.paragraph_format.space_before=Pt(1); p2.paragraph_format.space_after=Pt(1)
                p2.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                run(p2, str(val) if val else "—", bold=bold, size=9.5,
                    color=NAVY if sty in ("section","grand") else RGBColor(0,0,0))

            # Borders matching reference style
            if sty == "grand":
                for c in cells:
                    set_cell_borders(c, top="single", bottom="single", color="000000", size="8")
            elif sty == "total":
                for c in cells:
                    set_cell_borders(c, bottom="single", color="000000", size="6")
            elif sty == "section":
                for c in cells:
                    set_cell_borders(c, bottom="single", color="000000", size="6")
            else:
                for c in cells: set_cell_borders(c)

        doc.add_paragraph().paragraph_format.space_after=Pt(4)
        return tbl

    def tv(obj,key):
        d=obj.get(key,{}) if obj else {}
        return d.get("current"), d.get("prior")

    pl = data.get("pl",{}); bs = data.get("bs",{}); cf = data.get("cf",{})

    # ── COVER ──────────────────────────────────────────────────────────────────
    if s["inc_cover"]:
        doc.add_paragraph().paragraph_format.space_before=Pt(60)
        p=doc.add_paragraph(); run(p, s["company_name"], bold=True, size=24, color=NAVY)
        p2=doc.add_paragraph()
        pPr=p2._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single')
        bot.set(qn('w:sz'),'12'); bot.set(qn('w:color'),'2E75B6')
        pBdr.append(bot); pPr.append(pBdr); p2.paragraph_format.space_after=Pt(20)
        p3=doc.add_paragraph(); run(p3,"ÉTATS FINANCIERS / FINANCIAL STATEMENTS",bold=True,size=16,color=NAVY)
        for line in [
            f"Exercice clos le / Year ended: {s['period_end']}",
            f"Exercice comparatif / Comparative year: {py}",
            f"Monnaie / Currency: {cur}",
            f"Préparé par / Prepared by: {s['preparer']}",
            f"Date: {datetime.now().strftime('%B %d, %Y')}",
            "Confidentiel — Usage interne / Confidential — Internal use only"
        ]:
            px=doc.add_paragraph(); run(px,line,size=10,color=GREY,italic=line.startswith("Conf"))
        page_break()

    # ── P&L ────────────────────────────────────────────────────────────────────
    if s["inc_pl"] and pl.get("available"):
        company_header(
            "État des résultats / Income Statement",
            f"Exercice clos le {s['period_end']}, avec informations comparatives de {py}"
        )
        rows=[]
        # Header row
        rows.append({"s":"section","l":"","v1":cy,"v2":py})

        # Revenue
        for x in pl.get("revenue",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        c,p_=tv(pl,"total_revenue"); rows.append({"s":"total","l":"Ventes / Revenue","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        # COGS
        c,p_=tv(pl,"total_cogs"); rows.append({"s":"data","l":"Coût des ventes / Cost of sales (Annexe 1)","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        # Gross profit
        c,p_=tv(pl,"gross_profit")
        tr=tv(pl,"total_revenue"); 
        rows.append({"s":"total","l":"Bénéfice brut / Gross profit","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        # Expenses
        rows.append({"s":"data","l":"Charges / Expenses","v1":"","v2":""})
        c,p_=tv(pl,"total_opex"); rows.append({"s":"data","l":"    Frais d'exploitation / Operating expenses (Annexe 2)","v1":fmt(c),"v2":fmt(p_)})
        c,p_=tv(pl,"total_financial"); rows.append({"s":"data","l":"    Frais financiers / Financial expenses (Annexe 3)","v1":fmt(c),"v2":fmt(p_)})
        dep_c=pl.get("depreciation",{}).get("current"); dep_p=pl.get("depreciation",{}).get("prior")
        if dep_c: rows.append({"s":"data","l":"    Amortissement / Depreciation","v1":fmt(dep_c),"v2":fmt(dep_p)})
        tot_exp_c = (pl.get("total_opex",{}).get("current") or 0) + (pl.get("total_financial",{}).get("current") or 0) + (dep_c or 0)
        tot_exp_p = (pl.get("total_opex",{}).get("prior") or 0) + (pl.get("total_financial",{}).get("prior") or 0) + (dep_p or 0)
        rows.append({"s":"total","l":"Total charges","v1":fmt(tot_exp_c),"v2":fmt(tot_exp_p)})
        rows.append({"s":"blank"})

        # Profit before other
        c,p_=tv(pl,"profit_before_other"); rows.append({"s":"pbt","l":"Bénéfice (perte) avant autres revenus","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        # Other income
        if pl.get("other_income"):
            c,p_=tv(pl,"total_other"); rows.append({"s":"data","l":"Autres revenus / Other income (Annexe 4)","v1":fmt(c),"v2":fmt(p_)})
            rows.append({"s":"blank"})

        # Net profit
        c,p_=tv(pl,"net_profit"); rows.append({"s":"grand","l":"Bénéfice net (perte nette) / Net profit (loss)","v1":fmt(c),"v2":fmt(p_)})
        make_fin_table(rows)
        p=doc.add_paragraph(); run(p,"Se reporter aux notes afférentes aux états financiers.",italic=True,size=8,color=GREY)
        page_break()

    # ── BALANCE SHEET ──────────────────────────────────────────────────────────
    if s["inc_bs"] and bs.get("available"):
        company_header("Bilan / Balance Sheet", f"Au {s['period_end']}, avec informations comparatives de {py}")
        rows=[]
        rows.append({"s":"section","l":"","v1":cy,"v2":py})

        rows.append({"s":"section","l":"Actif à court terme / Current assets","v1":"","v2":""})
        for x in bs.get("current_assets",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        c,p_=tv(bs,"total_current_assets"); rows.append({"s":"total","l":"Total actif à court terme","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        for x in bs.get("longterm_assets",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        rows.append({"s":"blank"})
        c,p_=tv(bs,"total_assets"); rows.append({"s":"grand","l":"Total actif / Total assets","v1":fmt(c)+"  $","v2":fmt(p_)+"  $"})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Passif à court terme / Current liabilities","v1":"","v2":""})
        for x in bs.get("current_liabilities",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        c,p_=tv(bs,"total_current_liabilities"); rows.append({"s":"total","l":"Total passif à court terme","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        for x in bs.get("longterm_liabilities",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        rows.append({"s":"blank"})
        c,p_=tv(bs,"total_liabilities"); rows.append({"s":"total","l":"Total passif / Total liabilities","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Avoir des actionnaires / Shareholders' equity","v1":"","v2":""})
        for x in bs.get("equity",[]): rows.append({"s":"data","l":x["label"],"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
        c,p_=tv(bs,"total_equity"); rows.append({"s":"total","l":"Total avoir des actionnaires","v1":fmt(c),"v2":fmt(p_)})
        rows.append({"s":"blank"})
        c,p_=tv(bs,"total_liabilities_equity"); rows.append({"s":"grand","l":"Total passif et avoir / Total L&E","v1":fmt(c)+"  $","v2":fmt(p_)+"  $"})
        make_fin_table(rows)
        p=doc.add_paragraph(); run(p,"Se reporter aux notes afférentes aux états financiers.",italic=True,size=8,color=GREY)
        page_break()

    # ── ANNEXES ────────────────────────────────────────────────────────────────
    if s["inc_annexes"] and pl.get("available"):
        ann = pl.get("annexes",{})
        company_header("Annexes / Schedules", f"Exercice clos le {s['period_end']}")
        for num,title_fr,title_en,key in [
            ("1","Coût des ventes","Cost of sales","cogs_detail"),
            ("2","Frais d'exploitation","Operating expenses","opex_detail"),
            ("3","Frais financiers","Financial expenses","financial_detail"),
            ("4","Autres revenus","Other income","other_detail"),
        ]:
            items = ann.get(key,[])
            if not items: continue
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4)
            run(p,f"Annexe {num} — {title_fr} / {title_en}",bold=True,size=10,color=NAVY)
            rows=[]
            rows.append({"s":"section","l":"","v1":cy,"v2":py})
            for x in items: rows.append({"s":"data","l":x.get("label",""),"v1":fmt(x.get("current")),"v2":fmt(x.get("prior"))})
            tot_c=sum(float(x.get("current") or 0) for x in items)
            tot_p=sum(float(x.get("prior") or 0) for x in items)
            rows.append({"s":"grand","l":f"Total — {title_en}","v1":fmt(tot_c)+"  $","v2":fmt(tot_p)+"  $"})
            make_fin_table(rows)

    # ── NOTES ──────────────────────────────────────────────────────────────────
    if s["inc_notes"]:
        company_header("Notes afférentes aux états financiers / Notes to Financial Statements",
                       f"Exercice clos le {s['period_end']}")
        for title,body in [
            ("Note 1 — Principales méthodes comptables / Significant accounting policies",
             f"Ces états financiers ont été préparés conformément aux Normes comptables canadiennes pour les entreprises à capital fermé (NCECF). / These financial statements have been prepared in accordance with Canadian Accounting Standards for Private Enterprises (ASPE)."),
            ("Note 2 — Base de présentation / Basis of presentation",
             f"Les états financiers sont présentés en dollars canadiens (CAD). L'exercice financier se termine le {s['period_end']}. / Financial statements are presented in Canadian dollars (CAD). The fiscal year ends {s['period_end']}."),
            ("Note 3 — Monnaie étrangère / Foreign currency",
             "Les opérations en devises étrangères sont converties au taux de change en vigueur à la date de la transaction. / Foreign currency transactions are translated at the exchange rate in effect at the transaction date."),
        ]:
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(3)
            run(p,title,bold=True,size=10,color=NAVY)
            p2=doc.add_paragraph(); run(p2,body,size=9,color=GREY)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# ── Excel builder (unchanged logic, cleaner) ──────────────────────────────────
def build_excel(data, s):
    wb=Workbook()
    NAVY,ACCENT,LIGHT,GREY,WHITE="1A2E4A","2E75B6","D9E8F5","F5F5F5","FFFFFF"
    cy=data.get("fiscal_year",s["fiscal_year"]); py=data.get("prior_year",s["prior_year"])
    def fl(c): return PatternFill("solid",fgColor=c)
    def fnt(bold=False,color="000000",size=10): return Font(name="Arial",bold=bold,color=color,size=size)
    def bdr():
        s=Side(style="thin",color="CCCCCC"); return Border(left=s,right=s,top=s,bottom=s)
    def bdr_m():
        t=Side(style="thin",color="CCCCCC"); m=Side(style="medium",color=ACCENT)
        return Border(left=t,right=t,top=t,bottom=m)
    def aln(h="left"): return Alignment(horizontal=h,vertical="center")
    FMT='#,##0;(#,##0);"-"'

    def add_sheet(title,rows_data):
        ws=wb.create_sheet(title); ws.sheet_view.showGridLines=False; ws.freeze_panes="B3"
        ws["A1"].value=f"{title} — {s['company_name']}"; ws["A1"].font=fnt(bold=True,color=NAVY,size=12)
        ws.merge_cells("A1:C1")
        ws.column_dimensions["A"].width=46; ws.column_dimensions["B"].width=18; ws.column_dimensions["C"].width=18
        for i,(lbl,col) in enumerate([("Account",1),(f"FY{cy} ({s['currency']} $)",2),(f"FY{py} ({s['currency']} $)",3)],1):
            c=ws.cell(row=2,column=i,value=lbl)
            c.font=fnt(bold=True,color=WHITE); c.fill=fl(NAVY); c.border=bdr(); c.alignment=aln("center" if i>1 else "left")
        r=3; zebra=False
        for rd in rows_data:
            st2=rd.get("s","data"); lbl=rd.get("l",""); v1=rd.get("v1"); v2=rd.get("v2")
            if st2=="blank":
                for col in range(1,4):
                    c=ws.cell(row=r,column=col); c.fill=fl(WHITE)
                    sv=Side(style="thin",color="EEEEEE"); c.border=Border(left=sv,right=sv,top=sv,bottom=sv)
                r+=1; continue
            bg=NAVY if st2=="grand" else LIGHT if st2=="total" else ACCENT if st2=="section" else GREY if zebra else WHITE
            fc=WHITE if st2 in ("grand","section") else "000000"
            bold=st2 in ("grand","total","section")
            if st2=="data": zebra=not zebra
            else: zebra=False
            c=ws.cell(row=r,column=1,value=lbl)
            c.font=fnt(bold=bold,color=fc,size=9); c.fill=fl(bg)
            c.border=bdr_m() if st2 in ("grand","total") else bdr(); c.alignment=aln()
            for col,v in [(2,v1),(3,v2)]:
                fc2=ws.cell(row=r,column=col,value=v)
                fc2.font=fnt(bold=bold,color=fc,size=9); fc2.fill=fl(bg)
                fc2.number_format=FMT; fc2.border=bdr_m() if st2 in ("grand","total") else bdr(); fc2.alignment=aln("right")
            r+=1

    def tv(obj,key):
        d=obj.get(key,{}) if obj else {}; return d.get("current"),d.get("prior")

    ws=wb.active; ws.title="Cover"
    ws["A1"].value=s["company_name"]; ws["A1"].font=fnt(bold=True,color=NAVY,size=14)
    ws["A2"].value=f"Financial Statements FY{cy} vs FY{py}"
    ws["A3"].value=f"Period end: {s['period_end']}"
    ws.column_dimensions["A"].width=40

    pl=data.get("pl",{}); bs=data.get("bs",{}); cf=data.get("cf",{})

    if pl.get("available"):
        rows=[]
        for x in pl.get("revenue",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        c,p_=tv(pl,"total_revenue"); rows.append({"s":"total","l":"Total revenue","v1":c,"v2":p_}); rows.append({"s":"blank"})
        c,p_=tv(pl,"total_cogs"); rows.append({"s":"data","l":"Cost of sales","v1":c,"v2":p_}); rows.append({"s":"blank"})
        c,p_=tv(pl,"gross_profit"); rows.append({"s":"grand","l":"Gross profit","v1":c,"v2":p_}); rows.append({"s":"blank"})
        c,p_=tv(pl,"total_opex"); rows.append({"s":"data","l":"Operating expenses","v1":c,"v2":p_})
        c,p_=tv(pl,"total_financial"); rows.append({"s":"data","l":"Financial expenses","v1":c,"v2":p_})
        dep_c=pl.get("depreciation",{}).get("current"); dep_p=pl.get("depreciation",{}).get("prior")
        if dep_c: rows.append({"s":"data","l":"Depreciation","v1":dep_c,"v2":dep_p})
        c,p_=tv(pl,"profit_before_other"); rows.append({"s":"total","l":"Profit before other income","v1":c,"v2":p_}); rows.append({"s":"blank"})
        if pl.get("other_income"):
            c,p_=tv(pl,"total_other"); rows.append({"s":"data","l":"Other income","v1":c,"v2":p_}); rows.append({"s":"blank"})
        c,p_=tv(pl,"net_profit"); rows.append({"s":"grand","l":"Net profit / (loss)","v1":c,"v2":p_})
        add_sheet("P&L",rows)
        ann=pl.get("annexes",{})
        ann_rows=[]
        for num,title,key in [("1","Cost of sales","cogs_detail"),("2","Operating expenses","opex_detail"),("3","Financial expenses","financial_detail"),("4","Other income","other_detail")]:
            items=ann.get(key,[])
            if not items: continue
            ann_rows.append({"s":"section","l":f"Annex {num} — {title}","v1":None,"v2":None})
            for x in items: ann_rows.append({"s":"data","l":x.get("label",""),"v1":x.get("current"),"v2":x.get("prior")})
            ann_rows.append({"s":"total","l":f"Total","v1":sum(float(x.get("current") or 0) for x in items),"v2":sum(float(x.get("prior") or 0) for x in items)})
            ann_rows.append({"s":"blank"})
        if ann_rows: add_sheet("Annexes",ann_rows)

    if bs.get("available"):
        rows=[]
        rows.append({"s":"section","l":"Current assets","v1":None,"v2":None})
        for x in bs.get("current_assets",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        c,p_=tv(bs,"total_current_assets"); rows.append({"s":"total","l":"Total current assets","v1":c,"v2":p_}); rows.append({"s":"blank"})
        for x in bs.get("longterm_assets",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        rows.append({"s":"blank"})
        c,p_=tv(bs,"total_assets"); rows.append({"s":"grand","l":"TOTAL ASSETS","v1":c,"v2":p_}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Current liabilities","v1":None,"v2":None})
        for x in bs.get("current_liabilities",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        c,p_=tv(bs,"total_current_liabilities"); rows.append({"s":"total","l":"Total current liabilities","v1":c,"v2":p_}); rows.append({"s":"blank"})
        for x in bs.get("longterm_liabilities",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        rows.append({"s":"blank"})
        c,p_=tv(bs,"total_liabilities"); rows.append({"s":"total","l":"Total liabilities","v1":c,"v2":p_}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Shareholders' equity","v1":None,"v2":None})
        for x in bs.get("equity",[]): rows.append({"s":"data","l":x["label"],"v1":x.get("current"),"v2":x.get("prior")})
        c,p_=tv(bs,"total_equity"); rows.append({"s":"total","l":"Total equity","v1":c,"v2":p_}); rows.append({"s":"blank"})
        c,p_=tv(bs,"total_liabilities_equity"); rows.append({"s":"grand","l":"TOTAL LIABILITIES & EQUITY","v1":c,"v2":p_})
        add_sheet("Balance Sheet",rows)

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── Main UI ───────────────────────────────────────────────────────────────────
st.markdown('<div style="font-size:.7rem;font-weight:600;text-transform:uppercase;letter-spacing:.08em;color:#2E75B6;margin-bottom:.4rem">Step 1 — Upload QuickBooks PDFs</div>',unsafe_allow_html=True)
st.markdown('<div class="info-box">Upload current year + prior year PDFs together — Claude extracts both and builds the two-column comparison automatically.</div>',unsafe_allow_html=True)

uploaded = st.file_uploader("Drop PDFs here",type=["pdf"],accept_multiple_files=True,label_visibility="collapsed")
if uploaded:
    st.markdown(f"**{len(uploaded)} file(s) ready:** "+" , ".join(f.name for f in uploaded))

if not get_api_key():
    st.markdown('<div class="info-box">👈 Enter your Anthropic API key in the sidebar to get started.</div>',unsafe_allow_html=True)

if st.button("⚡ Extract & Build Full Report",disabled=not(uploaded and get_api_key()),type="primary",use_container_width=True):
    with st.spinner("Reading PDFs with Claude... ~30 seconds"):
        try:
            data=extract(uploaded, settings["consolidated"])
            st.session_state["data_v2"]=data
        except Exception as e:
            st.error(f"Extraction failed: {e}")

if "data_v2" in st.session_state:
    data=st.session_state["data_v2"]
    cy=data.get("fiscal_year",fiscal_year); py_val=data.get("prior_year",prior_year)
    pl=data.get("pl",{}); bs=data.get("bs",{})

    st.success(f"✅ {data.get('company','')} · FY{cy} vs FY{py_val} · {'Consolidated' if consolidated else 'Detailed'}")
    st.markdown("---")

    if pl.get("available"):
        t_rev=pl.get("total_revenue",{}); t_np=pl.get("net_profit",{}); t_gp=pl.get("gross_profit",{})
        c1,c2,c3,c4=st.columns(4)
        rev_c=t_rev.get("current"); np_c=t_np.get("current"); gp_c=t_gp.get("current")
        with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-label">Revenue FY{cy}</div><div class="kpi-value">{fmt(rev_c,True)}</div><div class="kpi-sub">vs {fmt(t_rev.get("prior"),True)} prior</div></div>',unsafe_allow_html=True)
        with c2:
            cls="kpi-pos" if (gp_c or 0)>=0 else "kpi-neg"
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Gross profit</div><div class="kpi-value {cls}">{fmt(gp_c,True)}</div><div class="kpi-sub">Margin {pct(gp_c,rev_c)}</div></div>',unsafe_allow_html=True)
        with c3:
            cls="kpi-pos" if (np_c or 0)>=0 else "kpi-neg"
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Net profit</div><div class="kpi-value {cls}">{fmt(np_c,True)}</div><div class="kpi-sub">Margin {pct(np_c,rev_c)}</div></div>',unsafe_allow_html=True)
        with c4:
            ta=bs.get("total_assets",{}).get("current") if bs.get("available") else None
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Total assets</div><div class="kpi-value">{fmt(ta,True)}</div></div>',unsafe_allow_html=True)
        st.markdown("")

    st.markdown("---")
    col1,col2=st.columns(2)
    slug=f"FY{cy}"
    with col1:
        word_buf=build_word(data,settings)
        st.download_button("📄 Download Word (.docx)",data=word_buf,
            file_name=f"ActiveMedia_FinancialStatements_{slug}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,type="primary")
    with col2:
        excel_buf=build_excel(data,settings)
        st.download_button("📥 Download Excel (.xlsx)",data=excel_buf,
            file_name=f"ActiveMedia_FinancialStatements_{slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
