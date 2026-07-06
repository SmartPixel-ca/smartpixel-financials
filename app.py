import streamlit as st
import anthropic
import base64
import json
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

st.set_page_config(
    page_title="Active Média — Financial Statements",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main-header { background: linear-gradient(135deg, #1A2E4A 0%, #2E75B6 100%); padding: 2rem 2.5rem; border-radius: 12px; margin-bottom: 2rem; color: white; }
.main-header h1 { margin: 0; font-size: 1.6rem; font-weight: 600; }
.main-header p { margin: 0.3rem 0 0; opacity: 0.75; font-size: 0.88rem; }
.kpi-card { background: white; border: 1px solid #E8ECF0; border-radius: 10px; padding: 1.2rem 1.4rem; }
.kpi-label { font-size: 0.7rem; font-weight: 600; text-transform: uppercase; letter-spacing: .06em; color: #7A8799; margin-bottom: 4px; }
.kpi-value { font-size: 1.5rem; font-weight: 600; color: #1A2E4A; font-variant-numeric: tabular-nums; }
.kpi-sub { font-size: 0.73rem; color: #7A8799; margin-top: 2px; }
.kpi-pos { color: #1A7A4A; } .kpi-neg { color: #C0392B; }
.section-label { font-size: 0.7rem; font-weight: 600; text-transform: uppercase; letter-spacing: .08em; color: #2E75B6; margin: 1.2rem 0 0.4rem; }
.info-box { background: #EBF4FF; border-left: 3px solid #2E75B6; padding: 0.9rem 1.1rem; border-radius: 0 8px 8px 0; margin: 1rem 0; font-size: 0.88rem; color: #1A2E4A; }
.step-box { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; padding: 1rem 1.2rem; margin: 0.5rem 0; }
</style>
""", unsafe_allow_html=True)

# ── API key — from Streamlit secrets (deployed) or sidebar (local) ────────────
def get_api_key():
    try:
        key = st.secrets.get("ANTHROPIC_API_KEY", "")
        if key and not key.startswith("sk-ant-your"):
            return key
    except:
        pass
    return st.session_state.get("api_key", "")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
    <h1>📊 Financial Statements</h1>
    <p>Active Média inc. — SmartPixel &nbsp;·&nbsp; Upload QuickBooks PDFs to generate a complete audited-style report</p>
</div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Settings")
    if not get_api_key():
        manual_key = st.text_input("Anthropic API key", type="password", placeholder="sk-ant-...")
        if manual_key:
            st.session_state["api_key"] = manual_key
    else:
        st.success("API key configured ✓")

    st.markdown("---")
    st.markdown("**Company info**")
    company_name = st.text_input("Company name", value="Active Média inc.")
    fiscal_year   = st.text_input("Current fiscal year", value="2025")
    prior_year    = st.text_input("Prior fiscal year", value="2024")
    period_end    = st.text_input("Period end date", value="January 31, 2025")
    auditor_note  = st.text_input("Prepared by", value="Management")
    currency      = st.selectbox("Currency", ["CAD", "USD", "EUR"], index=0)

    st.markdown("---")
    st.markdown("**Include in report**")
    inc_cover    = st.checkbox("Cover page", value=True)
    inc_pl       = st.checkbox("Income statement (P&L)", value=True)
    inc_bs       = st.checkbox("Balance sheet", value=True)
    inc_cf       = st.checkbox("Cash flow statement", value=True)
    inc_annexes  = st.checkbox("Annexes / schedules", value=True)
    inc_notes    = st.checkbox("Notes section", value=True)
    st.markdown("---")
    st.caption("SmartPixel Financial Tool · v2.0")

# ── Helpers ───────────────────────────────────────────────────────────────────
def fmt(n, short=False):
    if n is None: return "—"
    try: n = float(n)
    except: return "—"
    if n == 0: return "—"
    if short:
        if abs(n) >= 1_000_000: return f"{'(' if n<0 else ''}${abs(n)/1_000_000:.1f}M{')' if n<0 else ''}"
        if abs(n) >= 1_000:     return f"{'(' if n<0 else ''}${abs(n)/1_000:.0f}K{')' if n<0 else ''}"
    s = f"{abs(n):,.0f}"
    return f"({s})" if n < 0 else s

def pct(a, b):
    try: return f"{float(a)/float(b)*100:.1f}%"
    except: return "—"

# ── Extraction prompt ─────────────────────────────────────────────────────────
EXTRACT_PROMPT = """Extract ALL financial data from these QuickBooks PDFs for Active Média inc. / SmartPixel.

IMPORTANT: Group accounts into clean audited-style line items (like a real annual report), NOT individual QuickBooks account codes.
For example, group all salary accounts into "Salaires et avantages sociaux", all software into "Logiciel", etc.
Preserve bilingual French/English naming as it appears.

The PDFs may include current year AND prior year — extract both.

Return ONLY valid JSON, no markdown, no explanation:
{
  "company": "string",
  "fiscal_year": "string",
  "prior_year": "string",
  "period_end": "string",
  "currency": "CAD",
  "pl": {
    "available": true,
    "revenue": [{"label": "string", "current": number, "prior": number}],
    "total_revenue": {"current": number, "prior": number},
    "cogs": [{"label": "string", "current": number, "prior": number}],
    "total_cogs": {"current": number, "prior": number},
    "gross_profit": {"current": number, "prior": number},
    "operating_expenses": [{"label": "string", "current": number, "prior": number}],
    "total_opex": {"current": number, "prior": number},
    "ebit": {"current": number, "prior": number},
    "financial_expenses": [{"label": "string", "current": number, "prior": number}],
    "total_financial": {"current": number, "prior": number},
    "other_income": [{"label": "string", "current": number, "prior": number}],
    "total_other": {"current": number, "prior": number},
    "net_profit": {"current": number, "prior": number},
    "annexes": {
      "cogs_detail": [{"label": "string", "current": number, "prior": number}],
      "opex_detail": [{"label": "string", "current": number, "prior": number}],
      "financial_detail": [{"label": "string", "current": number, "prior": number}],
      "other_detail": [{"label": "string", "current": number, "prior": number}]
    }
  },
  "bs": {
    "available": true,
    "current_assets": [{"label": "string", "current": number, "prior": number}],
    "total_current_assets": {"current": number, "prior": number},
    "longterm_assets": [{"label": "string", "current": number, "prior": number}],
    "total_longterm_assets": {"current": number, "prior": number},
    "total_assets": {"current": number, "prior": number},
    "current_liabilities": [{"label": "string", "current": number, "prior": number}],
    "total_current_liabilities": {"current": number, "prior": number},
    "longterm_liabilities": [{"label": "string", "current": number, "prior": number}],
    "total_longterm_liabilities": {"current": number, "prior": number},
    "total_liabilities": {"current": number, "prior": number},
    "equity": [{"label": "string", "current": number, "prior": number}],
    "total_equity": {"current": number, "prior": number},
    "total_liabilities_equity": {"current": number, "prior": number}
  },
  "cf": {
    "available": true,
    "operating": [{"label": "string", "current": number, "prior": number}],
    "total_operating": {"current": number, "prior": number},
    "investing": [{"label": "string", "current": number, "prior": number}],
    "total_investing": {"current": number, "prior": number},
    "financing": [{"label": "string", "current": number, "prior": number}],
    "total_financing": {"current": number, "prior": number},
    "net_change": {"current": number, "prior": number},
    "opening_cash": {"current": number, "prior": number},
    "closing_cash": {"current": number, "prior": number}
  }
}
If a statement is not present set available to false."""

# ── Extract ───────────────────────────────────────────────────────────────────
def extract(pdf_files):
    client = anthropic.Anthropic(api_key=get_api_key())
    content = []
    for f in pdf_files:
        b64 = base64.b64encode(f.read()).decode()
        content.append({"type":"document","source":{"type":"base64","media_type":"application/pdf","data":b64},"title":f.name})
    content.append({"type":"text","text":EXTRACT_PROMPT})
    msg = client.messages.create(model="claude-sonnet-4-6", max_tokens=8096, messages=[{"role":"user","content":content}])
    raw = msg.content[0].text.strip().lstrip("```json").lstrip("```").rstrip("```").strip()
    return json.loads(raw)

# ── Word builder ──────────────────────────────────────────────────────────────
def build_word(data, settings):
    doc = Document()
    for s in doc.sections:
        s.page_width=Inches(8.5); s.page_height=Inches(11)
        s.left_margin=s.right_margin=Inches(1.0)
        s.top_margin=s.bottom_margin=Inches(1.0)

    NAVY=RGBColor(0x1A,0x2E,0x4A); BLUE=RGBColor(0x2E,0x75,0xB6)
    WHITE=RGBColor(0xFF,0xFF,0xFF); GREY=RGBColor(0x55,0x55,0x55)
    BLACK=RGBColor(0,0,0)

    cy = data.get("fiscal_year", settings["fiscal_year"])
    py = data.get("prior_year",  settings["prior_year"])
    cur = settings["currency"]

    def set_bg(cell, hex_color):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        shd=OxmlElement('w:shd')
        shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),hex_color)
        tcPr.append(shd)

    def set_bdr(cell, color="CCCCCC"):
        tc=cell._tc; tcPr=tc.get_or_add_tcPr()
        tcB=OxmlElement('w:tcBorders')
        for side in ['top','left','bottom','right']:
            b=OxmlElement(f'w:{side}')
            b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4'); b.set(qn('w:color'),color)
            tcB.append(b)
        tcPr.append(tcB)

    def run(para, text, bold=False, italic=False, size=10, color=None):
        r=para.add_run(text); r.bold=bold; r.italic=italic
        r.font.name="Arial"; r.font.size=Pt(size)
        if color: r.font.color.rgb=color

    def page_hdr(title, period):
        p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(2)
        run(p, settings["company_name"], bold=True, size=14, color=NAVY)
        p2=doc.add_paragraph(); p2.paragraph_format.space_after=Pt(2)
        run(p2, title, size=11, color=BLUE)
        p3=doc.add_paragraph()
        pPr=p3._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'8'); bot.set(qn('w:color'),'2E75B6')
        pBdr.append(bot); pPr.append(pBdr); p3.paragraph_format.space_after=Pt(10)
        run(p3, period, italic=True, size=9, color=GREY)

    def fin_table(rows_data, col1_w=4500, col2_label=None, col3_label=None):
        col2_label = col2_label or cy
        col3_label = col3_label or py
        tbl=doc.add_table(rows=1, cols=3); tbl.style='Table Grid'
        hcells=tbl.rows[0].cells
        for i,(c,h) in enumerate(zip(hcells,[" ",col2_label+f"\n{cur} $",col3_label+f"\n{cur} $"])):
            set_bg(c,"1A2E4A"); set_bdr(c,"1A2E4A")
            p=c.paragraphs[0]
            p.alignment=WD_ALIGN_PARAGRAPH.CENTER if i>0 else WD_ALIGN_PARAGRAPH.LEFT
            run(p,h,bold=True,size=9,color=WHITE)

        zebra=False
        for rd in rows_data:
            style=rd.get("s","data"); lbl=rd.get("l","")
            v1=rd.get("v1",""); v2=rd.get("v2","")
            tr=tbl.add_row(); cs=tr.cells
            if style=="section":
                for c in cs: set_bg(c,"2E75B6"); set_bdr(c,"2E75B6")
                p=cs[0].paragraphs[0]; run(p,lbl,bold=True,size=9,color=WHITE)
                zebra=False; continue
            if style=="total":
                bg="D9E8F5"; fc=BLACK; bold=True
            elif style=="grand":
                bg="1A2E4A"; fc=WHITE; bold=True
            elif style=="blank":
                for c in cs:
                    set_bg(c,"FFFFFF")
                    s=Side(style="thin",color="EEEEEE")
                    c._tc.get_or_add_tcPr()
                continue
            else:
                bg="F5F5F5" if zebra else "FFFFFF"; fc=BLACK; bold=False; zebra=not zebra
            for c in cs: set_bg(c,bg); set_bdr(c)
            p=cs[0].paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.LEFT
            run(p,"  "+lbl if style=="data" else lbl,bold=bold,size=9,color=fc)
            for c,v in [(cs[1],v1),(cs[2],v2)]:
                p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
                run(p,str(v),bold=bold,size=9,color=fc)

        for row in tbl.rows:
            row.cells[0].width=Twips(col1_w)
            row.cells[1].width=Twips(2000)
            row.cells[2].width=Twips(2000)
        doc.add_paragraph().paragraph_format.space_after=Pt(4)

    def rows_from(items, cur_key="current", pri_key="prior", lbl_key="label", indent=True):
        out=[]
        for i,x in enumerate(items):
            out.append({"s":"data","l":x.get(lbl_key,""),"v1":fmt(x.get(cur_key)),"v2":fmt(x.get(pri_key))})
        return out

    def tv(obj, key):
        if not obj: return {"current":None,"prior":None}
        return obj.get(key,{"current":None,"prior":None})

    # ── COVER PAGE ────────────────────────────────────────────────────────────
    if settings["inc_cover"]:
        doc.add_paragraph().paragraph_format.space_before=Pt(80)
        p=doc.add_paragraph(); run(p,settings["company_name"],bold=True,size=28,color=NAVY)
        p2=doc.add_paragraph()
        pPr=p2._p.get_or_add_pPr(); pBdr=OxmlElement('w:pBdr')
        bot=OxmlElement('w:bottom'); bot.set(qn('w:val'),'single'); bot.set(qn('w:sz'),'16'); bot.set(qn('w:color'),'2E75B6')
        pBdr.append(bot); pPr.append(pBdr); p2.paragraph_format.space_after=Pt(24)
        p3=doc.add_paragraph(); run(p3,"ÉTATS FINANCIERS / FINANCIAL STATEMENTS",bold=True,size=18,color=NAVY)
        for line in [
            f"Exercice clos le / Year ended: {settings['period_end']}",
            f"Exercice comparatif / Comparative year: {py}",
            f"Monnaie / Currency: {cur}",
            f"Préparé par / Prepared by: {settings['auditor_note']}",
            f"Date de production: {datetime.now().strftime('%B %d, %Y')}",
            "Confidentiel — Usage interne / Confidential — Internal use only"
        ]:
            px=doc.add_paragraph()
            run(px,line,size=10,color=GREY,italic=line.startswith("Conf"))
        doc.add_page_break()

    pl=data.get("pl",{}); bs=data.get("bs",{}); cf=data.get("cf",{})

    # ── P&L ───────────────────────────────────────────────────────────────────
    if settings["inc_pl"] and pl.get("available"):
        page_hdr("État des résultats / Income Statement",
                 f"Exercice clos le {settings['period_end']}, avec informations comparatives de {py}")

        rows=[]
        rows+=rows_from(pl.get("revenue",[]))
        t=tv(pl,"total_revenue")
        rows.append({"s":"total","l":"Ventes / Revenue","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Coût des ventes / Cost of sales (Annexe 1)"})
        rows+=rows_from(pl.get("cogs",[]))
        t=tv(pl,"total_cogs")
        rows.append({"s":"total","l":"Total coût des ventes","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(pl,"gross_profit")
        rows.append({"s":"grand","l":"Bénéfice brut / Gross profit","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Charges / Operating expenses (Annexe 2)"})
        rows+=rows_from(pl.get("operating_expenses",[]))
        t=tv(pl,"total_opex")
        rows.append({"s":"total","l":"Total charges d'exploitation","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Frais financiers / Financial expenses (Annexe 3)"})
        rows+=rows_from(pl.get("financial_expenses",[]))
        t=tv(pl,"total_financial")
        rows.append({"s":"total","l":"Total frais financiers","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(pl,"ebit")
        rows.append({"s":"grand","l":"Bénéfice (perte) avant autres revenus","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        if pl.get("other_income"):
            rows.append({"s":"section","l":"Autres revenus / Other income (Annexe 4)"})
            rows+=rows_from(pl.get("other_income",[]))
            t=tv(pl,"total_other")
            rows.append({"s":"total","l":"Total autres revenus","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
            rows.append({"s":"blank"})

        t=tv(pl,"net_profit")
        rows.append({"s":"grand","l":"Bénéfice net (perte nette) / Net profit (loss)","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        fin_table(rows)
        p=doc.add_paragraph()
        run(p,"Se reporter aux notes afférentes aux états financiers.",italic=True,size=8,color=GREY)
        doc.add_page_break()

    # ── BALANCE SHEET ─────────────────────────────────────────────────────────
    if settings["inc_bs"] and bs.get("available"):
        page_hdr("Bilan / Balance Sheet", f"Au {settings['period_end']}, avec informations comparatives de {py}")
        rows=[]
        rows.append({"s":"section","l":"Actif à court terme / Current assets"})
        rows+=rows_from(bs.get("current_assets",[]))
        t=tv(bs,"total_current_assets")
        rows.append({"s":"total","l":"Total actif à court terme","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Actif à long terme / Long-term assets"})
        rows+=rows_from(bs.get("longterm_assets",[]))
        t=tv(bs,"total_longterm_assets")
        rows.append({"s":"total","l":"Total actif à long terme","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(bs,"total_assets")
        rows.append({"s":"grand","l":"TOTAL ACTIF / TOTAL ASSETS","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Passif à court terme / Current liabilities"})
        rows+=rows_from(bs.get("current_liabilities",[]))
        t=tv(bs,"total_current_liabilities")
        rows.append({"s":"total","l":"Total passif à court terme","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Passif à long terme / Long-term liabilities"})
        rows+=rows_from(bs.get("longterm_liabilities",[]))
        t=tv(bs,"total_longterm_liabilities")
        rows.append({"s":"total","l":"Total passif à long terme","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(bs,"total_liabilities")
        rows.append({"s":"grand","l":"TOTAL PASSIF / TOTAL LIABILITIES","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Avoir des actionnaires / Shareholders' equity"})
        rows+=rows_from(bs.get("equity",[]))
        t=tv(bs,"total_equity")
        rows.append({"s":"total","l":"Total avoir des actionnaires","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(bs,"total_liabilities_equity")
        rows.append({"s":"grand","l":"TOTAL PASSIF ET AVOIR / TOTAL L&E","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        fin_table(rows)
        p=doc.add_paragraph()
        run(p,"Se reporter aux notes afférentes aux états financiers.",italic=True,size=8,color=GREY)
        doc.add_page_break()

    # ── CASH FLOW ─────────────────────────────────────────────────────────────
    if settings["inc_cf"] and cf.get("available"):
        page_hdr("Flux de trésorerie / Cash Flow Statement",
                 f"Exercice clos le {settings['period_end']}, avec informations comparatives de {py}")
        rows=[]
        rows.append({"s":"section","l":"Activités d'exploitation / Operating activities"})
        rows+=rows_from(cf.get("operating",[]))
        t=tv(cf,"total_operating")
        rows.append({"s":"total","l":"Flux nets — exploitation","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Activités d'investissement / Investing activities"})
        rows+=rows_from(cf.get("investing",[]))
        t=tv(cf,"total_investing")
        rows.append({"s":"total","l":"Flux nets — investissement","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        rows.append({"s":"section","l":"Activités de financement / Financing activities"})
        rows+=rows_from(cf.get("financing",[]))
        t=tv(cf,"total_financing")
        rows.append({"s":"total","l":"Flux nets — financement","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        rows.append({"s":"blank"})

        t=tv(cf,"net_change")
        rows.append({"s":"grand","l":"Variation nette de la trésorerie","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        t=tv(cf,"opening_cash")
        rows.append({"s":"data","l":"Trésorerie — début d'exercice","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        t=tv(cf,"closing_cash")
        rows.append({"s":"grand","l":"Trésorerie — fin d'exercice","v1":fmt(t.get("current")),"v2":fmt(t.get("prior"))})
        fin_table(rows)
        doc.add_page_break()

    # ── ANNEXES ───────────────────────────────────────────────────────────────
    if settings["inc_annexes"] and pl.get("available"):
        annexes=pl.get("annexes",{})
        page_hdr("Annexes / Schedules", f"Exercice clos le {settings['period_end']}")

        for num, title, key in [
            ("1","Coût des ventes / Cost of sales","cogs_detail"),
            ("2","Frais d'exploitation / Operating expenses","opex_detail"),
            ("3","Frais financiers / Financial expenses","financial_detail"),
            ("4","Autres revenus / Other income","other_detail"),
        ]:
            items=annexes.get(key,[])
            if not items: continue
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4)
            run(p,f"Annexe {num} — {title}",bold=True,size=11,color=NAVY)
            rows=rows_from(items)
            if rows:
                total_cur=sum(x.get("current",0) or 0 for x in items)
                total_pri=sum(x.get("prior",0) or 0 for x in items)
                rows.append({"s":"total","l":f"Total — {title}","v1":fmt(total_cur),"v2":fmt(total_pri)})
            fin_table(rows)

    # ── NOTES ─────────────────────────────────────────────────────────────────
    if settings["inc_notes"]:
        page_hdr("Notes afférentes aux états financiers / Notes to Financial Statements",
                 f"Exercice clos le {settings['period_end']}")
        notes=[
            ("Note 1 — Principales méthodes comptables / Significant accounting policies",
             "Ces états financiers ont été préparés conformément aux Normes comptables canadiennes pour les entreprises à capital fermé (NCECF). / These financial statements have been prepared in accordance with Canadian Accounting Standards for Private Enterprises (ASPE)."),
            ("Note 2 — Base de présentation / Basis of presentation",
             f"Les états financiers sont présentés en dollars canadiens ({cur}). L'exercice financier se termine le {settings['period_end']}. / Financial statements are presented in Canadian dollars ({cur}). The fiscal year ends {settings['period_end']}."),
            ("Note 3 — Monnaie étrangère / Foreign currency",
             "Les opérations en devises étrangères sont converties au taux de change en vigueur à la date de la transaction. / Foreign currency transactions are translated at the exchange rate in effect at the transaction date."),
        ]
        for title, body in notes:
            p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4)
            run(p,title,bold=True,size=10,color=NAVY)
            p2=doc.add_paragraph(); run(p2,body,size=9,color=GREY)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf

# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(data, settings):
    wb=Workbook()
    NAVY,ACCENT,LIGHT,GREY,WHITE,BLUE="1A2E4A","2E75B6","D9E8F5","F5F5F5","FFFFFF","0000FF"
    cy=data.get("fiscal_year",settings["fiscal_year"])
    py=data.get("prior_year",settings["prior_year"])

    def fl(c): return PatternFill("solid",fgColor=c)
    def fnt(bold=False,color="000000",size=10,italic=False): return Font(name="Arial",bold=bold,color=color,size=size,italic=italic)
    def bdr():
        s=Side(style="thin",color="CCCCCC")
        return Border(left=s,right=s,top=s,bottom=s)
    def bdr_m():
        t=Side(style="thin",color="CCCCCC"); m=Side(style="medium",color=ACCENT)
        return Border(left=t,right=t,top=t,bottom=m)
    def aln(h="left"): return Alignment(horizontal=h,vertical="center")
    FMT='#,##0;(#,##0);"-"'

    def add_sheet(title, rows_data, col_labels):
        ws=wb.create_sheet(title)
        ws.sheet_view.showGridLines=False; ws.freeze_panes="B3"
        ws["A1"].value=f"{title} — {settings['company_name']}"
        ws["A1"].font=fnt(bold=True,color=NAVY,size=12)
        ws.merge_cells(f"A1:C1")
        ws.column_dimensions["A"].width=46
        ws.column_dimensions["B"].width=18; ws.column_dimensions["C"].width=18
        for i,(lbl,col) in enumerate([(col_labels[0],1),(col_labels[1],2),(col_labels[2],3)],1):
            c=ws.cell(row=2,column=i,value=lbl)
            c.font=fnt(bold=True,color=WHITE); c.fill=fl(NAVY); c.border=bdr()
            c.alignment=aln("center" if i>1 else "left")
        r=3; zebra=False
        for rd in rows_data:
            s=rd.get("s","data"); lbl=rd.get("l",""); v1=rd.get("v1"); v2=rd.get("v2")
            if s=="blank":
                for col in range(1,4):
                    c=ws.cell(row=r,column=col); c.fill=fl(WHITE)
                    sv=Side(style="thin",color="EEEEEE"); c.border=Border(left=sv,right=sv,top=sv,bottom=sv)
                r+=1; continue
            bg=NAVY if s=="grand" else LIGHT if s=="total" else ACCENT if s=="section" else GREY if zebra else WHITE
            fc=WHITE if s in ("grand","section") else "000000"
            bold=s in ("grand","total","section")
            if s=="data": zebra=not zebra
            else: zebra=False
            c=ws.cell(row=r,column=1,value=lbl)
            c.font=fnt(bold=bold,color=fc,size=9); c.fill=fl(bg)
            c.border=bdr_m() if s in ("grand","total") else bdr(); c.alignment=aln()
            for col,v in [(2,v1),(3,v2)]:
                fc2=ws.cell(row=r,column=col,value=v)
                fc2.font=fnt(bold=bold,color=fc,size=9); fc2.fill=fl(bg)
                fc2.number_format=FMT; fc2.border=bdr_m() if s in ("grand","total") else bdr(); fc2.alignment=aln("right")
            r+=1

    def sheet_rows(items, cur_key="current", pri_key="prior"):
        return [{"s":"data","l":x.get("label",""),"v1":x.get(cur_key),"v2":x.get(pri_key)} for x in items]

    def tv(obj,key):
        if not obj: return {"current":None,"prior":None}
        return obj.get(key,{"current":None,"prior":None})

    # Remove default sheet
    ws=wb.active; ws.title="Cover"
    ws["A1"].value=settings["company_name"]; ws["A1"].font=fnt(bold=True,color=NAVY,size=14)
    ws["A2"].value=f"Financial Statements — FY{cy} vs FY{py}"
    ws["A3"].value=f"Period end: {settings['period_end']}"
    ws["A4"].value=f"Currency: {settings['currency']}"
    ws.column_dimensions["A"].width=40

    pl=data.get("pl",{}); bs=data.get("bs",{}); cf=data.get("cf",{})

    if pl.get("available"):
        rows=[]
        rows.append({"s":"section","l":"Revenue"})
        rows+=sheet_rows(pl.get("revenue",[]))
        t=tv(pl,"total_revenue"); rows.append({"s":"total","l":"Total revenue","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Cost of sales"})
        rows+=sheet_rows(pl.get("cogs",[]))
        t=tv(pl,"total_cogs"); rows.append({"s":"total","l":"Total COGS","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(pl,"gross_profit"); rows.append({"s":"grand","l":"Gross profit","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Operating expenses"})
        rows+=sheet_rows(pl.get("operating_expenses",[]))
        t=tv(pl,"total_opex"); rows.append({"s":"total","l":"Total OpEx","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Financial expenses"})
        rows+=sheet_rows(pl.get("financial_expenses",[]))
        t=tv(pl,"total_financial"); rows.append({"s":"total","l":"Total financial","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(pl,"ebit"); rows.append({"s":"grand","l":"Profit before other income","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        if pl.get("other_income"):
            rows.append({"s":"section","l":"Other income"})
            rows+=sheet_rows(pl.get("other_income",[]))
            t=tv(pl,"total_other"); rows.append({"s":"total","l":"Total other income","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(pl,"net_profit"); rows.append({"s":"grand","l":"Net profit / (loss)","v1":t["current"],"v2":t["prior"]})
        add_sheet("P&L",rows,[" ",f"FY{cy} ({settings['currency']} $)",f"FY{py} ({settings['currency']} $)"])

        if pl.get("annexes"):
            ann=pl["annexes"]; ann_rows=[]
            for num,title,key in [("1","Cost of sales","cogs_detail"),("2","Operating expenses","opex_detail"),("3","Financial expenses","financial_detail"),("4","Other income","other_detail")]:
                items=ann.get(key,[])
                if not items: continue
                ann_rows.append({"s":"section","l":f"Annex {num} — {title}"})
                ann_rows+=sheet_rows(items)
                ann_rows.append({"s":"total","l":f"Total — {title}","v1":sum(x.get("current",0) or 0 for x in items),"v2":sum(x.get("prior",0) or 0 for x in items)})
                ann_rows.append({"s":"blank"})
            if ann_rows:
                add_sheet("Annexes",ann_rows,[" ",f"FY{cy}",f"FY{py}"])

    if bs.get("available"):
        rows=[]
        rows.append({"s":"section","l":"Current assets"})
        rows+=sheet_rows(bs.get("current_assets",[]))
        t=tv(bs,"total_current_assets"); rows.append({"s":"total","l":"Total current assets","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Long-term assets"})
        rows+=sheet_rows(bs.get("longterm_assets",[]))
        t=tv(bs,"total_longterm_assets"); rows.append({"s":"total","l":"Total long-term assets","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(bs,"total_assets"); rows.append({"s":"grand","l":"TOTAL ASSETS","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Current liabilities"})
        rows+=sheet_rows(bs.get("current_liabilities",[]))
        t=tv(bs,"total_current_liabilities"); rows.append({"s":"total","l":"Total current liabilities","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Long-term liabilities"})
        rows+=sheet_rows(bs.get("longterm_liabilities",[]))
        t=tv(bs,"total_longterm_liabilities"); rows.append({"s":"total","l":"Total long-term liabilities","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(bs,"total_liabilities"); rows.append({"s":"grand","l":"TOTAL LIABILITIES","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Shareholders' equity"})
        rows+=sheet_rows(bs.get("equity",[]))
        t=tv(bs,"total_equity"); rows.append({"s":"total","l":"Total equity","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(bs,"total_liabilities_equity"); rows.append({"s":"grand","l":"TOTAL LIABILITIES & EQUITY","v1":t["current"],"v2":t["prior"]})
        add_sheet("Balance Sheet",rows,[" ",f"FY{cy} ({settings['currency']} $)",f"FY{py} ({settings['currency']} $)"])

    if cf.get("available"):
        rows=[]
        rows.append({"s":"section","l":"Operating activities"})
        rows+=sheet_rows(cf.get("operating",[]))
        t=tv(cf,"total_operating"); rows.append({"s":"total","l":"Net operating cash flow","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Investing activities"})
        rows+=sheet_rows(cf.get("investing",[]))
        t=tv(cf,"total_investing"); rows.append({"s":"total","l":"Net investing cash flow","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        rows.append({"s":"section","l":"Financing activities"})
        rows+=sheet_rows(cf.get("financing",[]))
        t=tv(cf,"total_financing"); rows.append({"s":"total","l":"Net financing cash flow","v1":t["current"],"v2":t["prior"]}); rows.append({"s":"blank"})
        t=tv(cf,"net_change"); rows.append({"s":"grand","l":"Net change in cash","v1":t["current"],"v2":t["prior"]})
        t=tv(cf,"opening_cash"); rows.append({"s":"data","l":"Opening cash","v1":t["current"],"v2":t["prior"]})
        t=tv(cf,"closing_cash"); rows.append({"s":"grand","l":"Closing cash","v1":t["current"],"v2":t["prior"]})
        add_sheet("Cash Flow",rows,[" ",f"FY{cy}",f"FY{py}"])

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf

# ── Main UI ───────────────────────────────────────────────────────────────────
settings = {
    "company_name": company_name, "fiscal_year": fiscal_year,
    "prior_year": prior_year, "period_end": period_end,
    "auditor_note": auditor_note, "currency": currency,
    "inc_cover": inc_cover, "inc_pl": inc_pl, "inc_bs": inc_bs,
    "inc_cf": inc_cf, "inc_annexes": inc_annexes, "inc_notes": inc_notes,
}

st.markdown('<div class="section-label">Step 1 — Upload QuickBooks PDFs</div>', unsafe_allow_html=True)
st.markdown('<div class="info-box">Upload current year + prior year PDFs together — Claude will extract both automatically and build the two-column comparison.</div>', unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Drop PDFs here", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed"
)
if uploaded:
    st.markdown(f"**{len(uploaded)} file(s) ready:** " + ", ".join(f.name for f in uploaded))

st.markdown("")
can_run = bool(uploaded) and bool(get_api_key())
if not get_api_key():
    st.markdown('<div class="info-box">👈 Enter your Anthropic API key in the sidebar to get started.</div>', unsafe_allow_html=True)

if st.button("⚡ Extract & Build Full Report", disabled=not can_run, type="primary", use_container_width=True):
    with st.spinner("Reading PDFs and extracting data with Claude... this takes about 30 seconds."):
        try:
            data = extract(uploaded)
            st.session_state["data_v2"] = data
        except Exception as e:
            st.error(f"Extraction failed: {e}")

if "data_v2" in st.session_state:
    data = st.session_state["data_v2"]
    cy = data.get("fiscal_year", fiscal_year)
    py_val = data.get("prior_year", prior_year)
    pl = data.get("pl", {}); bs = data.get("bs", {}); cf = data.get("cf", {})

    st.success(f"✅ Extracted — {data.get('company','')} · FY{cy} vs FY{py_val}")
    st.markdown("---")

    # KPIs
    if pl.get("available"):
        t_rev = pl.get("total_revenue", {}); t_np = pl.get("net_profit", {}); t_gp = pl.get("gross_profit", {})
        c1,c2,c3,c4 = st.columns(4)
        rev_cur = t_rev.get("current"); rev_pri = t_rev.get("prior")
        np_cur = t_np.get("current"); gp_cur = t_gp.get("current")
        with c1: st.markdown(f'<div class="kpi-card"><div class="kpi-label">Revenue FY{cy}</div><div class="kpi-value">{fmt(rev_cur,True)}</div><div class="kpi-sub">vs {fmt(rev_pri,True)} prior year</div></div>',unsafe_allow_html=True)
        with c2:
            cls="kpi-pos" if (gp_cur or 0)>=0 else "kpi-neg"
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Gross profit FY{cy}</div><div class="kpi-value {cls}">{fmt(gp_cur,True)}</div><div class="kpi-sub">Margin {pct(gp_cur,rev_cur)}</div></div>',unsafe_allow_html=True)
        with c3:
            cls="kpi-pos" if (np_cur or 0)>=0 else "kpi-neg"
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Net profit FY{cy}</div><div class="kpi-value {cls}">{fmt(np_cur,True)}</div><div class="kpi-sub">Margin {pct(np_cur,rev_cur)}</div></div>',unsafe_allow_html=True)
        with c4:
            ta = bs.get("total_assets",{}).get("current") if bs.get("available") else None
            st.markdown(f'<div class="kpi-card"><div class="kpi-label">Total assets</div><div class="kpi-value">{fmt(ta,True)}</div><div class="kpi-sub">As at {period_end}</div></div>',unsafe_allow_html=True)
        st.markdown("")

    # Preview tabs
    tabs_available = []
    if pl.get("available"): tabs_available.append("📋 P&L")
    if bs.get("available"): tabs_available.append("⚖️ Balance sheet")
    if cf.get("available"): tabs_available.append("💰 Cash flow")

    if tabs_available:
        import pandas as pd
        selected = st.tabs(tabs_available)
        idx=0

        if pl.get("available"):
            with selected[idx]:
                def pl_df(items):
                    return pd.DataFrame([{"Account":x.get("label",""), f"FY{cy} ({currency})":fmt(x.get("current")), f"FY{py_val} ({currency})":fmt(x.get("prior"))} for x in items])
                st.markdown(f"**Revenue**")
                st.table(pl_df(pl.get("revenue",[])))
                t=pl.get("total_revenue",{}); st.markdown(f"**Total revenue: {fmt(t.get('current'))}** (prior: {fmt(t.get('prior'))})")
                st.markdown(f"**Gross profit: {fmt(pl.get('gross_profit',{}).get('current'))}** · margin {pct(pl.get('gross_profit',{}).get('current'),t.get('current'))}")
                st.markdown("**Operating expenses**")
                st.table(pl_df(pl.get("operating_expenses",[])))
                t=pl.get("net_profit",{}); cls="🟢" if (t.get("current") or 0)>=0 else "🔴"
                st.markdown(f"### {cls} Net profit: **{fmt(t.get('current'))}** (prior: {fmt(t.get('prior'))})")
            idx+=1

        if bs.get("available"):
            with selected[idx]:
                def bs_df(items):
                    return pd.DataFrame([{"Account":x.get("label",""), f"FY{cy} ({currency})":fmt(x.get("current")), f"FY{py_val} ({currency})":fmt(x.get("prior"))} for x in items])
                st.markdown("**Current assets**"); st.table(bs_df(bs.get("current_assets",[])))
                st.markdown("**Long-term assets**"); st.table(bs_df(bs.get("longterm_assets",[])))
                t=bs.get("total_assets",{}); st.markdown(f"**Total assets: {fmt(t.get('current'))}** (prior: {fmt(t.get('prior'))})")
                st.markdown("**Current liabilities**"); st.table(bs_df(bs.get("current_liabilities",[])))
                st.markdown("**Long-term liabilities**"); st.table(bs_df(bs.get("longterm_liabilities",[])))
                t=bs.get("total_equity",{}); st.markdown(f"**Total equity: {fmt(t.get('current'))}** (prior: {fmt(t.get('prior'))})")
            idx+=1

        if cf.get("available"):
            with selected[idx]:
                def cf_df(items):
                    return pd.DataFrame([{"Item":x.get("label",""), f"FY{cy}":fmt(x.get("current")), f"FY{py_val}":fmt(x.get("prior"))} for x in items])
                st.markdown("**Operating**"); st.table(cf_df(cf.get("operating",[])))
                st.markdown("**Investing**"); st.table(cf_df(cf.get("investing",[])))
                st.markdown("**Financing**"); st.table(cf_df(cf.get("financing",[])))
                t=cf.get("closing_cash",{}); st.markdown(f"**Closing cash: {fmt(t.get('current'))}**")

    st.markdown("---")
    st.markdown('<div class="section-label">Download complete package</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    slug = f"FY{cy}"

    with col1:
        with st.spinner("Building Word document..."):
            word_buf = build_word(data, settings)
        st.download_button("📄 Download Word (.docx) — Full report",
            data=word_buf, file_name=f"ActiveMedia_FinancialStatements_{slug}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True, type="primary")

    with col2:
        with st.spinner("Building Excel..."):
            excel_buf = build_excel(data, settings)
        st.download_button("📥 Download Excel (.xlsx) — All sheets",
            data=excel_buf, file_name=f"ActiveMedia_FinancialStatements_{slug}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
