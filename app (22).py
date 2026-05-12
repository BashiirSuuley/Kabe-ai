"""
Kaabe v4 — Academic AI Platform (كعبه)
Fast · Bilingual EN/AR · Real Word/PPTX output
Zero external packages — streamlit + requests only
"""
import streamlit as st
import requests, json, re, base64, io, zipfile, random

st.set_page_config(page_title="Kaabe — Academic AI", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@700;800&family=DM+Sans:wght@400;500;600&family=Amiri:wght@400;700&display=swap');
*{box-sizing:border-box;}
html,body,[class*=css]{font-family:'DM Sans',sans-serif;background:#07090f!important;color:#dde3ee!important;}
.stApp{background:#07090f!important;}
section[data-testid=stSidebar]{background:#04060c!important;border-right:1px solid #141c2e;}
section[data-testid=stSidebar] *{color:#7a8daa!important;}
#MainMenu,footer,header{visibility:hidden;}
.hero{background:linear-gradient(135deg,#0c1524,#0f1e3a,#0c1524);border:1px solid #1c304e;border-radius:16px;padding:1.6rem 2.2rem;margin-bottom:1.2rem;display:flex;align-items:center;gap:1.4rem;}
.hero-ico{width:54px;height:54px;border-radius:12px;background:linear-gradient(135deg,#1e5fa8,#2d7dd2);display:flex;align-items:center;justify-content:center;font-size:1.7rem;flex-shrink:0;box-shadow:0 4px 14px rgba(45,125,210,.35);}
.hero h1{font-family:'Syne',sans-serif;font-size:1.75rem;font-weight:800;color:#fff;margin:0 0 2px;}
.hero .ar{font-family:'Amiri',serif;font-size:1rem;color:#4a9ede;margin:0 0 .25rem;direction:rtl;}
.hero p{font-size:.8rem;color:#5a7099;margin:0;}
.hero-right{margin-left:auto;display:flex;flex-direction:column;gap:4px;align-items:flex-end;}
.pill{display:inline-flex;align-items:center;gap:4px;background:#0d1628;border:1px solid #1a2a45;border-radius:20px;padding:3px 10px;font-size:.73rem;color:#8899bb;}
.pill b{color:#fff;}
.card{background:#0c1422;border:1px solid #141c2e;border-radius:11px;padding:1.1rem 1.3rem;margin-bottom:.8rem;}
.ctitle{font-family:'Syne',sans-serif;font-size:.92rem;font-weight:700;color:#fff;margin-bottom:.6rem;}
.mcq-wrap{background:#090e1a;border:1px solid #141c2e;border-radius:9px;padding:1rem 1.1rem;margin-bottom:.7rem;}
.mcq-q{font-weight:600;font-size:.9rem;color:#e0e8f8;margin-bottom:.5rem;}
.mcq-ar{font-family:'Amiri',serif;font-size:.92rem;color:#7ab3d4;direction:rtl;text-align:right;margin-bottom:.4rem;}
.mcq-opt{display:flex;align-items:flex-start;gap:8px;padding:.28rem .4rem;border-radius:5px;font-size:.83rem;color:#8899bb;margin-bottom:2px;border:1px solid transparent;}
.mcq-opt.ok{background:#071a0e;border-color:#14532d;color:#4ade80;}
.mcq-ltr{font-family:'Syne',sans-serif;font-weight:700;color:#2563eb;min-width:17px;}
.mcq-exp{background:#0c1e33;border-left:3px solid #2563eb;border-radius:0 5px 5px 0;padding:.4rem .7rem;font-size:.77rem;color:#7ab3d4;margin-top:.4rem;}
.fc{background:linear-gradient(135deg,#0b1730,#0f1e38);border:1px solid #1e3a6a;border-radius:13px;padding:1.8rem;text-align:center;min-height:170px;display:flex;flex-direction:column;align-items:center;justify-content:center;margin-bottom:.9rem;cursor:pointer;transition:all .2s;}
.fc:hover{border-color:#2563eb;transform:translateY(-2px);}
.fc-lbl{font-size:.64rem;color:#2563eb;text-transform:uppercase;letter-spacing:.12em;margin-bottom:.6rem;}
.fc-txt{font-size:1rem;font-weight:600;color:#e0e8f8;}
.fc-ar{font-family:'Amiri',serif;font-size:.97rem;color:#7ab3d4;direction:rtl;margin-top:.5rem;}
.sum-s{background:#090e1a;border:1px solid #141c2e;border-radius:9px;padding:1rem 1.1rem;margin-bottom:.65rem;}
.sum-h{font-family:'Syne',sans-serif;font-size:.88rem;font-weight:700;color:#2563eb;margin-bottom:.35rem;}
.sum-ar{font-family:'Amiri',serif;font-size:.9rem;color:#4a9ede;direction:rtl;text-align:right;margin-bottom:.3rem;}
.sum-txt{font-size:.83rem;color:#8899bb;line-height:1.7;}
.kterm{display:inline-block;background:#0c1e33;border:1px solid #1e3a6a;border-radius:5px;padding:2px 8px;font-size:.75rem;color:#60a5fa;margin:2px;}
.exam-sec{background:#0d1628;border:1px solid #1e3a6a;border-radius:7px;padding:.5rem .9rem;font-family:'Syne',sans-serif;font-size:.85rem;font-weight:700;color:#93c5fd;margin:.7rem 0 .4rem;}
.exam-q{padding:.4rem 0;border-bottom:1px solid #111826;font-size:.86rem;color:#ccd6f6;}
.exam-qn{color:#2563eb;font-weight:700;margin-right:5px;}
.mktag{float:right;background:#0c1628;border:1px solid #1e3a6a;border-radius:20px;padding:1px 7px;font-size:.69rem;color:#7ab3d4;}
.ans-ln{border-bottom:1px dashed #1e2a40;height:22px;margin:.25rem 0;}
.qmsg{padding:.7rem .9rem;border-radius:9px;margin-bottom:.5rem;font-size:.86rem;line-height:1.6;}
.quser{background:#0c1e33;border:1px solid #1e3a6a;color:#c0d4f5;text-align:right;}
.qai{background:#090e1a;border:1px solid #141c2e;color:#8ab4d4;}
.qlbl{font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#2563eb;margin-bottom:.22rem;}
.pb{background:#0d1628;border-radius:6px;height:4px;overflow:hidden;margin:.35rem 0;}
.pf{height:100%;border-radius:6px;background:linear-gradient(90deg,#2563eb,#10b981);transition:width .3s ease;}
.sbox{background:#090e1a;border:1px solid #141c2e;border-radius:9px;padding:.85rem 1rem;font-size:.78rem;color:#5a7099;line-height:1.65;}
.sbox b{color:#dde3ee;}
div[data-testid="stExpander"]{background:#090e1a;border:1px solid #141c2e;border-radius:8px;}
.stButton>button{font-family:'DM Sans',sans-serif!important;}
.stTextInput>div>div,.stTextArea textarea,.stSelectbox>div>div,.stNumberInput>div>div>input{background:#0d1628!important;border-color:#1a2a45!important;color:#dde3ee!important;}
.stTabs [data-baseweb=tab-list]{background:#0d1628;border-radius:7px;padding:3px;gap:3px;}
.stTabs [data-baseweb=tab]{color:#5a7099!important;font-family:'DM Sans',sans-serif!important;font-size:12px!important;}
.stTabs [aria-selected=true]{background:#1a2a45!important;color:#fff!important;border-radius:5px!important;}
</style>
""", unsafe_allow_html=True)

# ── THEMES ────────────────────────────────────────────────────
THEMES = {
    "Academic Dark":   {"bg":"0F1729","tb":"1E293B","ac":"60A5FA","tx":"F8FAFC","su":"94A3B8","bx":"182538","mn":"93C5FD"},
    "Royal Blue":      {"bg":"0A1940","tb":"0F2850","ac":"3B82F6","tx":"FFFFFF","su":"93C5FD","bx":"143264","mn":"BAE6FD"},
    "Forest Scholar":  {"bg":"0A1A0A","tb":"142E14","ac":"4ADE80","tx":"F0FDF4","su":"86EFAC","bx":"143214","mn":"BBF7D0"},
    "Executive White": {"bg":"FFFFFF","tb":"EFF6FF","ac":"1D4ED8","tx":"0F172A","su":"475569","bx":"F1F5F9","mn":"1E40AF"},
    "Crimson Power":   {"bg":"140505","tb":"280A0A","ac":"EF4444","tx":"FEF2F2","su":"FCA5A5","bx":"280A0A","mn":"FECACA"},
    "Purple Scholar":  {"bg":"0F0A1A","tb":"1E1040","ac":"A855F7","tx":"FAF5FF","su":"D8B4FE","bx":"1A1035","mn":"E9D5FF"},
    "Teal Modern":     {"bg":"021B1F","tb":"083344","ac":"06B6D4","tx":"ECFEFF","su":"67E8F9","bx":"0E3A45","mn":"A5F3FC"},
    "Gold Prestige":   {"bg":"1A1400","tb":"2D2200","ac":"F59E0B","tx":"FFFBEB","su":"FCD34D","bx":"2A1F00","mn":"FDE68A"},
}

# Two static dicts — picked at runtime based on ar_only
SECTION_TYPES_EN = {
    "mcq":          "MCQ (Multiple Choice)",
    "true_false":   "True / False",
    "short_answer": "Short Answer",
    "long_answer":  "Long Answer / Essay",
    "calculation":  "Calculation",
    "fill_blank":   "Fill in the Blank",
    "matching":     "Matching",
    "case_study":   "Case Study",
}
SECTION_TYPES_AR = {
    "mcq":          "اختيار من متعدد",
    "true_false":   "صح أو خطأ",
    "short_answer": "إجابة قصيرة",
    "long_answer":  "إجابة مطولة",
    "calculation":  "حساب ومسائل",
    "fill_blank":   "أكمل الفراغ",
    "matching":     "مطابقة",
    "case_study":   "دراسة حالة",
}

# ── GEMINI — fast, single best model first ────────────────────
def call_gemini(api_key, prompt, pdf_b64=None, tokens=3500, json_mode=True):
    """Tries 5 models in order. Auto-retries on 404 and 429 rate limits."""
    MODELS = [
        "gemini-2.0-flash",
        "gemini-2.0-flash-lite",
        "gemini-1.5-flash",
        "gemini-2.5-flash",
        "gemini-1.5-flash-latest",
    ]
    parts = []
    if pdf_b64:
        parts.append({"inline_data": {"mime_type": "application/pdf", "data": pdf_b64}})
    parts.append({"text": prompt})
    cfg = {"temperature": 0.1, "maxOutputTokens": tokens}
    if json_mode:
        cfg["responseMimeType"] = "application/json"
    last_err = ""
    for model in MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
        try:
            r = requests.post(
                url,
                json={"contents": [{"parts": parts}], "generationConfig": cfg},
                timeout=60
            )
        except requests.exceptions.Timeout:
            raise ValueError("⏱️ Request timed out. Try a shorter PDF or fewer questions.")
        except requests.exceptions.ConnectionError:
            raise ValueError("🌐 No internet connection.")
        if r.status_code == 200:
            try:
                return r.json()["candidates"][0]["content"]["parts"][0]["text"]
            except (KeyError, IndexError):
                raise ValueError("Unexpected API response format.")
        elif r.status_code == 404:
            last_err = f"Model {model} not found"
            continue
        elif r.status_code == 429:
            # Auto-retry on next model instead of crashing
            last_err = f"Rate limit on {model}"
            continue
        elif r.status_code == 403:
            raise ValueError("🔑 Invalid API key. Get your free key at aistudio.google.com/app/apikey")
        elif r.status_code == 400:
            raise ValueError("📄 PDF may be too large. Try a document under 5MB.")
        else:
            last_err = f"Error {r.status_code}"
            continue
    raise ValueError(
        f"All Gemini models are busy or unavailable. Last: {last_err}\n\n"
        f"✅ Try this:\n"
        f"• Wait 60 seconds and click Generate again\n"
        f"• Reduce number of slides/questions in sidebar\n"
        f"• The free tier allows 15 requests per minute"
    )

# ── JSON REPAIR ────────────────────────────────────────────────
def safe_parse(raw: str) -> dict:
    t = raw.strip()
    t = re.sub(r"^```(?:json)?", "", t).strip()
    t = re.sub(r"```$", "", t).strip()
    s = t.find("{"); e = t.rfind("}") + 1
    if s == -1: raise ValueError("No JSON found in AI response. Click Generate again.")
    c = t[s:e]
    # fix literal newlines in strings
    r2 = []; in_s = False; esc = False
    for ch in c:
        if esc: r2.append(ch); esc = False; continue
        if ch == "\\" and in_s: esc = True; r2.append(ch); continue
        if ch == '"': in_s = not in_s; r2.append(ch); continue
        if in_s and ch in "\n\r": r2.append(" "); continue
        r2.append(ch)
    c = "".join(r2)
    c = re.sub(r",\s*([\}\]])", r"\1", c)
    c = re.sub(r'(?<!\\)\\([^"\\/bfnrtu])', r' ', c)
    try: return json.loads(c)
    except:
        # close unclosed structures
        stk = []; in_s = False; esc = False
        for ch in c:
            if esc: esc = False; continue
            if ch == "\\" and in_s: esc = True; continue
            if ch == '"': in_s = not in_s; continue
            if not in_s:
                if ch in "{[": stk.append("}" if ch == "{" else "]")
                elif ch in "}]" and stk and stk[-1] == ch: stk.pop()
        c = c + "".join(reversed(stk))
        c = re.sub(r",\s*([\}\]])", r"\1", c)
        try: return json.loads(c)
        except: raise ValueError("Could not parse AI response. Click Generate again.")

# ══════════════════════════════════════════════════════════════
# WORD DOCX BUILDER — zero external packages
# ══════════════════════════════════════════════════════════════
def we(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

def build_docx(body_xml: str) -> bytes:
    CT = "application/vnd.openxmlformats-officedocument"
    content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="{CT}.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"   ContentType="{CT}.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="{CT}.wordprocessingml.settings+xml"/>
</Types>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    doc_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"   Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Amiri"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/></w:pPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="H1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1D3461"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="H2">
    <w:name w:val="heading 2"/>
    <w:pPr><w:spacing w:before="200" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="1E4D8C"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="AR">
    <w:name w:val="Arabic"/>
    <w:pPr><w:bidi/><w:jc w:val="right"/><w:spacing w:after="80"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Amiri" w:hAnsi="Amiri" w:cs="Amiri"/>
    <w:sz w:val="24"/><w:szCs w:val="24"/><w:color w:val="1E3A5F"/></w:rPr>
  </w:style>
</w:styles>"""
    settings = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
</w:settings>"""
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body_xml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>
    </w:sectPr>
  </w:body>
</w:document>"""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
    buf.seek(0)
    return buf.read()

def wp(text, bold=False, italic=False, sz=24, color=None, align="left", after=120, indent=0, rtl=False, font=None):
    b = "<w:b/>" if bold else ""
    i = "<w:i/>" if italic else ""
    co = f'<w:color w:val="{color}"/>' if color else ""
    fn_tag = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' if font else ""
    rt = "<w:rtl/>" if rtl else ""
    bidi = "<w:bidi/>" if rtl else ""
    alg = {"left":"left","center":"center","right":"right","justify":"both"}.get(align,"left")
    ind_tag = f'<w:ind w:left="{indent}"/>' if indent else ""
    return (f'<w:p><w:pPr><w:jc w:val="{alg}"/><w:spacing w:after="{after}"/>{ind_tag}{bidi}</w:pPr>'
            f'<w:r><w:rPr>{b}{i}{co}{fn_tag}{rt}<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{we(text)}</w:t></w:r></w:p>')

def wp_ar(text, sz=24, color="1E3A5F", after=100):
    return (f'<w:p><w:pPr><w:bidi/><w:jc w:val="right"/><w:spacing w:after="{after}"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="Amiri" w:hAnsi="Amiri" w:cs="Amiri"/>'
            f'<w:rtl/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/><w:color w:val="{color}"/></w:rPr>'
            f'<w:t xml:space="preserve">{we(text)}</w:t></w:r></w:p>')

def wp_runs(runs, align="left", after=120, indent=0):
    alg = {"left":"left","center":"center","right":"right"}.get(align,"left")
    ind_tag = f'<w:ind w:left="{indent}"/>' if indent else ""
    xml = ""
    for r in runs:
        b = "<w:b/>" if r.get("b") else ""
        i = "<w:i/>" if r.get("i") else ""
        sz = r.get("sz", 22)
        rc2 = r.get("co","")
        co = f'<w:color w:val="{rc2}"/>' if rc2 else ""
        fn = r.get("fn","")
        fn_tag = f'<w:rFonts w:ascii="{fn}" w:hAnsi="{fn}" w:cs="{fn}"/>' if fn else ""
        rt = "<w:rtl/>" if r.get("rtl") else ""
        xml += (f'<w:r><w:rPr>{b}{i}{co}{fn_tag}{rt}'
                f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
                f'<w:t xml:space="preserve">{we(r.get("t",""))}</w:t></w:r>')
    return f'<w:p><w:pPr><w:jc w:val="{alg}"/><w:spacing w:after="{after}"/>{ind_tag}</w:pPr>{xml}</w:p>'

def wempty(n=1):
    return '<w:p><w:pPr><w:spacing w:after="0"/></w:pPr></w:p>' * n

def whr(col="2563EB"):
    return (f'<w:p><w:pPr><w:pBdr>'
            f'<w:bottom w:val="single" w:sz="6" w:space="1" w:color="{col}"/>'
            f'</w:pBdr><w:spacing w:after="80"/></w:pPr></w:p>')

def wans(n=4):
    lines = ""
    for _ in range(n):
        lines += ('<w:p><w:pPr><w:spacing w:after="0"/>'
                  '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="BBBBBB"/></w:pBdr>'
                  '</w:pPr></w:p>')
    return lines + wempty()

def wtable(h_en, h_ar, rows, cws=None):
    """Real Word table with blue header, alternating rows."""
    n = len(h_en)
    tw = 9360
    if not cws:
        cws = [tw // n] * n

    def cell_hdr(en, ar, cw):
        return (f'<w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="1D3461"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>'
                f'<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
                f'<w:t>{we(en)}</w:t></w:r>'
                f'{"<w:r><w:rPr><w:rFonts w:ascii=\"Amiri\" w:hAnsi=\"Amiri\" w:cs=\"Amiri\"/><w:rtl/><w:color w:val=\"BAD0FF\"/><w:sz w:val=\"18\"/><w:szCs w:val=\"18\"/></w:rPr><w:t> | "+we(ar)+"</w:t></w:r>" if ar else ""}'
                f'</w:p></w:tc>')

    def cell_data(val, cw, fill):
        return (f'<w:tc><w:tcPr><w:tcW w:w="{cw}" w:type="dxa"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/></w:tcPr>'
                f'<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="60"/></w:pPr>'
                f'<w:r><w:rPr><w:sz w:val="20"/><w:szCs w:val="20"/><w:color w:val="1E293B"/></w:rPr>'
                f'<w:t xml:space="preserve">{we(str(val))}</w:t></w:r></w:p></w:tc>')

    hdr_cells = "".join(cell_hdr(h_en[i], h_ar[i] if i < len(h_ar) else "", cws[i] if i < len(cws) else cws[-1]) for i in range(n))
    body = ""
    for ri, row in enumerate(rows):
        fill = "EEF3FB" if ri % 2 == 0 else "F8FAFF"
        cells = "".join(cell_data(row[ci] if ci < len(row) else "", cws[ci] if ci < len(cws) else cws[-1], fill) for ci in range(n))
        body += f"<w:tr>{cells}</w:tr>"

    return (f'<w:tbl><w:tblPr>'
            f'<w:tblW w:w="{tw}" w:type="dxa"/>'
            f'<w:tblBorders>'
            f'<w:top    w:val="single" w:sz="6" w:color="2563EB"/>'
            f'<w:left   w:val="single" w:sz="6" w:color="2563EB"/>'
            f'<w:bottom w:val="single" w:sz="6" w:color="2563EB"/>'
            f'<w:right  w:val="single" w:sz="6" w:color="2563EB"/>'
            f'<w:insideH w:val="single" w:sz="4" w:color="BFDBFE"/>'
            f'<w:insideV w:val="single" w:sz="4" w:color="BFDBFE"/>'
            f'</w:tblBorders>'
            f'<w:tblCellMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
            f'<w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
            f'</w:tblPr>'
            f'<w:tr>{hdr_cells}</w:tr>{body}</w:tbl>{wempty()}')

def wcover(title_en, title_ar, meta):
    p = []
    p.append(f'<w:p><w:pPr><w:pBdr><w:top w:val="single" w:sz="12" w:space="1" w:color="1D3461"/></w:pBdr><w:spacing w:after="80"/></w:pPr></w:p>')
    p.append(wp(title_en, bold=True, sz=36, color="1D3461", align="center", after=60))
    if title_ar:
        p.append(wp_ar(title_ar, sz=30, color="1E4D8C", after=80))
    for en, ar in meta:
        p.append(wp_runs([{"t": en, "sz": 20, "b": True, "co": "1D3461"},
                          {"t": "  |  " + ar if ar else "", "sz": 20, "co": "374151", "fn": "Amiri", "rtl": bool(ar)}],
                         after=60))
    p.append(f'<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="1D3461"/></w:pBdr><w:spacing w:after="160"/></w:pPr></w:p>')
    return "".join(p)

# ── MCQ DOCX ──────────────────────────────────────────────────
def make_mcq_docx(data, show_ans, show_ar):
    p = []
    p.append(wcover("MCQ — " + data.get("subject",""), "أسئلة متعددة — " + data.get("subject_ar",""), []))
    for q in data.get("questions", []):
        p.append(wp_runs([
            {"t": f"Q{q.get('num','')}. ", "b": True, "sz": 24, "co": "1D3461"},
            {"t": q.get("question",""), "sz": 22},
            {"t": f"  [{q.get('difficulty','')}]", "sz": 18, "co": "888888", "i": True},
        ], after=60))
        if show_ar and q.get("question_ar"):
            p.append(wp_ar(q["question_ar"], sz=22, after=60))
        opts = q.get("options", {})
        if opts:
            rows = [[k, opts[k]] for k in sorted(opts.keys())]
            p.append(wtable(["Option", "Text"], ["الخيار", "النص"], rows, cws=[900, 8460]))
        if show_ans:
            corr = q.get("correct","")
            if corr:
                p.append(wp_runs([{"t":"✓ Answer: ","b":True,"sz":20,"co":"16A34A"},
                                   {"t":f"{corr} — {opts.get(corr,'')}","sz":20,"co":"16A34A"}], indent=360, after=40))
            if q.get("explanation"):
                p.append(wp(q["explanation"], italic=True, sz=20, color="374151", indent=360, after=60))
                if show_ar and q.get("explanation_ar"):
                    p.append(wp_ar(q["explanation_ar"], sz=20, after=60))
        p.append(wempty())
    return build_docx("".join(p))

# ── EXAM DOCX ─────────────────────────────────────────────────
def make_exam_docx(data, incl_ms, show_ar):
    p = []
    p.append(wcover(
        data.get("title","EXAMINATION"),
        data.get("title_ar","اختبار"),
        [
            (f"Institution: {data.get('institution','')}", data.get("institution_ar","")),
            (f"Course: {data.get('course','')}  |  Duration: {data.get('duration','')}  |  Total: {data.get('total_marks','')} marks",
             f"المادة: {data.get('course_ar','')}  |  المدة: {data.get('duration','')}  |  المجموع: {data.get('total_marks','')}"),
            (f"Lecturer: {data.get('lecturer','')}", data.get("lecturer_ar","")),
            (f"Date: {data.get('date','_______________')}", f"التاريخ: {data.get('date','_______________')}"),
            ("Student Name: ___________________________", "اسم الطالب: ___________________________"),
        ]
    ))
    p.append(wp("Instructions | التعليمات", bold=True, sz=26, color="1D3461", after=80))
    for inst in data.get("instructions", []):
        en = inst.get("en","") if isinstance(inst,dict) else str(inst)
        ar = inst.get("ar","") if isinstance(inst,dict) else ""
        p.append(wp_runs([{"t":"• ","b":True,"co":"2563EB","sz":22},
                           {"t":en,"sz":22},
                           {"t":"  |  "+ar if ar and show_ar else "","sz":20,"co":"374151","fn":"Amiri","rtl":bool(ar)}], after=50))
    p.append(wempty())

    for sec in data.get("sections", []):
        p.append(whr())
        name = sec.get("name",""); name_ar = sec.get("name_ar","")
        desc = sec.get("description",""); mks = sec.get("marks","")
        p.append(wp_runs([
            {"t": f"{name}  —  {desc}  ({mks} marks)", "b": True, "sz": 26, "co": "1D3461"},
        ], after=60))
        if show_ar and name_ar:
            p.append(wp_ar(f"{name_ar}  —  {sec.get('description_ar','')}  ({mks} درجة)", sz=22, after=80))
        ins = sec.get("instructions",""); ins_ar = sec.get("instructions_ar","")
        if ins:
            p.append(wp(ins + ("  |  " + ins_ar if ins_ar and show_ar else ""), italic=True, sz=20, color="555555", after=100))

        stype = sec.get("type","short_answer")

        # Matching — render as real table
        if stype == "matching" and sec.get("col_a") and sec.get("col_b"):
            rows = list(zip(
                [f"{i+1}. {v}" for i,v in enumerate(sec["col_a"])],
                [f"{chr(65+i)}. {v}" for i,v in enumerate(sec["col_b"])]
            ))
            p.append(wtable(["Column A | العمود أ","Column B | العمود ب"],["",""],rows,cws=[4680,4680]))

        for q in sec.get("questions", []):
            qn = q.get("num",""); qt = q.get("text",""); qar = q.get("text_ar",""); qmk = q.get("marks","")
            p.append(wp_runs([
                {"t": f"Q{qn}.  ", "b": True, "sz": 24, "co": "1D3461"},
                {"t": qt, "sz": 22},
                {"t": f"  [{qmk} marks]", "b": True, "sz": 20, "co": "2563EB", "i": True},
            ], after=60))
            if show_ar and qar:
                p.append(wp_ar(qar, sz=20, after=60))

            if stype == "mcq" and q.get("options"):
                rows = [[k, q["options"][k]] for k in sorted(q["options"].keys())]
                p.append(wtable(["","Option | الخيار"],["",""],rows,cws=[720,8640]))
            elif q.get("parts"):
                for pt in q["parts"]:
                    p.append(wp_runs([
                        {"t": f"  ({pt.get('part','')})  ", "b": True, "sz": 22, "co": "2563EB"},
                        {"t": pt.get("text",""), "sz": 22},
                        {"t": f"  [{pt.get('marks','')} marks]", "sz": 18, "co": "888888", "i": True},
                    ], indent=360, after=50))
                    if show_ar and pt.get("text_ar"):
                        p.append(wp_ar(pt["text_ar"], sz=20, after=50))
                    p.append(wans(pt.get("answer_lines", 3)))
            elif stype not in ("mcq", "matching"):
                p.append(wans(q.get("answer_lines", 4)))

        p.append(wempty())

    if incl_ms:
        p.append(wp("─"*80, sz=8, color="CCCCCC"))
        p.append(wempty())
        p.append(wcover("MARK SCHEME", "مخطط التصحيح", [(data.get("title",""),"")]))
        for sec in data.get("sections",[]):
            p.append(wp(sec.get("name",""), bold=True, sz=24, color="1D3461", after=80))
            for q in sec.get("questions",[]):
                p.append(wp_runs([
                    {"t":f"Q{q.get('num','')}  [{q.get('marks','')} marks]:  ","b":True,"sz":22,"co":"1D3461"},
                    {"t":q.get("model_answer",""),"sz":22},
                ], after=60))
                if show_ar and q.get("model_answer_ar"):
                    p.append(wp_ar(q["model_answer_ar"], sz=20, after=60))
                if q.get("correct"):
                    corr = q["correct"]; opts = q.get("options",{})
                    p.append(wp_runs([{"t":f"   ✓ {corr}: {opts.get(corr,'')}","b":True,"sz":20,"co":"16A34A"}], indent=360, after=40))
                if q.get("parts"):
                    for pt in q["parts"]:
                        p.append(wp_runs([
                            {"t":f"   ({pt.get('part','')}) ","b":True,"sz":20,"co":"2563EB"},
                            {"t":pt.get("model_answer",""),"sz":20},
                        ], indent=360, after=40))
                p.append(wempty())
    return build_docx("".join(p))

# ── SUMMARY DOCX ──────────────────────────────────────────────
def make_summary_docx(data, show_ar):
    p = []
    p.append(wcover(data.get("title","Summary"), data.get("title_ar","ملخص"),
                    [(data.get("subject",""), data.get("subject_ar",""))]))
    p.append(wp("Overview | نظرة عامة", bold=True, sz=26, color="1D3461", after=80))
    p.append(wp(data.get("overview",""), sz=22, after=80))
    if show_ar and data.get("overview_ar"):
        p.append(wp_ar(data["overview_ar"], sz=22, after=120))

    for sec in data.get("sections",[]):
        p.append(wp(sec.get("heading",""), bold=True, sz=26, color="1E4D8C", after=70))
        if show_ar and sec.get("heading_ar"):
            p.append(wp_ar(sec["heading_ar"], sz=24, color="2563EB", after=60))
        p.append(wp(sec.get("summary",""), sz=22, after=70))
        if show_ar and sec.get("summary_ar"):
            p.append(wp_ar(sec["summary_ar"], sz=20, after=80))
        kps = sec.get("key_points",[]); kps_ar = (sec.get("key_points_ar",[]) or []) + [""]*20
        for kp, kp_ar in zip(kps, kps_ar):
            line = f"▸  {kp}" + (f"  |  {kp_ar}" if kp_ar and show_ar else "")
            p.append(wp(line, sz=22, color="2563EB", after=40))
        p.append(wempty())

    if data.get("key_terms"):
        p.append(wp("Key Terms & Definitions | المصطلحات والتعريفات", bold=True, sz=26, color="1D3461", after=80))
        rows = [[kt.get("term",""), kt.get("term_ar",""), kt.get("definition",""), kt.get("definition_ar","")] for kt in data["key_terms"]]
        p.append(wtable(["Term","المصطلح","Definition","التعريف"],["","","",""],rows,cws=[1560,1560,3120,3120]))

    if data.get("formulas"):
        p.append(wempty())
        p.append(wp("Key Formulas | الصيغ الأساسية", bold=True, sz=26, color="1D3461", after=80))
        rows = [[f.get("name",""), f.get("name_ar",""), f.get("formula",""), f.get("meaning",""), f.get("meaning_ar","")] for f in data["formulas"]]
        p.append(wtable(["Name","الاسم","Formula / الصيغة","Meaning","المعنى"],["","","","",""],rows,cws=[1560,1560,2200,2200,1840]))

    if data.get("key_conclusions"):
        p.append(wempty())
        p.append(wp("Key Conclusions | الاستنتاجات", bold=True, sz=26, color="1D3461", after=80))
        concs_ar = (data.get("key_conclusions_ar",[]) or []) + [""]*20
        for c, c_ar in zip(data["key_conclusions"], concs_ar):
            line = f"✓  {c}" + (f"  |  {c_ar}" if c_ar and show_ar else "")
            p.append(wp(line, bold=True, sz=22, color="16A34A", after=50))
    return build_docx("".join(p))

# ══════════════════════════════════════════════════════════════
# PPTX BUILDER — pure zipfile + xml, native tables
# ══════════════════════════════════════════════════════════════
def emu(i): return int(i * 914400)
def pt_(p): return int(p * 12700)
SW = emu(13.33); SH = emu(7.5)

def xe(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;").replace("'","&apos;")

def rect(n, x, y, w, h, fill, ln=None, lw=0):
    line = (f'<a:ln w="{pt_(lw)}"><a:solidFill><a:srgbClr val="{ln}"/></a:solidFill></a:ln>'
            if ln and lw > 0 else '<a:ln><a:noFill/></a:ln>')
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="1" name="{n}"/>'
            f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/>'
            f'<a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{line}</p:spPr></p:sp>')

def para(text, sz, bold=False, italic=False, col="FFFFFF", alg="l", font="Calibri", bul="", spc=0, rtl=False):
    a = {"l":"l","c":"ctr","r":"r"}.get(alg,"l")
    b = ' b="1"' if bold else ""; i = ' i="1"' if italic else ""
    bul_xml = f'<a:buChar char="{xe(bul)}"/>' if bul else "<a:buNone/>"
    sa = f'<a:spcAft><a:spcPts val="{spc*100}"/></a:spcAft>' if spc else ""
    rtl_attr = ' rtl="1"' if rtl else ""
    return (f'<a:p><a:pPr algn="{a}" indent="0" marL="228600"{rtl_attr}>{bul_xml}{sa}</a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="{int(sz*100)}" dirty="0"{b}{i}>'
            f'<a:solidFill><a:srgbClr val="{col}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/></a:rPr>'
            f'<a:t>{xe(text)}</a:t></a:r></a:p>')

def bilpara(en, ar, sz_e=16, sz_a=13, col_e="FFFFFF", col_a="94A3B8"):
    """Single paragraph with English then Arabic."""
    return (f'<a:p><a:pPr algn="l" indent="0" marL="228600"><a:buNone/></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="{int(sz_e*100)}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{col_e}"/></a:solidFill>'
            f'<a:latin typeface="Calibri"/></a:rPr><a:t>{xe(en)}</a:t></a:r>'
            f'<a:r><a:rPr lang="ar-SA" sz="{int(sz_a*100)}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{col_a}"/></a:solidFill>'
            f'<a:latin typeface="Amiri"/></a:rPr>'
            f'<a:t>  |  {xe(ar)}</a:t></a:r></a:p>')

def txbox(n, x, y, w, h, paras):
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="2" name="{n}"/>'
            f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/>'
            f'<a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            f'<p:txBody><a:bodyPr wrap="square" lIns="45720" rIns="45720" tIns="45720" bIns="45720"/>'
            f'<a:lstStyle/>{"".join(paras)}</p:txBody></p:sp>')

def slide_xml(shapes, bg):
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            f' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
            f' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            f'<p:cSld><p:bg><p:bgPr>'
            f'<a:solidFill><a:srgbClr val="{bg}"/></a:solidFill>'
            f'<a:effectLst/></p:bgPr></p:bg><p:spTree>'
            f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{SW}" cy="{SH}"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="{SW}" cy="{SH}"/></a:xfrm></p:grpSpPr>'
            f'{"".join(shapes)}</p:spTree></p:cSld></p:sld>')

def slide_hdr(t_en, t_ar, T):
    shapes = [
        rect("top", 0, 0, 13.33, 0.13, T["ac"]),
        rect("tb",  0.4, 0.22, 12.53, 0.92, T["tb"]),
        txbox("te", 0.55, 0.24, 8.5,  0.88, [para(t_en, 26, bold=True, col=T["tx"], font="Georgia")]),
    ]
    if t_ar:
        shapes.append(txbox("ta", 9.1, 0.24, 3.6, 0.88,
                            [para(t_ar, 17, bold=True, col=T["ac"], alg="r", font="Amiri", rtl=True)]))
    return shapes

def pptx_native_table(h_en, h_ar, rows, T, x=0.4, y=1.3, w=12.53):
    """Native PPTX table — renders as real editable table in PowerPoint."""
    n = len(h_en)
    cw = int(emu(w) / n)
    rh = emu(0.44)

    def tc_hdr(en, ar):
        return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
                f'<a:p><a:pPr algn="ctr"/>'
                f'<a:r><a:rPr lang="en-US" b="1" sz="1200" dirty="0">'
                f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
                f'<a:latin typeface="Calibri"/></a:rPr><a:t>{xe(en)}</a:t></a:r>'
                f'{"<a:r><a:rPr lang=\"ar-SA\" sz=\"1000\" dirty=\"0\"><a:solidFill><a:srgbClr val=\"BAD0FF\"/></a:solidFill><a:latin typeface=\"Amiri\"/></a:rPr><a:t> | "+xe(ar)+"</a:t></a:r>" if ar else ""}'
                f'</a:p></a:txBody>'
                f'<a:tcPr><a:solidFill><a:srgbClr val="{T["ac"]}"/></a:solidFill></a:tcPr></a:tc>')

    def tc_data(val, bg):
        return (f'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>'
                f'<a:p><a:pPr algn="ctr"/>'
                f'<a:r><a:rPr lang="en-US" sz="1100" dirty="0">'
                f'<a:solidFill><a:srgbClr val="{T["su"]}"/></a:solidFill>'
                f'<a:latin typeface="Calibri"/></a:rPr>'
                f'<a:t>{xe(str(val))}</a:t></a:r></a:p></a:txBody>'
                f'<a:tcPr><a:solidFill><a:srgbClr val="{bg}"/></a:solidFill></a:tcPr></a:tc>')

    hdr_row = "".join(tc_hdr(h_en[i], h_ar[i] if i < len(h_ar) else "") for i in range(n))
    data_rows = ""
    for ri, row in enumerate(rows):
        bg = T["bx"] if ri % 2 == 0 else T["tb"]
        cells = "".join(tc_data(row[ci] if ci < len(row) else "", bg) for ci in range(n))
        data_rows += f'<a:tr h="{rh}">{cells}</a:tr>'

    tbl = (f'<a:tbl>'
           f'<a:tblPr firstRow="1" bandRow="1"/>'
           f'<a:tblGrid>{"".join(f"<a:gridCol w=\"{cw}\"/>" for _ in range(n))}</a:tblGrid>'
           f'<a:tr h="{rh}">{hdr_row}</a:tr>'
           f'{data_rows}</a:tbl>')

    h_total = rh * (len(rows) + 1)
    return (f'<p:graphicFrame>'
            f'<p:nvGraphicFramePr>'
            f'<p:cNvPr id="10" name="tbl"/>'
            f'<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>'
            f'<p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/>'
            f'<a:ext cx="{emu(w)}" cy="{h_total}"/></p:xfrm>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f'{tbl}</a:graphicData></a:graphic></p:graphicFrame>')

# ── Slide type builders ───────────────────────────────────────
def s_title(s, T):
    en = s.get("title",""); ar = s.get("title_ar","")
    sub_en = s.get("subtitle",""); sub_ar = s.get("subtitle_ar","")
    body = s.get("body","")
    return slide_xml([
        rect("bg",  0, 0, 13.33, 7.5,  T["bg"]),
        rect("bar", 0, 0, 0.24,  7.5,  T["ac"]),
        rect("bot", 0.5, 7.1, 4.5, 0.06, T["ac"]),
        txbox("te", 0.5, 1.0, 12.3, 2.0, [para(en, 42, bold=True, col=T["tx"], font="Georgia")]),
        txbox("ta", 0.5, 3.1, 12.3, 0.85,
              [para(ar, 24, bold=True, col=T["ac"], alg="r", font="Amiri", rtl=True)] if ar else
              [para("", 12, col=T["bg"])]),
        txbox("se", 0.5, 4.05, 12.3, 0.65, [para(sub_en, 20, col=T["su"])]),
        txbox("sa", 0.5, 4.72, 12.3, 0.55,
              [para(sub_ar, 15, col=T["ac"], alg="r", font="Amiri", rtl=True)] if sub_ar else
              [para("", 12, col=T["bg"])]),
        txbox("au", 0.5, 5.5, 12.3, 0.5, [para(body, 13, italic=True, col=T["su"])]),
    ], T["bg"])

def s_bullets(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    bullets = s.get("bullets", [])
    bullets_ar = (s.get("bullets_ar") or []) + [""] * 20
    ps = []
    for en_b, ar_b in zip(bullets, bullets_ar):
        if ar_b:
            ps.append(bilpara(en_b, ar_b, col_e=T["su"], col_a=T["ac"]))
        else:
            ps.append(para(en_b, 17, col=T["su"], bul="▸", spc=3))
    sh.append(txbox("bd", 0.55, 1.32, 12.2, 5.9, ps))
    return slide_xml(sh, T["bg"])

def s_lecture(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    ps = []
    def sec(lbl_en, lbl_ar, content, content_ar="", mono=False):
        if not content: return
        ps.append(bilpara(lbl_en, lbl_ar, sz_e=13, sz_a=11, col_e=T["ac"], col_a=T["ac"]))
        fn = "Courier New" if mono else "Calibri"
        col = T["mn"] if mono else T["su"]
        for line in str(content).split("\\n"):
            ps.append(para(line.strip() or " ", 15 if mono else 16, col=col, font=fn, spc=1))
        if content_ar and not mono:
            ps.append(para(content_ar, 13, col=T["ac"], alg="r", font="Amiri", rtl=True, spc=1))
        ps.append(para(" ", 5, col=T["bg"]))
    sec("📖 THEORY", "النظرية", s.get("theory",""), s.get("theory_ar",""))
    sec("📐 FORMULA", "الصيغة",  s.get("formula",""), mono=True)
    if s.get("where"):
        ps.append(para("   " + s["where"], 13, italic=True, col=T["su"]))
        if s.get("where_ar"): ps.append(para("   " + s["where_ar"], 12, italic=True, col=T["ac"], alg="r", font="Amiri", rtl=True))
        ps.append(para(" ", 5, col=T["bg"]))
    sec("✏️ EXAMPLE", "مثال", s.get("example",""), s.get("example_ar",""))
    sec("🔢 CALCULATION", "الحساب", s.get("calculation",""), mono=True)
    sh.append(txbox("bd", 0.55, 1.32, 12.2, 5.9, ps))
    return slide_xml(sh, T["bg"])

def s_formula(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    sh.append(rect("fb", 1.5, 1.38, 10.3, 1.38, T["tb"], T["ac"], 1.2))
    sh.append(txbox("fm", 1.65, 1.41, 9.95, 1.32,
                    [para(s.get("formula",""), 30, bold=True, col=T["tx"], font="Courier New", alg="c")]))
    ps = []
    if s.get("where"):
        ps.append(para("Where:", 14, bold=True, col=T["ac"]))
        for part in str(s["where"]).split(","):
            ps.append(para(part.strip(), 14, col=T["su"], bul="•"))
        if s.get("where_ar"):
            ps.append(para(s["where_ar"], 13, col=T["ac"], alg="r", font="Amiri", rtl=True))
        ps.append(para(" ", 5, col=T["bg"]))
    if s.get("example"):
        ps.append(bilpara("Worked Example:", "مثال محلول:", sz_e=13, sz_a=12, col_e=T["ac"], col_a=T["ac"]))
        ps.append(para(s["example"], 14, col=T["su"]))
        if s.get("example_ar"): ps.append(para(s["example_ar"], 13, col=T["ac"], alg="r", font="Amiri", rtl=True))
        ps.append(para(" ", 5, col=T["bg"]))
    if s.get("calculation"):
        ps.append(bilpara("Solution:", "الحل:", sz_e=13, sz_a=12, col_e=T["ac"], col_a=T["ac"]))
        for line in str(s["calculation"]).split("\\n"):
            ps.append(para(line.strip() or " ", 14, col=T["mn"], font="Courier New"))
    if ps:
        sh.append(txbox("rs", 0.55, 2.9, 12.2, 4.35, ps))
    return slide_xml(sh, T["bg"])

def s_two_col(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    sh += [
        rect("lc", 0.35, 1.28, 6.1, 5.88, T["bx"], T["ac"], 0.7),
        rect("lh", 0.35, 1.28, 6.1, 0.48, T["ac"]),
        txbox("lt", 0.45, 1.30, 5.9, 0.44,
              [para(s.get("left_title",""), 13, bold=True, col="FFFFFF", alg="c")]),
        txbox("lb", 0.45, 1.84, 5.9, 5.22,
              [para(b, 15, col=T["su"], bul="▸", spc=3) for b in s.get("left_points",[])]),
        rect("rc", 6.88, 1.28, 6.1, 5.88, T["bx"], T["ac"], 0.7),
        rect("rh", 6.88, 1.28, 6.1, 0.48, T["ac"]),
        txbox("rt", 6.98, 1.30, 5.9, 0.44,
              [para(s.get("right_title",""), 13, bold=True, col="FFFFFF", alg="c")]),
        txbox("rb", 6.98, 1.84, 5.9, 5.22,
              [para(b, 15, col=T["su"], bul="▸", spc=3) for b in s.get("right_points",[])]),
    ]
    return slide_xml(sh, T["bg"])

def s_stat(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    xp = [0.35, 4.6, 8.85]; ww = 3.9
    for i, si in enumerate(s.get("stats",[])[:3]):
        x = xp[i]
        sh += [
            rect(f"sc{i}", x,      1.32, ww,    3.65, T["bx"], T["ac"], 0.8),
            rect(f"sl{i}", x,      1.32, ww,    0.12, T["ac"]),
            txbox(f"sv{i}", x+0.1, 1.57, ww-0.2, 1.38,
                  [para(si.get("value",""), 50, bold=True, col=T["ac"], font="Georgia", alg="c")]),
            txbox(f"lb{i}", x+0.1, 3.0,  ww-0.2, 0.65,
                  [para(si.get("label",""), 13, col=T["su"], alg="c")]),
        ]
        if si.get("label_ar"):
            sh.append(txbox(f"la{i}", x+0.1, 3.65, ww-0.2, 0.48,
                            [para(si["label_ar"], 12, col=T["ac"], alg="c", font="Amiri", rtl=True)]))
    if s.get("body"):
        sh.append(txbox("bd", 0.5, 5.1, 12.3, 1.0,
                        [para(s["body"], 14, italic=True, col=T["su"])]))
    return slide_xml(sh, T["bg"])

def s_table(s, T):
    sh = slide_hdr(s.get("title",""), s.get("title_ar",""), T)
    h_en = s.get("headers",  []); h_ar = s.get("headers_ar", [])
    rows = s.get("rows", [])
    if h_en and rows:
        sh.append(pptx_native_table(h_en, h_ar, rows, T, x=0.4, y=1.32, w=12.53))
    else:
        sh.append(txbox("nb", 0.5, 1.5, 12.0, 1.0,
                        [para("No table data available.", 16, col=T["su"])]))
    return slide_xml(sh, T["bg"])

def s_conclusion(s, T):
    sh = slide_hdr(s.get("title","Key Takeaways"), s.get("title_ar","النقاط الرئيسية"), T)
    sh += [
        rect("cb", 0.4,  1.28, 12.53, 5.92, T["bx"], T["ac"], 0.8),
        rect("bb", 0, 7.36, 13.33, 0.14, T["ac"]),
    ]
    bullets    = s.get("bullets", [])
    bullets_ar = (s.get("bullets_ar") or []) + [""] * 20
    ps = []
    for en_b, ar_b in zip(bullets, bullets_ar):
        if ar_b:
            ps.append(bilpara(f"✓  {en_b}", ar_b, col_e=T["su"], col_a=T["ac"]))
        else:
            ps.append(para(en_b, 17, col=T["su"], bul="✓", spc=5))
    sh.append(txbox("bd", 0.65, 1.46, 12.0, 5.72, ps))
    return slide_xml(sh, T["bg"])

def build_slide(s, T):
    t = s.get("type","bullets")
    if   t == "title":        return s_title(s, T)
    elif t == "lecture":      return s_lecture(s, T)
    elif t == "formula":      return s_formula(s, T)
    elif t == "two_col":      return s_two_col(s, T)
    elif t == "stat_callout": return s_stat(s, T)
    elif t == "table":        return s_table(s, T)
    elif t == "conclusion":   return s_conclusion(s, T)
    else:                     return s_bullets(s, T)

def assemble_pptx(xmls):
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        n = len(xmls)
        ct_s = "".join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(n))
        z.writestr("[Content_Types].xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>{ct_s}</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
        z.writestr("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Kaabe</Application></Properties>')
        z.writestr("ppt/theme/theme1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Kaabe"><a:themeElements><a:clrScheme name="Kaabe"><a:dk1><a:sysClr lastClr="000000" val="windowText"/></a:dk1><a:lt1><a:sysClr lastClr="FFFFFF" val="window"/></a:lt1><a:dk2><a:srgbClr val="1D3461"/></a:dk2><a:lt2><a:srgbClr val="EFF6FF"/></a:lt2><a:accent1><a:srgbClr val="2563EB"/></a:accent1><a:accent2><a:srgbClr val="0D9488"/></a:accent2><a:accent3><a:srgbClr val="A855F7"/></a:accent3><a:accent4><a:srgbClr val="F59E0B"/></a:accent4><a:accent5><a:srgbClr val="EF4444"/></a:accent5><a:accent6><a:srgbClr val="10B981"/></a:accent6><a:hlink><a:srgbClr val="2563EB"/></a:hlink><a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink></a:clrScheme><a:fontScheme name="Kaabe"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface="Amiri"/><a:cs typeface="Amiri"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface="Amiri"/><a:cs typeface="Amiri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>')
        z.writestr("ppt/slideMasters/slideMaster1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle><a:lstStyle/></p:titleStyle><p:bodyStyle><a:lstStyle/></p:bodyStyle><p:otherStyle><a:lstStyle/></p:otherStyle></p:txStyles></p:sldMaster>')
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/></Relationships>')
        z.writestr("ppt/slideLayouts/slideLayout1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>')
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
        ids = "\n".join(f'<p:sldId id="{256+i}" r:id="rId{i+3}"/>' for i in range(n))
        prs_rels = "\n".join(f'<Relationship Id="rId{i+3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(n))
        for i, xml in enumerate(xmls):
            z.writestr(f"ppt/slides/slide{i+1}.xml", xml)
            z.writestr(f"ppt/slides/_rels/slide{i+1}.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
        z.writestr("ppt/presentation.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="{SW}" cy="{SH}" type="custom"/><p:notesSz cx="{emu(7.5)}" cy="{emu(10)}"/></p:presentation>')
        z.writestr("ppt/_rels/presentation.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>{prs_rels}</Relationships>')
    buf.seek(0)
    return buf.read()

# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
for k, v in {
    "tool":"slides","pdf_b64":None,"filename":"",
    "chat":[],"fc_idx":0,"fc_show":False,
    "mcq_data":None,"sum_data":None,"exam_data":None,"fc_data":None,
    "exam_sections":[
        {"name":"Section A","name_ar":"القسم أ","type":"mcq",         "marks":20,"num_q":10,"desc":"Multiple choice questions","desc_ar":"أسئلة الاختيار من متعدد"},
        {"name":"Section B","name_ar":"القسم ب","type":"short_answer","marks":30,"num_q":5, "desc":"Short answer questions","desc_ar":"أسئلة الإجابة القصيرة"},
        {"name":"Section C","name_ar":"القسم ج","type":"long_answer", "marks":30,"num_q":3, "desc":"Long answer questions","desc_ar":"أسئلة الإجابة المطولة"},
    ]
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    # ── Language first — everything else depends on it ──────
    st.markdown("### 🌐 Language")
    lang_mode = st.radio("", ["🇬🇧  English", "🇸🇦  Arabic"], label_visibility="collapsed")
    ar_only   = lang_mode == "🇸🇦  Arabic"
    st.markdown("---")

    # ── API Key ──────────────────────────────────────────────
    st.markdown("### 🔑 " + ("مفتاح API — Gemini" if ar_only else "Gemini API Key"))
    st.markdown(
        '<div class="sbox">' +
        ("مجاني على " if ar_only else "Free at ") + '<a href="https://aistudio.google.com/app/apikey" target="_blank" style="color:#60a5fa">aistudio.google.com</a>' + ("</div>" if ar_only else "<br>No credit card needed</div>"),
        unsafe_allow_html=True)
    api_key = st.text_input("", "", type="password", placeholder="AIzaSy...", label_visibility="collapsed")
    st.markdown("---")

    # ── Upload ───────────────────────────────────────────────
    st.markdown("### 📄 " + ("رفع ملف PDF" if ar_only else "Upload PDF"))
    uploaded = st.file_uploader("", type=["pdf"], label_visibility="collapsed")
    if uploaded:
        b64 = base64.standard_b64encode(uploaded.getvalue()).decode()
        st.session_state.pdf_b64 = b64
        st.session_state.filename = uploaded.name
        kb = len(uploaded.getvalue()) // 1024
        st.markdown(
            f'<div style="background:#071a0e;border:1px solid #14532d;border-radius:8px;padding:.55rem .9rem;font-size:.78rem;color:#4ade80">' +
            f"✓ {uploaded.name}<br>" +
            f'<span style="color:#5a7099">{kb} KB</span></div>',
            unsafe_allow_html=True)
    st.markdown("---")

    # ── Theme ────────────────────────────────────────────────
    st.markdown("### 🎨 " + ("سمة الشرائح" if ar_only else "Slide Theme"))
    sel_theme = st.selectbox("", list(THEMES.keys()), label_visibility="collapsed")
    st.markdown("---")

    # ── Settings ─────────────────────────────────────────────
    st.markdown("### ⚙️ " + ("الإعدادات" if ar_only else "Settings"))
    num_slides  = st.slider("Slides" if not ar_only else "عدد الشرائح",  6, 20, 10)
    num_mcq     = st.slider("MCQ Qs" if not ar_only else "عدد الأسئلة",  5, 30, 15)
    num_fc      = st.slider("Flashcards" if not ar_only else "البطاقات",  5, 25, 12)
    difficulty  = st.selectbox(
        "Exam difficulty" if not ar_only else "مستوى الاختبار",
        ["Easy","Medium","Hard","Mixed"] if not ar_only else ["سهل","متوسط","صعب","مختلط"])
    course_name = st.text_input(
        "Course" if not ar_only else "اسم المادة",
        placeholder="e.g. BBA Mathematics" if not ar_only else "مثال: الرياضيات")
    institution = st.text_input(
        "Institution" if not ar_only else "المؤسسة",
        placeholder="e.g. University of Somalia" if not ar_only else "مثال: جامعة الصومال")
    lecturer    = st.text_input(
        "Lecturer" if not ar_only else "اسم المحاضر",
        placeholder="e.g. Dr. Ahmed" if not ar_only else "مثال: د. أحمد")
    focus_topic = st.text_input(
        "Focus topic" if not ar_only else "موضوع التركيز",
        placeholder="e.g. probability, forces" if not ar_only else "مثال: الاحتمالات، القوى")

# ══════════════════════════════════════════════════════════════
# HERO
# ══════════════════════════════════════════════════════════════
fn = st.session_state.filename
if ar_only:
    _hero_title    = "كعبه"
    _hero_subtitle = "منصة الذكاء الاصطناعي الأكاديمية"
    _hero_desc     = "شرائح · أسئلة · اختبار · ملخص · بطاقات · معلم ذكي"
    _no_pdf        = "لا يوجد ملف"
    _lang_pill     = "عربي"
else:
    _hero_title    = "Kaabe"
    _hero_subtitle = "Academic AI Platform"
    _hero_desc     = "PDF → Slides · MCQ · Exam · Summary · Flashcards · AI Tutor"
    _no_pdf        = "No PDF"
    _lang_pill     = "English"

st.markdown(f"""
<div class="hero">
  <div class="hero-ico">🎓</div>
  <div>
    <h1>{_hero_title}</h1>
    <p style="font-size:.95rem;color:#4a9ede;margin:0 0 .2rem;{"font-family:'Amiri',serif;direction:rtl;" if ar_only else ""}">{_hero_subtitle}</p>
    <p>{_hero_desc}</p>
  </div>
  <div class="hero-right">
    <span class="pill">📄 <b>{(fn[:20]+"…") if len(fn)>22 else fn or _no_pdf}</b></span>
    <span class="pill">🌐 <b>{_lang_pill}</b></span>
    <span class="pill">🎨 <b>{sel_theme}</b></span>
  </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# TOOL NAV
# ══════════════════════════════════════════════════════════════
TOOLS = [
    ("slides",     "📊", "Lecture Slides",  "شرائح المحاضرة"),
    ("mcq",        "❓", "MCQ Generator",   "أسئلة متعددة"),
    ("exam",       "📝", "Exam Paper",      "ورقة اختبار"),
    ("summary",    "📋", "Smart Summary",   "ملخص ذكي"),
    ("flashcards", "🃏", "Flashcards",      "بطاقات دراسية"),
    ("qa",         "💬", "AI Tutor",        "معلم ذكي"),
]
tc = st.columns(len(TOOLS))
for i, (tid, ico, en_l, ar_l) in enumerate(TOOLS):
    with tc[i]:
        lbl = f"{ico} {ar_l}" if ar_only else f"{ico} {en_l}"
        if st.button(lbl, key=f"t_{tid}", use_container_width=True,
                     type="primary" if st.session_state.tool == tid else "secondary"):
            st.session_state.tool = tid; st.rerun()
st.markdown("---")

# ── Helpers ───────────────────────────────────────────────────
def check_ready():
    if not api_key:
        st.warning("🔑 أدخل مفتاح Gemini في الشريط الجانبي." if ar_only else "🔑 Enter your Gemini API key in the sidebar.")
        return False
    if not st.session_state.pdf_b64:
        st.warning("📄 ارفع ملف PDF من الشريط الجانبي." if ar_only else "📄 Upload a PDF in the sidebar.")
        return False
    return True

def prog(se, pe, msg, pct):
    se.markdown(f'<div style="text-align:center;font-size:.86rem;color:#5a7099;margin:.35rem 0">{msg}</div>', unsafe_allow_html=True)
    pe.markdown(f'<div class="pb"><div class="pf" style="width:{pct}%"></div></div>', unsafe_allow_html=True)

def lang_instr():
    if ar_only: return "Provide ALL content in Arabic only. Every field must be in Arabic."
    return "Provide all content in English only."

tool = st.session_state.tool

# ══════════════════════════════════════════════════════════════
# 1 · LECTURE SLIDES
# ══════════════════════════════════════════════════════════════
if tool == "slides":
    st.markdown("## 📊 " + ("شرائح المحاضرة" if ar_only else "Lecture Slides"))
    c1, c2 = st.columns([2, 1])
    with c1:
        pres_style = st.selectbox("Style", ["University Lecture","Research Paper","Technical Report","Business Analysis"])
        incl_f = st.checkbox("Extract formulas & equations", value=True)
        incl_e = st.checkbox("Include worked examples",      value=True)
        incl_t = st.checkbox("Extract tables (real tables)", value=True)
    with c2:
        if ar_only:
            st.markdown('<div class="sbox"><b>أنواع الشرائح:</b><br>📖 نظرية<br>📐 صيغة رياضية<br>✏️ مثال محلول<br>🔢 خطوات الحساب<br>📊 جدول حقيقي<br>🔄 عمودان<br>📈 إحصائيات<br>✓ الخلاصة</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="sbox"><b>Slide types generated:</b><br>📖 Theory explanation<br>📐 Formula + variables<br>✏️ Worked example<br>🔢 Step-by-step calculation<br>📊 Native table<br>🔄 Two-column comparison<br>📈 Stats callout<br>✓ Conclusion</div>', unsafe_allow_html=True)

    if st.button("✨ " + ("توليد الشرائح" if ar_only else "Generate Slides"), type="primary", use_container_width=True):
        if not check_ready(): st.stop()
        se = st.empty(); pe = st.empty()
        try:
            prog(se, pe, "🧠 AI reading PDF…", 20)
            fi = f"Focus on: {focus_topic}." if focus_topic else ""
            ci = f"Course: {course_name}." if course_name else ""
            # Language-specific labels for the prompt
            if ar_only:
                type_labels = {
                    "title":        "عنوان (صفحة العنوان فقط)",
                    "bullets":      "نقاط (حقائق ومعلومات رئيسية)",
                    "lecture":      "محاضرة (نظرية + صيغة + مثال + حساب)",
                    "formula":      "صيغة (معادلة رياضية + مثال محلول)",
                    "two_col":      "عمودان (مقارنة أو تصنيف)",
                    "stat_callout": "إحصائيات (أرقام وبيانات بارزة)",
                    "table":        "جدول (بيانات منظمة)",
                    "conclusion":   "خلاصة (النقاط الرئيسية والتوصيات)",
                }
                lang_note = "اكتب جميع النصوص باللغة العربية فقط. استخدم المصطلحات الأكاديمية الصحيحة."
                title_lbl      = "العنوان"
                subtitle_lbl   = "العنوان الفرعي"
                theory_lbl     = "الشرح النظري"
                formula_lbl    = "الصيغة الرياضية"
                where_lbl      = "تعريف المتغيرات"
                example_lbl    = "مثال رقمي"
                calc_lbl       = "خطوات الحل"
                bullets_lbl    = "النقاط الرئيسية"
                conc_title     = "الخلاصة والتوصيات"
                author_lbl     = "اسم المحاضر / المؤلف"
            else:
                type_labels = {
                    "title":        "title (cover slide only)",
                    "bullets":      "bullets (key facts and information)",
                    "lecture":      "lecture (theory + formula + example + calculation)",
                    "formula":      "formula (equation + worked example)",
                    "two_col":      "two_col (comparison or classification)",
                    "stat_callout": "stat_callout (key numbers and statistics)",
                    "table":        "table (structured data)",
                    "conclusion":   "conclusion (takeaways and recommendations)",
                }
                lang_note      = "Write ALL text in English only."
                title_lbl      = "title"
                subtitle_lbl   = "subtitle"
                theory_lbl     = "theory explanation"
                formula_lbl    = "exact formula"
                where_lbl      = "variable definitions"
                example_lbl    = "numerical example"
                calc_lbl       = "step-by-step solution"
                bullets_lbl    = "key points"
                conc_title     = "Key Takeaways"
                author_lbl     = "author/lecturer name"

            # Build a concrete slide list so Gemini knows EXACTLY how many to produce
            slide_list = []
            slide_list.append('{"slide_num":1,"type":"title","title":"...","subtitle":"...","body":"..."}')
            for i in range(2, num_slides):
                types = ["bullets","lecture","formula","two_col","table","stat_callout","bullets","lecture","formula","two_col","bullets","lecture","stat_callout","two_col","bullets"]
                t = types[(i-2) % len(types)]
                if t == "bullets":
                    slide_list.append(f'{{"slide_num":{i},"type":"bullets","title":"...","bullets":["...","...","...","...","..."]}}')
                elif t == "lecture":
                    slide_list.append(f'{{"slide_num":{i},"type":"lecture","title":"...","theory":"...","formula":"...","where":"...","example":"...","calculation":"step1 \\n step2 \\n answer"}}')
                elif t == "formula":
                    slide_list.append(f'{{"slide_num":{i},"type":"formula","title":"...","formula":"...","where":"...","example":"...","calculation":"step1 \\n step2 \\n answer"}}')
                elif t == "two_col":
                    slide_list.append(f'{{"slide_num":{i},"type":"two_col","title":"...","left_title":"...","left_points":["...","...","..."],"right_title":"...","right_points":["...","..."]}}')
                elif t == "table":
                    slide_list.append(f'{{"slide_num":{i},"type":"table","title":"...","headers":["Col1","Col2","Col3"],"rows":[["...","...","..."],["...","...","..."]]}}')
                elif t == "stat_callout":
                    slide_list.append(f'{{"slide_num":{i},"type":"stat_callout","title":"...","stats":[{{"value":"...","label":"..."}},{{"value":"...","label":"..."}},{{"value":"...","label":"..."}}],"body":"..."}}')
            slide_list.append(f'{{"slide_num":{num_slides},"type":"conclusion","title":"...","bullets":["...","...","...","...","..."]}}')
            slides_template = ",\n    ".join(slide_list)

            lang_note = (
                "LANGUAGE: Write ALL text in Arabic only. Use proper academic Arabic. Every field must be in Arabic."
                if ar_only else
                "LANGUAGE: Write ALL text in English only."
            )
            ci = f"Course: {course_name}." if course_name else ""
            fi = f"Pay special attention to: {focus_topic}." if focus_topic else ""

            prompt = f"""You are a senior university professor and expert slide designer.
Read the uploaded PDF completely and create EXACTLY {num_slides} lecture slides.
{ci} {fi}
{lang_note}

CRITICAL RULES — follow every one without exception:
1. You MUST produce EXACTLY {num_slides} slides — not fewer, not more.
2. Use ONLY real content from the PDF — exact facts, numbers, definitions, formulas.
3. NEVER say "from the document" or "the document says" — state facts directly.
4. NEVER leave placeholder text like "..." — replace every "..." with real content.
5. Every bullet must be a specific real fact with actual data, not vague.
6. For EVERY formula: write it exactly, define every variable, give a worked numerical example with full calculation steps.
7. For calculations: use \\n between steps (NOT a real newline — write the 2-character sequence backslash-n).
8. For tables: extract exact headers and all data rows from the PDF.
9. Cover ALL major topics in the document across the {num_slides} slides.
10. Last slide MUST be type "conclusion" with real key takeaways.

SLIDE COUNT: You must generate all {num_slides} slides in the "slides" array.
Do NOT stop early. Do NOT truncate. Complete all {num_slides} slides.

SLIDE TYPES available:
- "title": cover slide (slide 1 only)
- "bullets": key points list (5 detailed bullets)
- "lecture": theory + formula + worked example + calculation
- "formula": standalone equation + variable definitions + worked example
- "two_col": side-by-side comparison with 4 points each
- "table": real table with headers and rows
- "stat_callout": 3 key statistics with labels
- "conclusion": final takeaways (last slide only)

Return ONLY raw valid JSON. No markdown. No code fences. No explanation.
Strings must be on ONE line. Use the 2-character sequence \\n for line breaks in calculations.

{{
  "title": "exact document title here",
  "subtitle": "course or subject name",
  "author": "author or lecturer name if found",
  "slides": [
    {slides_template}
  ]
}}

REMINDER: Replace every single "..." with REAL content from the PDF.
You MUST generate ALL {num_slides} slides. This is mandatory."""
            raw  = call_gemini(api_key, prompt, st.session_state.pdf_b64, tokens=8000)
            prog(se, pe, "🔧 Parsing…", 50)
            plan = safe_parse(raw)
            prog(se, pe, "🎨 Building PPTX…", 65)
            T    = THEMES[sel_theme]
            # Build all slides in one pass — no per-slide UI updates for speed
            xmls = [build_slide(s, T) for s in plan.get("slides", [])]
            prog(se, pe, "💾 Assembling…", 92)
            pptx = assemble_pptx(xmls)
            se.empty(); pe.empty()
            st.success(f"✅ {len(xmls)} " + ("شريحة جاهزة!" if ar_only else "slides ready!"))
            safe = re.sub(r"[^a-zA-Z0-9_\- ]","",plan.get("title","slides"))[:40].strip().replace(" ","_") or "kaabe_slides"
            st.download_button("⬇️ " + ("تنزيل الشرائح" if ar_only else "Download Slides (.pptx)"), data=pptx, file_name=f"{safe}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               use_container_width=True)
            with st.expander("📋 " + ("مخطط الشرائح" if ar_only else "Slide outline")):
                for s in plan.get("slides",[]):
                    st.markdown(f"**{s.get('slide_num')}. {s.get('title','')}** `{s.get('type','')}`")
                    if s.get("formula"): st.code(s["formula"])
                    for b in s.get("bullets",[]): st.markdown(f"  - {b}")
        except Exception as e:
            se.empty(); pe.empty(); st.error(str(e))

# ══════════════════════════════════════════════════════════════
# 2 · MCQ
# ══════════════════════════════════════════════════════════════
elif tool == "mcq":
    st.markdown("## ❓ " + ("أسئلة متعددة" if ar_only else "MCQ Generator"))
    c1, c2, c3 = st.columns(3)
    with c1: mcq_diff = st.selectbox("Difficulty", ["Easy","Medium","Hard","Mixed"])
    with c2: mcq_type = st.selectbox("Type", ["Conceptual","Calculation-based","Mixed"])
    with c3: show_ans = st.checkbox("Show answers & explanations", value=True)

    if st.button("🎯 " + ("توليد الأسئلة" if ar_only else "Generate MCQ"), type="primary", use_container_width=True):
        if not check_ready(): st.stop()
        se = st.empty(); pe = st.empty()
        try:
            prog(se, pe, "🧠 Generating questions…", 35)
            fi = f"Focus on: {focus_topic}." if focus_topic else ""
            prompt = f"""You are a university examiner. Create {num_mcq} MCQ questions. Difficulty: {mcq_diff}. Type: {mcq_type}. {fi}
{lang_instr()}
RULES: NEVER say "from the document" — write standalone academic questions.
Each question: 4 options A/B/C/D, exactly one correct answer.
Return raw JSON only:
{{"subject":"...","subject_ar":"...","questions":[
{{"num":1,"question":"...?","question_ar":"...?","options":{{"A":"...","B":"...","C":"...","D":"..."}},"options_ar":{{"A":"...","B":"...","C":"...","D":"..."}},"correct":"A","explanation":"Why A is correct.","explanation_ar":"لماذا A صحيح.","difficulty":"{mcq_diff}","topic":"...","topic_ar":"..."}}
]}}
Generate exactly {num_mcq} questions covering all major topics."""
            raw  = call_gemini(api_key, prompt, st.session_state.pdf_b64, tokens=8000)
            prog(se, pe, "✅ Done!", 90)
            data = safe_parse(raw)
            st.session_state.mcq_data = data
            se.empty(); pe.empty()
        except Exception as e:
            se.empty(); pe.empty(); st.error(str(e))

    if st.session_state.mcq_data:
        data = st.session_state.mcq_data
        qs   = data.get("questions", [])
        subj = data.get("subject",""); subj_ar = data.get("subject_ar","")
        st.markdown(f'<div style="margin-bottom:.9rem"><span class="pill">📚 <b>{subj}{" | "+subj_ar if subj_ar and show_ar else ""}</b></span><span class="pill">❓ <b>{len(qs)}</b> questions</span></div>', unsafe_allow_html=True)
        docx = make_mcq_docx(data, show_ans, ar_only)
        st.download_button("⬇️ " + ("تنزيل الأسئلة" if ar_only else "Download MCQ (.docx)"), data=docx, file_name="kaabe_mcq.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
        for q in qs:
            corr = q.get("correct","")
            st.markdown(
                f'<div class="mcq-wrap">'
                f'<div style="display:flex;justify-content:space-between;margin-bottom:.38rem">'
                f'<span style="font-size:.69rem;color:#2563eb;font-weight:700">Q{q.get("num","")} · {q.get("topic","")}'
                f'{" | "+q.get("topic_ar","") if q.get("topic_ar") and show_ar else ""}</span>'
                f'<span style="font-size:.69rem;background:#0c1e33;border:1px solid #1e3a6a;border-radius:20px;padding:1px 7px;color:#7ab3d4">{q.get("difficulty","")}</span>'
                f'</div>'
                f'<div class="mcq-q">{q.get("question","")}</div>'
                f'{"<div class=\"mcq-ar\">"+q.get("question_ar","")+"</div>" if q.get("question_ar") and show_ar else ""}',
                unsafe_allow_html=True)
            opts_ar = q.get("options_ar", {})
            for k, v in q.get("options",{}).items():
                cls = "ok" if show_ans and k == corr else ""
                ar_opt = opts_ar.get(k,"")
                st.markdown(
                    f'<div class="mcq-opt {cls}"><span class="mcq-ltr">{k}</span>'
                    f'{v}{" | "+ar_opt if ar_opt and show_ar else ""}</div>',
                    unsafe_allow_html=True)
            if show_ans and q.get("explanation"):
                exp_ar = q.get("explanation_ar","")
                st.markdown(
                    f'<div class="mcq-exp">💡 {q["explanation"]}'
                    f'{"<br><span style=\"font-family:Amiri,serif;direction:rtl;display:block;text-align:right\">"+exp_ar+"</span>" if exp_ar and show_ar else ""}'
                    f'</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 3 · EXAM PAPER
# ══════════════════════════════════════════════════════════════
elif tool == "exam":
    st.markdown("## 📝 " + ("ورقة الاختبار" if ar_only else "Exam Paper"))
    st.markdown('<div class="card"><div class="ctitle">🏗️ Build Your Exam Structure | بناء هيكل الاختبار</div><div style="font-size:.8rem;color:#5a7099">Add sections and choose the question type for each. AI will generate real questions matching your structure.</div></div>', unsafe_allow_html=True)

    c1,c2,c3,c4 = st.columns(4)
    with c1: exam_dur  = st.selectbox("Duration",    ["30 min","1 hour","1.5 hours","2 hours","2.5 hours","3 hours"])
    with c2: total_mks = st.selectbox("Total marks", ["25","50","60","80","100","120"])
    with c3: exam_diff2= st.selectbox("Difficulty",  ["Easy","Medium","Hard","Mixed"])
    with c4: incl_ms   = st.checkbox("Include mark scheme", value=True)

    st.markdown("### 📋 Sections | الأقسام")
    updated = []
    for i, sec in enumerate(st.session_state.exam_sections):
        with st.expander(f"Section {i+1} — {sec['name']} ({(SECTION_TYPES_AR if ar_only else SECTION_TYPES_EN).get(sec['type'],sec['type']).split('|')[0].strip()})", expanded=(i==0)):
            rc1,rc2,rc3,rc4,rc5 = st.columns([2,2,1,1,1])
            with rc1: sn  = st.text_input("Name (EN)",  sec["name"],               key=f"sn{i}")
            with rc2:
                tkeys  = list(SECTION_TYPES_EN.keys())
                cur    = tkeys.index(sec["type"]) if sec["type"] in tkeys else 0
                styp   = st.selectbox("Type", tkeys, index=cur, format_func=lambda x: (SECTION_TYPES_AR if ar_only else SECTION_TYPES_EN).get(x,x), key=f"st{i}")
            with rc3: smk = st.number_input("Marks",1,200,sec["marks"], key=f"sm{i}")
            with rc4: snq = st.number_input("Qs",  1,30, sec["num_q"],  key=f"sq{i}")
            with rc5:
                st.markdown("<div style='height:26px'></div>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"del{i}"):
                    st.session_state.exam_sections.pop(i); st.rerun()
            sna = st.text_input("Name (AR)", sec.get("name_ar",""), key=f"sna{i}")
            sd  = st.text_input("Instructions (EN)", sec.get("desc",""),    key=f"sd{i}")
            sda = st.text_input("Instructions (AR)", sec.get("desc_ar",""), key=f"sda{i}")
            updated.append({"name":sn,"name_ar":sna,"type":styp,"marks":smk,"num_q":snq,"desc":sd,"desc_ar":sda})
    st.session_state.exam_sections = updated

    ac1, ac2 = st.columns(2)
    with ac1:
        if st.button("➕ " + ("إضافة قسم" if ar_only else "Add Section"), use_container_width=True):
            n2 = len(st.session_state.exam_sections)+1
            st.session_state.exam_sections.append({"name":f"Section {chr(64+n2)}","name_ar":f"القسم {n2}","type":"short_answer","marks":20,"num_q":4,"desc":"","desc_ar":""})
            st.rerun()
    with ac2:
        tot = sum(s["marks"] for s in st.session_state.exam_sections)
        st.markdown(f'<div style="background:#090e1a;border:1px solid #1a2a45;border-radius:8px;padding:.5rem .9rem;font-size:.82rem;color:#8899bb;text-align:center">Configured: <b style="color:#60a5fa">{tot}</b> / <b style="color:#fff">{total_mks}</b> marks</div>', unsafe_allow_html=True)

    st.markdown("---")
    if st.button("📝 " + ("توليد الاختبار" if ar_only else "Generate Exam Paper"), type="primary", use_container_width=True):
        if not check_ready(): st.stop()
        se = st.empty(); pe = st.empty()
        try:
            prog(se, pe, "🧠 AI generating exam…", 20)
            ci = f"Course: {course_name}." if course_name else ""
            li = f"Lecturer: {lecturer}." if lecturer else ""
            fi = f"Focus on: {focus_topic}." if focus_topic else ""
            sec_spec = "\n".join(f'  "{s["name"]}": {s["num_q"]} {(SECTION_TYPES_AR if ar_only else SECTION_TYPES_EN).get(s["type"],"").split("|")[0].strip()} questions, {s["marks"]} marks. Instructions: {s["desc"] or "Standard."}' for s in st.session_state.exam_sections)
            total_qs = sum(s["num_q"] for s in st.session_state.exam_sections)

            def stmpl(sec):
                t=sec["type"]; mk=sec["marks"]; n=sec["num_q"]; qmk=max(1,round(mk/n))
                base=f'"name":"{sec["name"]}","name_ar":"{sec["name_ar"]}","type":"{t}","description":"{(SECTION_TYPES_AR if ar_only else SECTION_TYPES_EN).get(t,"").split("|")[0].strip()}","description_ar":"{(SECTION_TYPES_AR if ar_only else SECTION_TYPES_EN).get(t,"").split("|")[-1].strip()}","marks":{mk},"instructions":"{sec["desc"] or "Answer all questions."}","instructions_ar":"{sec["desc_ar"] or "أجب على جميع الأسئلة."}"'
                if t=="mcq":
                    return '{'+base+',"questions":[{"num":"1","text":"Question?","text_ar":"السؤال؟","marks":'+str(qmk)+',"answer_lines":1,"options":{"A":"option","B":"option","C":"option","D":"option"},"correct":"A","model_answer":"A — explanation","model_answer_ar":"أ — الشرح"}]}'
                elif t=="true_false":
                    return '{'+base+',"questions":[{"num":"1","text":"Statement.","text_ar":"العبارة.","marks":'+str(qmk)+',"answer_lines":1,"model_answer":"True/False — reason","model_answer_ar":"صح/خطأ — السبب"}]}'
                elif t=="fill_blank":
                    return '{'+base+',"questions":[{"num":"1","text":"Sentence with _______ blank.","text_ar":"جملة مع فراغ.","marks":'+str(qmk)+',"answer_lines":1,"model_answer":"correct word","model_answer_ar":"الكلمة الصحيحة"}]}'
                elif t=="short_answer":
                    return '{'+base+',"questions":[{"num":"1","text":"Short question.","text_ar":"السؤال.","marks":'+str(qmk)+',"answer_lines":4,"model_answer":"Key points for full marks","model_answer_ar":"النقاط الأساسية"}]}'
                elif t=="long_answer":
                    return '{'+base+',"questions":[{"num":"1","text":"Full question.","text_ar":"السؤال الكامل.","marks":'+str(qmk)+',"answer_lines":12,"parts":[{"part":"a","text":"Part a","text_ar":"الجزء أ","marks":'+str(round(qmk/3))+',"answer_lines":4,"model_answer":"model","model_answer_ar":"النموذج"},{"part":"b","text":"Part b","text_ar":"الجزء ب","marks":'+str(round(qmk/3))+',"answer_lines":4,"model_answer":"model","model_answer_ar":"النموذج"},{"part":"c","text":"Part c","text_ar":"الجزء ج","marks":'+str(qmk-2*round(qmk/3))+',"answer_lines":4,"model_answer":"model","model_answer_ar":"النموذج"}],"model_answer":"Full answer","model_answer_ar":"الإجابة الكاملة"}]}'
                elif t=="calculation":
                    return '{'+base+',"questions":[{"num":"1","text":"Numerical problem.","text_ar":"مسألة حسابية.","marks":'+str(qmk)+',"answer_lines":10,"parts":[{"part":"a","text":"Setup","text_ar":"الإعداد","marks":'+str(round(qmk/3))+',"answer_lines":4,"model_answer":"Step 1...","model_answer_ar":"الخطوة 1"},{"part":"b","text":"Solve","text_ar":"الحل","marks":'+str(round(qmk/3))+',"answer_lines":5,"model_answer":"Calculation","model_answer_ar":"الحساب"},{"part":"c","text":"Interpret","text_ar":"التفسير","marks":'+str(qmk-2*round(qmk/3))+',"answer_lines":2,"model_answer":"Interpretation","model_answer_ar":"التفسير"}],"model_answer":"Full solution","model_answer_ar":"الحل الكامل"}]}'
                elif t=="matching":
                    return '{'+base+',"col_a":["Term 1","Term 2","Term 3","Term 4","Term 5"],"col_b":["Definition A","Definition B","Definition C","Definition D","Definition E"],"questions":[{"num":"1","text":"Match Column A with Column B.","text_ar":"طابق العمود أ مع العمود ب.","marks":'+str(mk)+',"answer_lines":2,"model_answer":"1-B 2-D 3-A 4-C 5-E","model_answer_ar":"1-ب 2-د 3-أ 4-ج 5-هـ"}]}'
                else:
                    return '{'+base+',"questions":[{"num":"1","text":"Question.","text_ar":"السؤال.","marks":'+str(qmk)+',"answer_lines":8,"model_answer":"Model answer","model_answer_ar":"نموذج الإجابة"}]}'

            secs_json = "[" + ",\n".join(stmpl(s) for s in st.session_state.exam_sections) + "]"

            prompt = f"""You are a chief university examiner. Read this PDF and create a complete professional exam.
{ci} {li} Duration: {exam_dur}. Total marks: {total_mks}. Difficulty: {exam_diff2}. {fi}
{lang_instr()}

STRUCTURE (follow exactly):
{sec_spec}

CRITICAL RULES:
- NEVER say "from the document", "according to the text", "as stated in", "the document" — write standalone academic questions.
- Write as an examiner testing knowledge directly.
- MCQ: one correct answer, three plausible distractors.
- Matching: provide col_a (5 real terms) and col_b (5 real definitions from the content).
- Calculation: use real numbers from the content.
- Complete model answers with all marking points.

Return raw JSON only, no fences:
{{"title":"Exam Title","title_ar":"عنوان الاختبار","institution":"{institution or 'University'}","institution_ar":"{institution or 'الجامعة'}","course":"{course_name or 'Course'}","course_ar":"{course_name or 'المادة'}","lecturer":"{lecturer or ''}","lecturer_ar":"{lecturer or ''}","duration":"{exam_dur}","total_marks":{total_mks},"date":"________________",
"instructions":[{{"en":"Read all questions carefully.","ar":"اقرأ جميع الأسئلة بعناية."}},{{"en":"Show all working for calculations.","ar":"اعرض جميع خطوات الحل."}},{{"en":"Mobile phones not permitted.","ar":"لا يُسمح بالهواتف المحمولة."}}],
"sections":{secs_json}}}
Replace ALL template questions with REAL questions from the PDF content. Generate exactly {total_qs} questions total."""
            raw  = call_gemini(api_key, prompt, st.session_state.pdf_b64, tokens=5000)
            prog(se, pe, "✅ Parsing exam…", 80)
            data = safe_parse(raw)
            st.session_state.exam_data = data
            se.empty(); pe.empty()
        except Exception as e:
            se.empty(); pe.empty(); st.error(str(e))

    if st.session_state.exam_data:
        data = st.session_state.exam_data
        docx = make_exam_docx(data, incl_ms, ar_only)
        st.success("✅ " + ("ورقة الاختبار جاهزة!" if ar_only else "Exam paper ready!"))
        safe = re.sub(r"[^a-zA-Z0-9_\- ]","",data.get("title","exam"))[:40].strip().replace(" ","_") or "kaabe_exam"
        st.download_button("⬇️ " + ("تنزيل الاختبار" if ar_only else "Download Exam (.docx)"), data=docx, file_name=f"{safe}.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)

        d = data
        st.markdown(f"""
        <div style="background:#090e1a;border:1px solid #1a2a45;border-radius:10px;padding:1.1rem 1.4rem;margin-top:.8rem">
          <div style="text-align:center;border-bottom:1px solid #1a2a45;padding-bottom:.7rem;margin-bottom:.7rem">
            <div style="font-size:.8rem;color:#60a5fa;font-weight:700">{d.get('institution','')}</div>
            <div style="font-size:1.1rem;font-weight:700;color:#fff;margin:.22rem 0">{d.get('title','')}</div>
            <div style="font-family:'Amiri',serif;font-size:1rem;color:#a78bfa;direction:rtl">{d.get('title_ar','')}</div>
            <div style="font-size:.76rem;color:#5a7099;margin-top:.22rem">{d.get('course','')} · {d.get('duration','')} · {d.get('total_marks','')} marks</div>
          </div>
        """, unsafe_allow_html=True)
        st.markdown("**Instructions:**")
        for inst in d.get("instructions",[]):
            en = inst.get("en","") if isinstance(inst,dict) else str(inst)
            ar = inst.get("ar","") if isinstance(inst,dict) else ""
            st.markdown(f"- {en}" + (f"  |  *{ar}*" if ar and show_ar else ""))
        st.markdown("---")
        for sec in d.get("sections",[]):
            stype = sec.get("type","short_answer")
            name_ar_s = f" | {sec.get('name_ar','')}" if sec.get("name_ar") and show_ar else ""
            st.markdown(f'<div class="exam-sec">{sec.get("name","")}{name_ar_s}  —  {sec.get("description","")}  ({sec.get("marks","")} marks)</div>', unsafe_allow_html=True)
            if stype=="matching" and sec.get("col_a"):
                mc1, mc2 = st.columns(2)
                with mc1:
                    st.markdown("**Column A**")
                    for i2,v in enumerate(sec["col_a"]): st.markdown(f"{i2+1}. {v}")
                with mc2:
                    st.markdown("**Column B**")
                    for i2,v in enumerate(sec.get("col_b",[])): st.markdown(f"{chr(65+i2)}. {v}")
            for q in sec.get("questions",[]):
                qar = q.get("text_ar","")
                st.markdown(f'<div class="exam-q"><span class="exam-qn">Q{q.get("num","")}.&nbsp;</span>{q.get("text","")}<span class="mktag">[{q.get("marks","")} marks]</span></div>', unsafe_allow_html=True)
                if qar and show_ar:
                    st.markdown(f'<div style="font-family:Amiri,serif;direction:rtl;text-align:right;font-size:.83rem;color:#7ab3d4;padding:.18rem 0">{qar}</div>', unsafe_allow_html=True)
                if stype=="mcq" and q.get("options"):
                    for k,v in q["options"].items():
                        st.markdown(f'<div style="padding:.15rem .5rem .15rem 1.4rem;font-size:.8rem;color:#8899bb">{k}) {v}</div>', unsafe_allow_html=True)
                elif q.get("parts"):
                    for pt in q.get("parts",[]):
                        pt_ar = pt.get("text_ar","")
                        st.markdown(f'<div class="exam-q" style="padding-left:1.6rem"><span class="exam-qn">({pt.get("part","")})&nbsp;</span>{pt.get("text","")}<span class="mktag">[{pt.get("marks","")} mk]</span></div>', unsafe_allow_html=True)
                        if pt_ar and show_ar:
                            st.markdown(f'<div style="font-family:Amiri,serif;direction:rtl;text-align:right;font-size:.8rem;color:#7ab3d4;padding-right:1.4rem">{pt_ar}</div>', unsafe_allow_html=True)
                        for _ in range(min(pt.get("answer_lines",3),4)):
                            st.markdown('<div class="ans-ln"></div>', unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        if incl_ms:
            with st.expander("📋 " + ("مخطط التصحيح" if ar_only else "Mark Scheme")):
                for sec in d.get("sections",[]):
                    st.markdown(f"**{sec.get('name','')}**")
                    for q in sec.get("questions",[]):
                        st.markdown(f"**Q{q.get('num','')}** [{q.get('marks','')} mk]: {q.get('model_answer','')}")
                        if show_ar and q.get("model_answer_ar"): st.markdown(f"*{q['model_answer_ar']}*")
                        if q.get("correct"):
                            c2 = q["correct"]; opts2 = q.get("options",{})
                            st.markdown(f"  ✓ **{c2}**: {opts2.get(c2,'')}")
                        if q.get("parts"):
                            for pt in q["parts"]: st.markdown(f"  **({pt.get('part','')})** {pt.get('model_answer','')}")

# ══════════════════════════════════════════════════════════════
# 4 · SMART SUMMARY
# ══════════════════════════════════════════════════════════════
elif tool == "summary":
    st.markdown("## 📋 " + ("الملخص الذكي" if ar_only else "Smart Summary"))
    c1,c2 = st.columns(2)
    with c1:
        sum_depth = st.selectbox("Depth",["Quick overview","Standard","Detailed"])
        incl_kw   = st.checkbox("Key terms & definitions", value=True)
    with c2:
        incl_fo   = st.checkbox("Extract all formulas",    value=True)
        incl_mm   = st.checkbox("Mind map structure",      value=True)

    if st.button("📋 " + ("توليد الملخص" if ar_only else "Generate Summary"), type="primary", use_container_width=True):
        if not check_ready(): st.stop()
        se = st.empty(); pe = st.empty()
        try:
            prog(se, pe, "🧠 AI summarizing…", 35)
            fi = f"Focus on: {focus_topic}." if focus_topic else ""
            prompt = f"""Create a {sum_depth} academic summary of this PDF. {fi}
{lang_instr()}
NEVER say "from the document" — state all facts directly.
Return raw JSON only:
{{"title":"...","title_ar":"...","subject":"...","subject_ar":"...",
"overview":"2-3 sentence overview","overview_ar":"نظرة عامة",
"sections":[{{"heading":"...","heading_ar":"...","summary":"3-5 sentences","summary_ar":"...","key_points":["point 1","point 2","point 3"],"key_points_ar":["نقطة 1","نقطة 2","نقطة 3"]}}],
"key_terms":[{{"term":"...","term_ar":"...","definition":"...","definition_ar":"...","example":"..."}}],
"formulas":[{{"name":"...","name_ar":"...","formula":"...","meaning":"...","meaning_ar":"...","variables":"..."}}],
"mind_map":{{"center":"main topic","center_ar":"الموضوع","branches":[{{"topic":"branch","topic_ar":"فرع","subtopics":["s1","s2","s3"],"subtopics_ar":["ف1","ف2","ف3"]}}]}},
"key_conclusions":["c1","c2","c3","c4"],"key_conclusions_ar":["ن1","ن2","ن3","ن4"]}}"""
            raw  = call_gemini(api_key, prompt, st.session_state.pdf_b64, tokens=8000)
            prog(se, pe, "✅ Done!", 90)
            data = safe_parse(raw)
            st.session_state.sum_data = data
            se.empty(); pe.empty()
        except Exception as e:
            se.empty(); pe.empty(); st.error(str(e))

    if st.session_state.sum_data:
        data = st.session_state.sum_data
        docx = make_summary_docx(data, ar_only)
        st.download_button("⬇️ " + ("تنزيل الملخص" if ar_only else "Download Summary (.docx)"), data=docx, file_name="kaabe_summary.docx",
                           mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                           use_container_width=True)
        ov_ar = data.get("overview_ar","")
        st.markdown(f'<div class="card"><div class="ctitle">📌 Overview</div><div style="font-size:.88rem;color:#8899bb;line-height:1.8">{data.get("overview","")}</div>{"<div class=\"sum-ar\" style=\"margin-top:.5rem\">"+ov_ar+"</div>" if ov_ar and show_ar else ""}</div>', unsafe_allow_html=True)
        for sec in data.get("sections",[]):
            h_ar = sec.get("heading_ar",""); s_ar = sec.get("summary_ar","")
            st.markdown(f'<div class="sum-s"><div class="sum-h">{sec.get("heading","")}</div>{"<div class=\"sum-ar\">"+h_ar+"</div>" if h_ar and show_ar else ""}<div class="sum-txt">{sec.get("summary","")}</div>{"<div class=\"sum-txt\" style=\"direction:rtl;text-align:right;font-family:Amiri,serif;margin-top:.4rem\">"+s_ar+"</div>" if s_ar and show_ar else ""}</div>', unsafe_allow_html=True)
            kps = sec.get("key_points",[]); kps_ar = (sec.get("key_points_ar") or []) + [""]*20
            for kp, kp_ar in zip(kps, kps_ar):
                st.markdown(f'<div style="padding:.13rem .5rem .13rem .9rem;font-size:.82rem;color:#8ab4d4">▸ {kp}{" | "+kp_ar if kp_ar and show_ar else ""}</div>', unsafe_allow_html=True)
        c1_s, c2_s = st.columns(2)
        with c1_s:
            if incl_kw and data.get("key_terms"):
                st.markdown('<div class="card"><div class="ctitle">📚 Key Terms</div>', unsafe_allow_html=True)
                for kt in data["key_terms"]:
                    t_ar = kt.get("term_ar",""); d_ar = kt.get("definition_ar","")
                    st.markdown(f'<div style="margin-bottom:.65rem"><span class="kterm">{kt.get("term","")}</span>{"<span class=\"kterm\" style=\"font-family:Amiri,serif\">"+t_ar+"</span>" if t_ar and show_ar else ""}<div style="font-size:.78rem;color:#8899bb;margin-top:.22rem">{kt.get("definition","")}</div>{"<div style=\"font-size:.76rem;color:#5a7099;direction:rtl;text-align:right;font-family:Amiri,serif\">"+d_ar+"</div>" if d_ar and show_ar else ""}</div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
        with c2_s:
            if incl_fo and data.get("formulas"):
                st.markdown('<div class="card"><div class="ctitle">📐 Formulas</div>', unsafe_allow_html=True)
                for f in data["formulas"]:
                    n_ar = f.get("name_ar",""); m_ar = f.get("meaning_ar","")
                    st.markdown(f'<div style="margin-bottom:.75rem"><div style="font-family:monospace;font-size:.92rem;color:#60a5fa;background:#090e1a;padding:.32rem .65rem;border-radius:5px;margin-bottom:.22rem">{f.get("formula","")}</div><div style="font-size:.73rem;color:#5a7099">{f.get("name","")}{" | "+n_ar if n_ar and show_ar else ""} — {f.get("meaning","")}{" | "+m_ar if m_ar and show_ar else ""}</div></div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)
        if incl_mm and data.get("mind_map"):
            mm = data["mind_map"]; c_ar = mm.get("center_ar","")
            st.markdown(f'<div class="card"><div class="ctitle">🗺️ Mind Map</div><div style="text-align:center;font-family:Syne,sans-serif;font-size:1.05rem;font-weight:700;color:#2563eb;padding:.6rem;background:#090e1a;border-radius:7px;margin-bottom:.7rem">{mm.get("center","")}{" | "+c_ar if c_ar and show_ar else ""}</div>', unsafe_allow_html=True)
            brc = st.columns(min(len(mm.get("branches",[])),4))
            for i2, br in enumerate(mm.get("branches",[])):
                with brc[i2 % len(brc)]:
                    t_ar = br.get("topic_ar","")
                    st.markdown(f'<div style="background:#090e1a;border:1px solid #1e3060;border-radius:7px;padding:.6rem;margin-bottom:.45rem"><div style="font-weight:600;color:#60a5fa;font-size:.8rem;margin-bottom:.32rem">{br.get("topic","")}{" | "+t_ar if t_ar and show_ar else ""}</div>', unsafe_allow_html=True)
                    subs = br.get("subtopics",[]); subs_ar = (br.get("subtopics_ar") or []) + [""]*20
                    for s2, s2_ar in zip(subs, subs_ar):
                        st.markdown(f'<div style="font-size:.73rem;color:#5a7099;padding:1px 0">· {s2}{" | "+s2_ar if s2_ar and show_ar else ""}</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        st.markdown('<div class="card"><div class="ctitle">✅ Key Conclusions</div>', unsafe_allow_html=True)
        concs = data.get("key_conclusions",[]); concs_ar = (data.get("key_conclusions_ar") or []) + [""]*20
        for c3, c3_ar in zip(concs, concs_ar):
            st.markdown(f'<div style="padding:.2rem 0;font-size:.86rem;color:#8ab4d4">✓ &nbsp;{c3}{" | "+c3_ar if c3_ar and show_ar else ""}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# 5 · FLASHCARDS
# ══════════════════════════════════════════════════════════════
elif tool == "flashcards":
    st.markdown("## 🃏 " + ("البطاقات التعليمية" if ar_only else "Flashcard Study Set"))
    c1,c2 = st.columns(2)
    with c1: fc_type = st.selectbox("Type",  ["Definition","Concept","Formula & Application","Mixed"])
    with c2: fc_lvl  = st.selectbox("Level", ["Introductory","Intermediate","Advanced"])

    if st.button("🃏 " + ("توليد البطاقات" if ar_only else "Generate Flashcards"), type="primary", use_container_width=True):
        if not check_ready(): st.stop()
        se = st.empty(); pe = st.empty()
        try:
            prog(se, pe, "🧠 Creating flashcards…", 40)
            fi = f"Focus on: {focus_topic}." if focus_topic else ""
            prompt = f"""Create {num_fc} study flashcards. Type: {fc_type}. Level: {fc_lvl}. {fi}
{lang_instr()}
NEVER say "from the document" — write standalone academic content.
Return raw JSON only:
{{"subject":"...","subject_ar":"...","cards":[{{"num":1,"front":"Question or term","front_ar":"السؤال أو المصطلح","back":"Complete answer","back_ar":"الإجابة الكاملة","type":"{fc_type}","topic":"topic","topic_ar":"الموضوع","hint":"memory hint","hint_ar":"تلميح"}}]}}
Generate exactly {num_fc} cards covering all important topics."""
            raw  = call_gemini(api_key, prompt, st.session_state.pdf_b64, tokens=3000)
            prog(se, pe, "✅ Done!", 90)
            data = safe_parse(raw)
            st.session_state.fc_data = data
            st.session_state.fc_idx = 0
            st.session_state.fc_show = False
            se.empty(); pe.empty()
        except Exception as e:
            se.empty(); pe.empty(); st.error(str(e))

    if st.session_state.fc_data:
        data   = st.session_state.fc_data
        cards  = data.get("cards",[])
        idx    = st.session_state.fc_idx
        show   = st.session_state.fc_show
        if cards:
            card   = cards[idx % len(cards)]
            pct    = int((idx+1) / len(cards) * 100)
            top_ar = card.get("topic_ar","")
            st.markdown(f'<div class="pb"><div class="pf" style="width:{pct}%"></div></div>', unsafe_allow_html=True)
            st.markdown(f'<div style="text-align:center;font-size:.72rem;color:#5a7099;margin-bottom:.45rem">Card {idx+1}/{len(cards)} · {card.get("topic","")}{" | "+top_ar if top_ar and show_ar else ""} · <span style="color:#2563eb">{card.get("type","")}</span></div>', unsafe_allow_html=True)
            main_text = card.get("back","") if show else card.get("front","")
            main_ar   = card.get("back_ar","") if show else card.get("front_ar","")
            hint = card.get("hint",""); hint_ar = card.get("hint_ar","")
            st.markdown(
                f'<div class="fc">'
                f'<div class="fc-lbl">{"✅ ANSWER | الإجابة" if show else "❓ QUESTION | السؤال"}</div>'
                f'<div class="fc-txt">{main_text}</div>'
                f'{"<div class=\"fc-ar\">"+main_ar+"</div>" if main_ar and show_ar else ""}'
                f'{"<div style=\"font-size:.71rem;color:#5a7099;margin-top:.65rem\">💡 "+hint+(" | "+hint_ar if hint_ar and show_ar else "")+"</div>" if show and hint else ""}'
                f'</div>', unsafe_allow_html=True)
            bc1,bc2,bc3,bc4 = st.columns(4)
            with bc1:
                if st.button("⬅️ Prev", use_container_width=True):
                    st.session_state.fc_idx = max(0,idx-1); st.session_state.fc_show=False; st.rerun()
            with bc2:
                if st.button("🔄 Flip", use_container_width=True, type="primary"):
                    st.session_state.fc_show = not show; st.rerun()
            with bc3:
                if st.button("➡️ Next", use_container_width=True):
                    st.session_state.fc_idx = min(len(cards)-1,idx+1); st.session_state.fc_show=False; st.rerun()
            with bc4:
                if st.button("🔀 Shuffle", use_container_width=True):
                    random.shuffle(cards); st.session_state.fc_data["cards"]=cards
                    st.session_state.fc_idx=0; st.session_state.fc_show=False; st.rerun()
        txt = f"KAABE FLASHCARDS — {data.get('subject','')} | {data.get('subject_ar','')}\n{'='*55}\n\n"
        for c in cards:
            txt += f"Card {c.get('num','')} [{c.get('type','')}] — {c.get('topic','')}\n"
            txt += f"Q: {c.get('front','')}\nQ(AR): {c.get('front_ar','')}\n"
            txt += f"A: {c.get('back','')}\nA(AR): {c.get('back_ar','')}\n"
            if c.get("hint"): txt += f"💡 {c['hint']} | {c.get('hint_ar','')}\n"
            txt += "\n"
        st.download_button("⬇️ " + ("تنزيل البطاقات" if ar_only else "Download Flashcards (.txt)"), data=txt.encode(), file_name="kaabe_flashcards.txt", mime="text/plain")
        with st.expander("📋 " + ("عرض جميع البطاقات" if ar_only else "View All Cards")):
            for c in cards:
                with st.expander(f"Card {c.get('num','')} · {str(c.get('front',''))[:55]}"):
                    st.markdown(f"**Q:** {c.get('front','')}  |  {c.get('front_ar','')}"); st.markdown(f"**A:** {c.get('back','')}  |  {c.get('back_ar','')}")
                    if c.get("hint"): st.markdown(f"💡 *{c['hint']}* | *{c.get('hint_ar','')}*")

# ══════════════════════════════════════════════════════════════
# 6 · AI TUTOR CHAT
# ══════════════════════════════════════════════════════════════
elif tool == "qa":
    st.markdown("## 💬 " + ("المعلم الذكي" if ar_only else "AI Tutor"))
    if not st.session_state.pdf_b64:
        st.warning("📄 Upload a PDF to start chatting.")
    else:
        st.markdown('<div style="font-size:.77rem;color:#5a7099;margin-bottom:.55rem">Quick questions | أسئلة سريعة:</div>', unsafe_allow_html=True)
        qcols = st.columns(4)
        quick_qs = [
            ("Summarise main topics", "لخص الموضوعات"),
            ("List all key formulas",  "اذكر الصيغ"),
            ("All definitions",        "جميع التعريفات"),
            ("Key conclusions",        "الاستنتاجات"),
            ("Create study plan",      "خطة دراسة"),
            ("Explain hardest concept","أصعب مفهوم"),
            ("3 practice questions",   "3 أسئلة تدريبية"),
            ("Explain in Arabic",      "شرح بالعربية"),
        ]
        for i2,(qq_en,qq_ar) in enumerate(quick_qs):
            lbl = qq_ar if ar_only else qq_en
            with qcols[i2%4]:
                if st.button(lbl, key=f"qq{i2}", use_container_width=True):
                    if api_key:
                        q_txt = qq_ar if ar_only else qq_en
                        st.session_state.chat.append({"role":"user","content":q_txt})

        for msg in st.session_state.chat:
            if msg["role"]=="user":
                st.markdown(f'<div class="qmsg quser"><div class="qlbl">{"أنت" if ar_only else "You"}</div>{msg["content"]}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="qmsg qai"><div class="qlbl">🎓 Kaabe AI Tutor</div>{msg["content"]}</div>', unsafe_allow_html=True)

        if st.session_state.chat and st.session_state.chat[-1]["role"]=="user" and api_key:
            with st.spinner("🎓 Thinking…"):
                try:
                    history = "\n".join([f"{'Student' if m['role']=='user' else 'Professor'}: {m['content']}" for m in st.session_state.chat[-6:]])
                    ar_instr = "Answer in Arabic only." if ar_only else ""
                    prompt2 = f"""You are an expert university professor and tutor.
Answer academically but clearly. Include formulas, examples, calculations where relevant.
NEVER say "from the document" — answer as a knowledgeable professor.
{ar_instr}
Conversation:\n{history}
Answer the latest question based on the uploaded content."""
                    response = call_gemini(api_key, prompt2, st.session_state.pdf_b64, tokens=1500, json_mode=False)
                    st.session_state.chat.append({"role":"assistant","content":response})
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))

        c1i, c2i = st.columns([5,1])
        with c1i:
            ph = "اسأل أي سؤال…" if ar_only else "Ask anything about your document…"
            user_input = st.text_input("", placeholder=ph, label_visibility="collapsed", key="ci")
        with c2i:
            send = st.button("📨", use_container_width=True, type="primary")
        if send and user_input.strip():
            if not api_key: st.warning("Enter API key.")
            else: st.session_state.chat.append({"role":"user","content":user_input.strip()}); st.rerun()
        if st.session_state.chat:
            if st.button("🗑️ " + ("مسح المحادثة" if ar_only else "Clear chat")):
                st.session_state.chat = []; st.rerun()
