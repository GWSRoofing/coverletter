import streamlit as st
import anthropic
import json, re, os, shutil, tempfile, zipfile
from pathlib import Path
from lxml import etree
import defusedxml.minidom
from datetime import date

st.set_page_config(page_title="GWS Cover Letter Generator", page_icon="🏠", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Sans+3:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }

.block-container {
    padding-top: 1.5rem !important;
    padding-left: 3rem !important;
    padding-right: 3rem !important;
}

/* Title */
.app-title {
    font-family: 'Source Sans 3', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    color: #1a2744;
    margin: 0;
    padding: 0;
    line-height: 1.2;
}

/* Field labels */
.field-label {
    font-size: 0.7rem;
    font-weight: 700;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #6b7280;
    margin-bottom: 6px;
    margin-top: 0;
    display: block;
}

/* Preview box */
.preview-box {
    background: #f8f9fb;
    border: 1px solid #e2e4e8;
    border-radius: 6px;
    padding: 28px 32px;
    font-size: 0.88rem;
    line-height: 1.8;
}

.preview-empty {
    background: #f8f9fb;
    border: 1px solid #e2e4e8;
    border-radius: 6px;
    display: flex;
    align-items: center;
    justify-content: center;
    color: #9ca3af;
    font-style: italic;
    font-size: 0.85rem;
}

.preview-site { font-weight: 700; text-decoration: underline; }

/* Divider */
.divider {
    border: none;
    border-top: 1px solid #e5e7eb;
    margin: 12px 0 20px 0;
}

/* Buttons — match quote generator exactly */
.stButton > button {
    background: #1a2744 !important;
    color: white !important;
    border: none !important;
    border-radius: 5px !important;
    font-family: 'Source Sans 3', sans-serif !important;
    font-weight: 600 !important;
    font-size: 0.95rem !important;
    padding: 12px 20px !important;
    width: 100% !important;
}

.stButton > button:hover { background: #243456 !important; }

/* Input labels */
.stSelectbox > label,
.stTextArea > label,
.stTextInput > label {
    font-size: 0.7rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: #6b7280 !important;
}

/* Text area and select styling */
div[data-testid="stTextArea"] textarea,
div[data-testid="stTextInput"] input {
    font-family: 'Source Sans 3', sans-serif !important;
    font-size: 0.88rem !important;
    border-color: #e2e4e8 !important;
    border-radius: 5px !important;
    color: #374151 !important;
}

div[data-testid="stSelectbox"] > div > div {
    border-color: #e2e4e8 !important;
    border-radius: 5px !important;
    font-size: 0.88rem !important;
}

/* Remove extra spacing */
div[data-testid="stVerticalBlock"] > div { gap: 0.4rem; }

.section-divider { border: none; border-top: 1px solid #e5e7eb; margin: 14px 0; }
</style>
""", unsafe_allow_html=True)

TEMPLATE_PATH = Path(__file__).parent / "template.docx"
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

ESTIMATORS = {
    'Gary Sparrowhawk': {'initials':'GS', 'email':'gary@gwsroofing.co.uk'},
    'Joe Sparrowhawk':  {'initials':'JS', 'email':'joe@gwsroofing.co.uk'},
    'Gary Dolling':     {'initials':'GD', 'email':'gdolling@gwsroofing.co.uk'},
    'Sam Baldwin':      {'initials':'SB', 'email':'sam@gwsroofing.co.uk'},
}

AI_SYSTEM = """You process dictated cover letter text for GWS Roofing.
Extract and lightly refine fields — fix grammar/punctuation minimally, preserve meaning.
Convert contractions to formal (it's to it is). Ignore filler phrases and self-corrections.
Return ONLY raw JSON, no markdown, no explanation.
Rules: clientName=title case, siteAddress=title case, dear=salutation name only (first name),
scope=sentence case, worksDescription=array split on "new paragraph", guarantee=string or null.
Do NOT extract or include date — it will be set automatically.
Return exactly: {"clientName":"","clientEmail":"","siteAddress":"","dear":"","scope":"","worksDescription":["p1","p2"],"guarantee":null}"""

def _pretty_print_xml(xml_file):
    try:
        content = xml_file.read_text(encoding='utf-8')
        dom = defusedxml.minidom.parseString(content)
        xml_file.write_bytes(dom.toprettyxml(indent='  ', encoding='utf-8'))
    except Exception:
        pass

def _condense_xml(xml_file):
    try:
        with open(xml_file, encoding='utf-8') as f:
            dom = defusedxml.minidom.parse(f)
        for element in dom.getElementsByTagName('*'):
            if element.tagName.endswith(':t'):
                continue
            for child in list(element.childNodes):
                if (child.nodeType == child.TEXT_NODE and child.nodeValue
                        and child.nodeValue.strip() == '') or child.nodeType == child.COMMENT_NODE:
                    element.removeChild(child)
        xml_file.write_bytes(dom.toxml(encoding='UTF-8'))
    except Exception:
        pass

def unpack_docx(docx_path, out_dir):
    with zipfile.ZipFile(docx_path, 'r') as zf:
        zf.extractall(out_dir)
    for xml_file in list(Path(out_dir).rglob('*.xml')) + list(Path(out_dir).rglob('*.rels')):
        _pretty_print_xml(xml_file)

def pack_docx(in_dir, out_path):
    in_path = Path(in_dir)
    for f in in_path.rglob('*.xml'):
        _condense_xml(f)
    for f in in_path.rglob('*.rels'):
        _condense_xml(f)
    all_files = [f for f in in_path.rglob('*') if f.is_file()]
    content_types = [f for f in all_files if f.name == '[Content_Types].xml']
    others = [f for f in all_files if f.name != '[Content_Types].xml']
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for f in content_types + others:
            zf.write(f, f.relative_to(in_path))

def find_para_with(root, placeholder):
    for p in root.iter(f'{{{W}}}p'):
        if placeholder in ''.join((t.text or '') for t in p.iter(f'{{{W}}}t')):
            return p
    return None

def add_spacing(para, after=160):
    pPr = para.find(f'{{{W}}}pPr')
    if pPr is None:
        pPr = etree.Element(f'{{{W}}}pPr')
        para.insert(0, pPr)
    ex = pPr.find(f'{{{W}}}spacing')
    if ex is not None:
        pPr.remove(ex)
    sp = etree.Element(f'{{{W}}}spacing')
    sp.set(f'{{{W}}}after', str(after))
    sp.set(f'{{{W}}}line', '240')
    sp.set(f'{{{W}}}lineRule', 'auto')
    rPr = pPr.find(f'{{{W}}}rPr')
    if rPr is not None:
        pPr.insert(list(pPr).index(rPr), sp)
    else:
        pPr.append(sp)

def make_works_para(template_para, text):
    new_p = etree.fromstring(etree.tostring(template_para))
    for child in list(new_p):
        if child.tag.split('}')[-1] != 'pPr':
            new_p.remove(child)
    add_spacing(new_p, 160)
    run = etree.SubElement(new_p, f'{{{W}}}r')
    rpr_src = template_para.find(f'.//{{{W}}}r/{{{W}}}rPr')
    if rpr_src is not None:
        run.insert(0, etree.fromstring(etree.tostring(rpr_src)))
    t = etree.SubElement(run, f'{{{W}}}t')
    t.text = text
    return new_p

def build_docx(fields, estimator_name):
    est = ESTIMATORS[estimator_name]
    work_dir = tempfile.mkdtemp()
    try:
        unpack_docx(TEMPLATE_PATH, work_dir)
        doc_path = os.path.join(work_dir, 'word', 'document.xml')
        tree = etree.parse(doc_path)
        root = tree.getroot()

        simple = {
            '#Initials/Ali':                est['initials'] + '/Ali',
            '#Date ':                       fields['date'] + ' ',
            '#Client name ':                fields['clientName'] + ' ',
            'Email: #Client email Address': 'Email: ' + fields['clientEmail'],
            'Dear #Dear':                   'Dear ' + fields['dear'],
            '#Estimator name':              estimator_name,
            '#Estimator email':             est['email'],
            '#Scope of works':              fields['scope'],
        }
        for t_elem in root.iter(f'{{{W}}}t'):
            for ph, val in simple.items():
                if t_elem.text and ph in t_elem.text:
                    t_elem.text = t_elem.text.replace(ph, val)
                    if val.endswith(' ') or val.startswith(' '):
                        t_elem.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

        p_site = find_para_with(root, '#Site Address')
        if p_site is not None:
            for t in p_site.iter(f'{{{W}}}t'):
                if '#Site Address' in (t.text or ''):
                    t.text = fields['siteAddress']

        p_works = find_para_with(root, '#Works description')
        if p_works is not None:
            works_paras = [p.strip() for p in fields['worksDescription'] if p.strip()]
            if works_paras:
                for t in p_works.iter(f'{{{W}}}t'):
                    if '#Works description' in (t.text or ''):
                        t.text = works_paras[0]
                add_spacing(p_works, 160)
                parent = p_works.getparent()
                idx = list(parent).index(p_works)
                for i, pt in enumerate(works_paras[1:], 1):
                    parent.insert(idx + i, make_works_para(p_works, pt))

        p_guar = find_para_with(root, '#Guarantee')
        if p_guar is not None:
            guarantee = (fields.get('guarantee') or '').strip()
            if guarantee:
                for t in p_guar.iter(f'{{{W}}}t'):
                    if '#Guarantee' in (t.text or ''):
                        t.text = guarantee
            else:
                p_guar.getparent().remove(p_guar)

        tree.write(doc_path, xml_declaration=True, encoding='UTF-8', standalone=True)
        out_path = os.path.join(work_dir, 'output.docx')
        pack_docx(work_dir, out_path)
        with open(out_path, 'rb') as f:
            return f.read()
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)

def process_with_ai(text, api_key):
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model='claude-sonnet-4-5', max_tokens=1000,
        system=AI_SYSTEM,
        messages=[{'role':'user', 'content': f'Dictated text:\n{text}'}]
    )
    raw = response.content[0].text.replace('```json','').replace('```','').strip()
    return json.loads(raw)

def render_preview(f, est):
    wps = f.get('worksDescription', [])
    if isinstance(wps, str):
        wps = [p.strip() for p in wps.split('\n') if p.strip()]
    works_html = ''.join(f'<p style="margin-bottom:10px">{p}</p>' for p in wps)
    guar_html = f'<p>{f["guarantee"]}</p>' if f.get('guarantee') else ''
    st.markdown(f"""<div class="preview-box">
        <p>{est['initials']}/Ali</p>
        <p>{f.get('date','')}</p><br>
        <p>{f.get('clientName','')}</p>
        <p>Email: {f.get('clientEmail','')}</p><br>
        <p>Re: <span class="preview-site">{f.get('siteAddress','')}</span></p><br>
        <p>Dear {f.get('dear','')}</p><br>
        <p><strong>Scope of work:</strong> {f.get('scope','')}</p><br>
        {works_html}{guar_html}<br>
        <p>Regards,</p><br>
        <p>{f.get('estimatorName','')}</p>
        <p><strong>Email:</strong> {est.get('email','')}</p>
    </div>""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for k in ['fields', 'docx_bytes', 'filename']:
    if k not in st.session_state:
        st.session_state[k] = None

api_key = os.environ.get("ANTHROPIC_API_KEY", "")
today = date.today().strftime('%d/%m/%Y')

# ── Header ────────────────────────────────────────────────────────────────────
col_logo, col_title = st.columns([1, 3])
with col_logo:
    st.image("GWS Roofing Logo.jpg", width=130)
with col_title:
    st.markdown("<p class='app-title'>Cover Letter Generator</p>", unsafe_allow_html=True)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ── Main layout ───────────────────────────────────────────────────────────────
left, right = st.columns([1, 1], gap='large')

with left:
    st.markdown("<span class='field-label'>Estimator</span>", unsafe_allow_html=True)
    estimator = st.selectbox("Estimator", list(ESTIMATORS.keys()),
                             label_visibility="collapsed")

    st.markdown("<span class='field-label'>Dictation</span>", unsafe_allow_html=True)

    # Calculate preview height to match right column
    # Base dictation box height
    dict_height = 310

    dictation = st.text_area("Dictation", height=dict_height, label_visibility="collapsed",
        placeholder=(
            "Client name — Full name/s, with title where appropriate (Mr/Mrs/Miss/Ms)\n"
            "Client email — Spell out if unusual\n"
            "Site address — Include postcode\n"
            "Dear — First name only (e.g. \"Dear Daniel\")\n"
            "Scope of works — Headline description of areas covered\n"
            "Works description — Main body of letter. Say \"new paragraph\" to split sections.\n"
            "Guarantee — Only mention if applicable"))

    col_process, col_reset = st.columns(2)
    with col_process:
        process_btn = st.button("Process with AI", key="process")
    with col_reset:
        reset_btn = st.button("Reset", key="reset")

    if reset_btn:
        st.session_state.fields = None
        st.session_state.docx_bytes = None
        st.session_state.filename = None
        st.rerun()

    if process_btn:
        if not dictation.strip():
            st.error("Please enter your dictation.")
        else:
            with st.spinner("Processing your dictation…"):
                try:
                    parsed = process_with_ai(dictation, api_key)
                    parsed['estimatorName'] = estimator
                    parsed['date'] = today
                    st.session_state.fields = parsed
                    st.session_state.docx_bytes = None
                    st.session_state.filename = None
                    st.rerun()
                except Exception as e:
                    st.error(f"AI error: {e}")

    # ── Editable fields ───────────────────────────────────────────────────────
    if st.session_state.fields is not None:
        f = st.session_state.fields
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

        def field_height(text, min_h=68):
            chars = len(text or '')
            rows = max(1, chars // 55 + 1)
            return max(min_h, rows * 28 + 20)

        client_name  = st.text_area("Client name",
            value=f.get('clientName',''), height=field_height(f.get('clientName','')), key="e_cn")
        client_email = st.text_area("Client email",
            value=f.get('clientEmail',''), height=field_height(f.get('clientEmail','')), key="e_ce")
        site_address = st.text_area("Site address",
            value=f.get('siteAddress',''), height=field_height(f.get('siteAddress','')), key="e_sa")
        dear         = st.text_area("Dear",
            value=f.get('dear',''), height=field_height(f.get('dear','')), key="e_dr")
        scope        = st.text_area("Scope of works",
            value=f.get('scope',''), height=field_height(f.get('scope','')), key="e_sc")

        works_paras = f.get('worksDescription', [])
        if isinstance(works_paras, str):
            works_paras = [p.strip() for p in works_paras.split('\n') if p.strip()]

        updated_works = []
        for i, para in enumerate(works_paras):
            updated = st.text_area(
                f"Works description — paragraph {i+1}",
                value=para,
                height=field_height(para),
                key=f"e_wd_{i}")
            updated_works.append(updated)

        guarantee = st.text_area("Guarantee (leave blank if not applicable)",
            value=f.get('guarantee') or '', height=68, key="e_gu")

        if st.button("Update preview", key="update_preview"):
            st.session_state.fields['clientName']       = client_name
            st.session_state.fields['clientEmail']      = client_email
            st.session_state.fields['siteAddress']      = site_address
            st.session_state.fields['dear']             = dear
            st.session_state.fields['scope']            = scope
            st.session_state.fields['worksDescription'] = updated_works
            st.session_state.fields['guarantee']        = guarantee or None
            st.session_state.docx_bytes = None
            st.rerun()

with right:
    st.markdown("<span class='field-label'>Preview</span>", unsafe_allow_html=True)

    if st.session_state.fields is None:
        st.markdown(f"""<div class="preview-empty" style="height:{dict_height + 60}px;">
            Your letter preview will appear here</div>""", unsafe_allow_html=True)
    else:
        f = st.session_state.fields
        est = ESTIMATORS.get(f.get('estimatorName',''), {'initials':'??','email':''})
        render_preview(f, est)

        st.markdown("<br>", unsafe_allow_html=True)

        if st.session_state.docx_bytes is None:
            if st.button("Generate Word Document", key="generate"):
                with st.spinner("Building your Word document…"):
                    try:
                        docx_bytes = build_docx(f, f.get('estimatorName',''))
                        addr = re.sub(r'[^\w\s]','', f.get('siteAddress','document')).strip()
                        fname = f"GWS letter {addr}.docx"
                        st.session_state.docx_bytes = docx_bytes
                        st.session_state.filename = fname
                        st.rerun()
                    except Exception as e:
                        st.error(f"Document error: {e}")
        else:
            st.success("✅ Document ready!")
            st.download_button(
                label=f"⬇ Download {st.session_state.filename}",
                data=st.session_state.docx_bytes,
                file_name=st.session_state.filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )

