import streamlit as st
import anthropic
import json, re, os, shutil, tempfile, zipfile
from pathlib import Path
from lxml import etree
import defusedxml.minidom

st.set_page_config(page_title="GWS Cover Letter", page_icon="🏠", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Libre+Baskerville:ital,wght@0,400;0,700;1,400&family=Source+Sans+3:wght@300;400;500;600&display=swap');
html, body, [class*="css"] { font-family: 'Source Sans 3', sans-serif; }
.gws-header { background:#1a2744; color:white; padding:18px 28px; border-radius:10px; margin-bottom:24px; }
.gws-logo { background:#c0392b; color:white; width:44px; height:44px; border-radius:8px; display:inline-flex; align-items:center; justify-content:center; font-family:'Libre Baskerville',serif; font-weight:700; font-size:.9rem; margin-right:14px; vertical-align:middle; }
.gws-title { font-family:'Libre Baskerville',serif; font-size:1.3rem; font-weight:700; display:inline; vertical-align:middle; }
.preview-box { background:white; border:1px solid #d5d0c8; border-radius:10px; padding:28px 32px; font-size:.92rem; line-height:1.75; }
.preview-site { font-weight:700; text-decoration:underline; }
.hint-box { background:white; border:1px solid #d5d0c8; border-left:4px solid #1a2744; border-radius:8px; padding:14px 18px; font-size:.82rem; color:#5a5550; line-height:1.7; margin-bottom:16px; }
.stButton>button { background:#1a2744 !important; color:white !important; border:none !important; border-radius:8px !important; font-weight:500 !important; width:100% !important; }
.stButton>button:hover { background:#243456 !important; }
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
Rules: date=DD/MM/YYYY, clientName=title case, siteAddress=title case, dear=salutation name only,
scope=sentence case, worksDescription=array split on "new paragraph", guarantee=string or null.
Return exactly: {"date":"","clientName":"","clientEmail":"","siteAddress":"","dear":"","scope":"","worksDescription":["p1","p2"],"guarantee":null}"""

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
    # [Content_Types].xml must be first entry in the zip for Word compatibility
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

for k in ['fields','confirmed','docx_bytes','filename']:
    if k not in st.session_state:
        st.session_state[k] = None

col_logo, col_title = st.columns([1, 3])
with col_logo:
    st.image("GWS Roofing Logo.jpg", width=180)
with col_title:
    st.markdown("<h2 style='color:#1a2744; font-family:Libre Baskerville,serif; padding-top:18px; margin-left:-40px;'>Cover Letter</h2>", unsafe_allow_html=True)

left, right = st.columns([1,1], gap='large')

with left:
    st.markdown("#### Letter Details")
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
        
    st.divider()
    mode = st.radio("Input method",
        ["🎤 Dictate", "✏️ Fill fields manually"],
        horizontal=True, label_visibility="collapsed")
    st.divider()

    if mode == "🎤 Dictate":

        estimator = st.selectbox("Estimator", list(ESTIMATORS.keys()))
        dictation = st.text_area("Dictation", height=260, label_visibility="collapsed",
            placeholder=(
                "Date 22nd April 2026\nClient name Mr John Smith\n"
                "Client email john@example.com\n"
                "Site address 12 Oak Lane, London W5 3AB\nDear Mr Smith\n"
                "Scope of works Full re-roof in plain concrete tile\n"
                "Works description Strip all existing coverings. "
                "New paragraph Install breathable membrane. "
                "New paragraph Relay in plain tile throughout.\n"
                "Guarantee 10-year workmanship guarantee."))
        if st.button("✨ Process with AI"):
            if not api_key:
                st.error("Please enter your Anthropic API key.")
            elif not dictation.strip():
                st.error("Please enter your dictation.")
            else:
                with st.spinner("AI is processing your dictation…"):
                    try:
                        parsed = process_with_ai(dictation, api_key)
                        parsed['estimatorName'] = estimator
                        st.session_state.fields = parsed
                        st.session_state.confirmed = False
                        st.session_state.docx_bytes = None
                        st.rerun()
                    except Exception as e:
                        st.error(f"AI error: {e}")
    else:
        estimator = st.selectbox("Estimator", list(ESTIMATORS.keys()))
        c1, c2 = st.columns(2)
        with c1:
            date         = st.text_input("Date (DD/MM/YYYY)", placeholder="22/04/2026")
            client_name  = st.text_input("Client name", placeholder="Mr John Smith")
            site_address = st.text_input("Site address", placeholder="12 Oak Lane, London W5 3AB")
            scope        = st.text_input("Scope of works", placeholder="Full re-roof in plain tile")
        with c2:
            client_email = st.text_input("Client email", placeholder="john@example.com")
            dear         = st.text_input("Dear", placeholder="Mr Smith")
            guarantee    = st.text_input("Guarantee (optional)", placeholder="10-year workmanship guarantee")
        works = st.text_area("Works description (blank line between paragraphs)", height=140,
            placeholder="First paragraph.\n\nSecond paragraph.\n\nThird paragraph.")
        if st.button("👁 Preview Letter"):
            st.session_state.fields = {
                'estimatorName': estimator, 'date': date,
                'clientName': client_name, 'clientEmail': client_email,
                'siteAddress': site_address, 'dear': dear, 'scope': scope,
                'worksDescription': [p.strip() for p in works.split('\n\n') if p.strip()],
                'guarantee': guarantee or None,
            }
            st.session_state.confirmed = False
            st.session_state.docx_bytes = None
            st.rerun()

with right:
    st.markdown("#### Preview")
    if st.session_state.fields is None:
        st.markdown("""<div class="preview-box" style="min-height:400px;display:flex;
            align-items:center;justify-content:center;color:#a09890;font-style:italic;">
            Your letter preview will appear here</div>""", unsafe_allow_html=True)
    else:
        f = st.session_state.fields
        est = ESTIMATORS.get(f.get('estimatorName',''), {'initials':'??','email':''})
        wps = f.get('worksDescription', [])
        if isinstance(wps, str):
            wps = [p.strip() for p in wps.split('\n') if p.strip()]
        works_html = ''.join(f'<p style="margin-bottom:10px">{p}</p>' for p in wps)
        guar_html  = f'<p>{f["guarantee"]}</p>' if f.get('guarantee') else ''

        st.markdown(f"""<div class="preview-box">
            <p>{est['initials']}/Ali</p><p>{f.get('date','')}</p><br>
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

        st.divider()

        if not st.session_state.confirmed:
            ca, cb = st.columns(2)
            with ca:
                if st.button("✅ Confirm & Generate Word Doc"):
                    with st.spinner("Building your Word document…"):
                        try:
                            docx_bytes = build_docx(f, f.get('estimatorName',''))
                            addr = re.sub(r'[^\w\s]','', f.get('siteAddress','document')).strip()
                            fname = f"GWS letter {addr}.docx"
                            st.session_state.docx_bytes = docx_bytes
                            st.session_state.filename = fname
                            st.session_state.confirmed = True
                            st.rerun()
                        except Exception as e:
                            st.error(f"Document error: {e}")
            with cb:
                if st.button("✏️ Edit"):
                    st.session_state.fields = None
                    st.rerun()

        if st.session_state.confirmed and st.session_state.docx_bytes:
            st.success("✅ Document ready!")
            st.download_button(
                label=f"⬇ Download {st.session_state.filename}",
                data=st.session_state.docx_bytes,
                file_name=st.session_state.filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            if st.button("↺ New letter"):
                for k in ['fields','confirmed','docx_bytes','filename']:
                    st.session_state[k] = None
                st.rerun()
