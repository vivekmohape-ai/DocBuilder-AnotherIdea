import streamlit as st
import zipfile
import io
import os
import re
import tempfile
from pathlib import Path
from PIL import Image as PILImage
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DocBuilder — Real Estate",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,600;1,300&family=DM+Sans:wght@300;400;500&display=swap');

:root {
    --cream: #F9F5EF;
    --charcoal: #1C1C1C;
    --gold: #B89355;
    --gold-light: #D4AD6F;
    --warm-grey: #8A8480;
    --border: #E0D9CF;
    --card-bg: #FFFFFF;
    --section-bg: #F4EFE8;
}

html, body, [class*="css"] {
    font-family: 'DM Sans', sans-serif;
    background-color: var(--cream) !important;
    color: var(--charcoal);
}

.main { background-color: var(--cream) !important; }
.block-container { padding: 2rem 3rem 4rem 3rem !important; max-width: 1100px; }

/* Hero header */
.hero-header {
    text-align: center;
    padding: 3rem 0 2rem 0;
    border-bottom: 1px solid var(--border);
    margin-bottom: 2.5rem;
}
.hero-title {
    font-family: 'Cormorant Garamond', serif;
    font-size: 3.2rem;
    font-weight: 300;
    letter-spacing: 0.08em;
    color: var(--charcoal);
    margin: 0;
    line-height: 1.1;
}
.hero-title span {
    color: var(--gold);
    font-style: italic;
}
.hero-sub {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.85rem;
    font-weight: 300;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: var(--warm-grey);
    margin-top: 0.6rem;
}

/* Step labels */
.step-label {
    font-family: 'DM Sans', sans-serif;
    font-size: 0.7rem;
    font-weight: 500;
    letter-spacing: 0.25em;
    text-transform: uppercase;
    color: var(--gold);
    margin-bottom: 0.3rem;
    display: block;
}
.step-heading {
    font-family: 'Cormorant Garamond', serif;
    font-size: 1.5rem;
    font-weight: 400;
    color: var(--charcoal);
    margin-bottom: 0.8rem;
}

/* Cards */
.upload-card {
    background: var(--card-bg);
    border: 1px solid var(--border);
    border-radius: 2px;
    padding: 1.8rem;
    margin-bottom: 1.5rem;
}

/* File uploader override */
.stFileUploader > label { display: none; }
[data-testid="stFileUploader"] {
    border: 1.5px dashed var(--border) !important;
    border-radius: 2px !important;
    background: var(--section-bg) !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--gold) !important;
}

/* Textarea */
.stTextArea textarea {
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.82rem !important;
    border: 1px solid var(--border) !important;
    border-radius: 2px !important;
    background: var(--section-bg) !important;
    color: var(--charcoal) !important;
    line-height: 1.7 !important;
}
.stTextArea textarea:focus {
    border-color: var(--gold) !important;
    box-shadow: 0 0 0 2px rgba(184,147,85,0.15) !important;
}

/* Buttons */
.stButton > button {
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.18em !important;
    text-transform: uppercase !important;
    background: var(--charcoal) !important;
    color: var(--cream) !important;
    border: none !important;
    border-radius: 1px !important;
    padding: 0.75rem 2.5rem !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background: var(--gold) !important;
    color: var(--charcoal) !important;
}
[data-testid="stDownloadButton"] > button {
    background: var(--gold) !important;
    color: var(--charcoal) !important;
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.78rem !important;
    font-weight: 500 !important;
    letter-spacing: 0.18em !important;
    text-transform: uppercase !important;
    border: none !important;
    border-radius: 1px !important;
    padding: 0.75rem 2.5rem !important;
    width: 100% !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: var(--charcoal) !important;
    color: var(--cream) !important;
}

/* Preview tags */
.preview-tag {
    display: inline-block;
    background: var(--section-bg);
    border: 1px solid var(--border);
    border-left: 3px solid var(--gold);
    padding: 0.4rem 0.8rem;
    font-size: 0.75rem;
    color: var(--charcoal);
    margin: 0.2rem 0.2rem 0.2rem 0;
    border-radius: 0;
    font-family: 'DM Sans', sans-serif;
}

/* Info / warning boxes */
.stAlert {
    border-radius: 2px !important;
    border-left-color: var(--gold) !important;
}

/* Divider */
.gold-divider {
    height: 1px;
    background: linear-gradient(90deg, var(--gold), transparent);
    border: none;
    margin: 2rem 0;
}

/* Success box */
.success-box {
    background: #F0EBE1;
    border: 1px solid var(--gold);
    border-left: 4px solid var(--gold);
    padding: 1.2rem 1.5rem;
    border-radius: 2px;
    margin: 1rem 0;
}

/* Format hint */
.format-hint {
    background: var(--section-bg);
    border: 1px solid var(--border);
    border-radius: 2px;
    padding: 1.2rem 1.5rem;
    font-size: 0.78rem;
    color: var(--warm-grey);
    line-height: 1.9;
    font-family: 'DM Sans', sans-serif;
}
.format-hint code {
    background: var(--border);
    padding: 0.1rem 0.35rem;
    border-radius: 2px;
    color: var(--charcoal);
    font-size: 0.75rem;
}

/* Metric cards */
.metric-row {
    display: flex;
    gap: 1rem;
    margin: 1rem 0;
}
.metric-card {
    flex: 1;
    background: var(--card-bg);
    border: 1px solid var(--border);
    padding: 1rem;
    text-align: center;
    border-radius: 2px;
}
.metric-number {
    font-family: 'Cormorant Garamond', serif;
    font-size: 2rem;
    font-weight: 300;
    color: var(--gold);
    line-height: 1;
}
.metric-label {
    font-size: 0.68rem;
    letter-spacing: 0.15em;
    text-transform: uppercase;
    color: var(--warm-grey);
    margin-top: 0.3rem;
}

/* Selectbox */
.stSelectbox > div > div {
    border: 1px solid var(--border) !important;
    border-radius: 2px !important;
    background: var(--section-bg) !important;
}

/* Hide streamlit branding */
#MainMenu, footer, header { visibility: hidden; }

/* Expander */
.streamlit-expanderHeader {
    font-family: 'DM Sans', sans-serif !important;
    font-size: 0.8rem !important;
    color: var(--warm-grey) !important;
    letter-spacing: 0.1em !important;
}
</style>
""", unsafe_allow_html=True)

# ── Helper: parse flow text ───────────────────────────────────────────────────
def parse_flow_text(flow_text: str):
    """
    Parse the pasted flow text into a list of sections.
    Each section: { project, entries: [{filename, caption}] }

    Supported formats:
    [PROJECT NAME]
    filename.jpg | Caption text
    filename.png | Caption text

    or

    PROJECT: Project Name
    filename.jpg - Caption text
    """
    sections = []
    current_project = None
    current_entries = []

    for raw_line in flow_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        # Section header: [PROJECT NAME] or PROJECT: Name or ## Name
        project_match = (
            re.match(r'^\[(.+)\]$', line) or
            re.match(r'^PROJECT\s*:\s*(.+)', line, re.IGNORECASE) or
            re.match(r'^#{1,3}\s+(.+)', line)
        )
        if project_match:
            if current_project and current_entries:
                sections.append({"project": current_project, "entries": current_entries})
            current_project = project_match.group(1).strip()
            current_entries = []
            continue

        # Entry line: filename | caption  or  filename - caption  or  filename, caption
        entry_match = (
            re.match(r'^(.+?)\s*[|\-,]\s*(.+)$', line)
        )
        if entry_match and current_project:
            fname = entry_match.group(1).strip()
            caption = entry_match.group(2).strip()
            current_entries.append({"filename": fname, "caption": caption})
        elif line and current_project:
            # Filename only, no caption
            current_entries.append({"filename": line, "caption": ""})

    if current_project and current_entries:
        sections.append({"project": current_project, "entries": current_entries})

    return sections


# ── Helper: extract images from ZIP ──────────────────────────────────────────
def extract_images_from_zip(zip_bytes: bytes):
    """
    Returns dict: { normalized_filename: PIL.Image, ... }
    Also returns folder_structure for display.
    """
    image_map = {}
    folder_structure = {}
    img_exts = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.tiff'}

    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for name in zf.namelist():
            p = Path(name)
            if p.suffix.lower() in img_exts and not name.startswith('__MACOSX'):
                try:
                    data = zf.read(name)
                    img = PILImage.open(io.BytesIO(data)).convert("RGB")
                    key = p.name.lower()
                    image_map[key] = img

                    # Track folder
                    folder = str(p.parent) if str(p.parent) != '.' else 'root'
                    folder_structure.setdefault(folder, []).append(p.name)
                except Exception:
                    pass

    return image_map, folder_structure


# ── Helper: set cell borders to none ─────────────────────────────────────────
def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        border.set(qn('w:sz'), '0')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), 'auto')
        tblBorders.append(border)
    tblPr.append(tblBorders)


# ── Helper: add image to cell ─────────────────────────────────────────────────
def add_image_to_cell(cell, img_bytes_io, width_inches):
    para = cell.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    run.add_picture(img_bytes_io, width=Inches(width_inches))


# ── Helper: add caption below image ──────────────────────────────────────────
def add_caption_to_cell(cell, caption_text, doc):
    para = cell.add_paragraph(caption_text)
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.runs[0] if para.runs else para.add_run(caption_text)
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(100, 95, 90)
    run.font.italic = True
    para.paragraph_format.space_before = Pt(3)
    para.paragraph_format.space_after = Pt(6)


# ── Core: build Word document ─────────────────────────────────────────────────
def build_word_doc(sections, image_map, doc_title="Real Estate Portfolio", progress_cb=None):
    doc = Document()

    # Page margins
    for section in doc.sections:
        section.top_margin = Cm(1.8)
        section.bottom_margin = Cm(1.8)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

    # Usable width in inches
    usable_width = 6.5  # ~16.5cm

    # Title page
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(doc_title)
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(28, 28, 28)
    title_para.paragraph_format.space_before = Pt(48)
    title_para.paragraph_format.space_after = Pt(6)

    sub_para = doc.add_paragraph()
    sub_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = sub_para.add_run("Marketing Collateral — Image Reference Document")
    sub_run.font.size = Pt(10)
    sub_run.font.color.rgb = RGBColor(138, 132, 128)
    sub_para.paragraph_format.space_after = Pt(48)

    doc.add_page_break()

    total_entries = sum(len(s['entries']) for s in sections)
    processed = 0

    for sec_idx, section_data in enumerate(sections):
        project_name = section_data["project"]
        entries = section_data["entries"]

        # Section heading
        heading_para = doc.add_paragraph()
        heading_run = heading_para.add_run(project_name.upper())
        heading_run.font.size = Pt(13)
        heading_run.font.bold = True
        heading_run.font.color.rgb = RGBColor(28, 28, 28)
        heading_run.font.all_caps = True
        heading_para.paragraph_format.space_before = Pt(16)
        heading_para.paragraph_format.space_after = Pt(4)

        # Gold rule under heading
        rule_para = doc.add_paragraph()
        rule_run = rule_para.add_run("─" * 60)
        rule_run.font.size = Pt(7)
        rule_run.font.color.rgb = RGBColor(184, 147, 85)
        rule_para.paragraph_format.space_after = Pt(10)

        # Group images into rows based on orientation
        i = 0
        while i < len(entries):
            entry = entries[i]
            fname = entry["filename"].lower()
            caption = entry["caption"]

            img = image_map.get(fname)
            if img is None:
                # Try partial match
                for key in image_map:
                    if fname in key or key in fname:
                        img = image_map[key]
                        break

            if img is None:
                # Missing image — show placeholder text
                missing_para = doc.add_paragraph(f"⚠ Image not found: {entry['filename']}")
                missing_para.runs[0].font.color.rgb = RGBColor(180, 80, 80)
                missing_para.runs[0].font.size = Pt(9)
                i += 1
                if progress_cb:
                    processed += 1
                    progress_cb(processed / total_entries)
                continue

            w, h = img.size
            is_landscape = w >= h

            # Decide pairing
            if is_landscape and i + 1 < len(entries):
                next_entry = entries[i + 1]
                next_fname = next_entry["filename"].lower()
                next_img = image_map.get(next_fname)
                if next_img is None:
                    for key in image_map:
                        if next_fname in key or key in next_fname:
                            next_img = image_map[key]
                            break

                nw, nh = (next_img.size if next_img else (1, 1))
                next_is_landscape = nw >= nh if next_img else True

                if next_img and next_is_landscape:
                    # Side-by-side 2-column layout
                    col_w = (usable_width - 0.2) / 2

                    table = doc.add_table(rows=2, cols=2)
                    table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    remove_table_borders(table)

                    # Set column widths
                    for col_idx in range(2):
                        for row_idx in range(2):
                            cell = table.cell(row_idx, col_idx)
                            tc = cell._tc
                            tcPr = tc.get_or_add_tcPr()
                            tcW = OxmlElement('w:tcW')
                            tcW.set(qn('w:w'), str(int(col_w * 1440)))
                            tcW.set(qn('w:type'), 'dxa')
                            tcPr.append(tcW)

                    # Image row
                    for col_idx, (cur_img, cur_entry) in enumerate([(img, entry), (next_img, next_entry)]):
                        cell = table.cell(0, col_idx)
                        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        img_io = io.BytesIO()
                        cur_img.save(img_io, format='JPEG', quality=92)
                        img_io.seek(0)
                        add_image_to_cell(cell, img_io, col_w - 0.1)

                    # Caption row
                    for col_idx, cur_entry in enumerate([entry, next_entry]):
                        if cur_entry["caption"]:
                            cap_cell = table.cell(1, col_idx)
                            add_caption_to_cell(cap_cell, cur_entry["caption"], doc)

                    space_para = doc.add_paragraph()
                    space_para.paragraph_format.space_after = Pt(8)

                    i += 2
                    if progress_cb:
                        processed += 2
                        progress_cb(min(processed / total_entries, 1.0))
                    continue

            # Single image (portrait or unpaired landscape)
            img_io = io.BytesIO()
            img.save(img_io, format='JPEG', quality=92)
            img_io.seek(0)

            if is_landscape:
                img_width = usable_width
            else:
                # Portrait: max 3.8 inches wide, centred
                img_width = min(3.8, usable_width * 0.55)

            table = doc.add_table(rows=2, cols=1)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            remove_table_borders(table)

            # Set full width
            tc = table.cell(0, 0)._tc
            tcPr = tc.get_or_add_tcPr()
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(int(usable_width * 1440)))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)

            img_cell = table.cell(0, 0)
            img_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            add_image_to_cell(img_cell, img_io, img_width)

            if caption:
                cap_cell = table.cell(1, 0)
                add_caption_to_cell(cap_cell, caption, doc)

            space_para = doc.add_paragraph()
            space_para.paragraph_format.space_after = Pt(8)

            i += 1
            if progress_cb:
                processed += 1
                progress_cb(min(processed / total_entries, 1.0))

        # Page break after each project section (except last)
        if sec_idx < len(sections) - 1:
            doc.add_page_break()

    out_io = io.BytesIO()
    doc.save(out_io)
    out_io.seek(0)
    return out_io


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-header">
    <p class="hero-sub">Hiranandani Communities · Marketing</p>
    <h1 class="hero-title">Doc<span>Builder</span></h1>
    <p class="hero-sub" style="margin-top:0.5rem">Image Portfolio Generator · Word Document</p>
</div>
""", unsafe_allow_html=True)

# ── Two-column layout ─────────────────────────────────────────────────────────
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.markdown('<span class="step-label">Step 01</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-heading">Upload Your Image Archive</p>', unsafe_allow_html=True)

    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    zip_file = st.file_uploader(
        "Upload ZIP",
        type=["zip"],
        help="Upload a ZIP file containing your project subfolders with images"
    )

    if zip_file:
        with st.spinner("Reading archive..."):
            zip_bytes = zip_file.read()
            image_map, folder_structure = extract_images_from_zip(zip_bytes)

        total_imgs = len(image_map)
        total_folders = len(folder_structure)

        st.markdown(f"""
        <div class="metric-row">
            <div class="metric-card">
                <div class="metric-number">{total_imgs}</div>
                <div class="metric-label">Images Found</div>
            </div>
            <div class="metric-card">
                <div class="metric-number">{total_folders}</div>
                <div class="metric-label">Folders</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.expander("▸ View detected folders & files"):
            for folder, files in sorted(folder_structure.items()):
                st.markdown(f"**📁 {folder}** ({len(files)} images)")
                for f in files[:8]:
                    st.markdown(f'<span class="preview-tag">{f}</span>', unsafe_allow_html=True)
                if len(files) > 8:
                    st.caption(f"... and {len(files) - 8} more")

    st.markdown('</div>', unsafe_allow_html=True)

    # Doc title
    st.markdown('<span class="step-label" style="margin-top:1rem; display:block">Step 02</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-heading">Document Settings</p>', unsafe_allow_html=True)
    st.markdown('<div class="upload-card">', unsafe_allow_html=True)
    doc_title = st.text_input("Document title", value="Hiranandani Communities — Marketing Collateral", label_visibility="visible")
    st.markdown('</div>', unsafe_allow_html=True)


with col2:
    st.markdown('<span class="step-label">Step 03</span>', unsafe_allow_html=True)
    st.markdown('<p class="step-heading">Define Image Sequence & Captions</p>', unsafe_allow_html=True)

    st.markdown("""
    <div class="format-hint">
    <strong>Format guide</strong><br><br>
    Use <code>[Project Name]</code> as a section header.<br>
    Then list each image as: <code>filename.jpg | Caption text</code><br><br>
    <strong>Example:</strong><br>
    <code>[Empress Hill 4 BHK]</code><br>
    <code>digital_screen_1.jpg | Digital Screen — Bedroom View</code><br>
    <code>digital_screen_2.jpg | Digital Screen — Living Room</code><br>
    <code>ooh_poster.jpg | OOH Billboard</code><br>
    <code>standee.jpg | Standee</code><br><br>
    <code>[Safety Week 2026]</code><br>
    <code>badge.png | Safety Week Badge</code><br>
    <code>standee_safety.jpg | Safety Week Standee</code><br><br>
    <em>Landscape images are auto-paired side-by-side.<br>
    Portrait images are displayed centred individually.</em>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<div style='height:0.8rem'></div>", unsafe_allow_html=True)

    flow_text = st.text_area(
        "Paste your image sequence here",
        height=320,
        placeholder="[Project Name]\nfilename.jpg | Caption\nfilename2.jpg | Caption\n\n[Next Project]\n...",
        label_visibility="collapsed"
    )

    if flow_text.strip():
        parsed = parse_flow_text(flow_text)
        if parsed:
            total_e = sum(len(s['entries']) for s in parsed)
            st.markdown(f"""
            <div class="success-box">
            ✓ Parsed <strong>{len(parsed)} sections</strong> · <strong>{total_e} image entries</strong>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("▸ Preview parsed sections"):
                for s in parsed:
                    st.markdown(f"**{s['project']}** — {len(s['entries'])} images")
                    for e in s['entries'][:4]:
                        st.caption(f"  {e['filename']}  ·  {e['caption'] or '(no caption)'}")
                    if len(s['entries']) > 4:
                        st.caption(f"  ... and {len(s['entries']) - 4} more")
        else:
            st.warning("Could not parse the flow text. Check the format above.")


# ── Divider ───────────────────────────────────────────────────────────────────
st.markdown('<div class="gold-divider"></div>', unsafe_allow_html=True)

# ── Generate button ───────────────────────────────────────────────────────────
gen_col1, gen_col2, gen_col3 = st.columns([1.5, 1, 1.5])

with gen_col2:
    generate = st.button("Generate Document", use_container_width=True)

if generate:
    if not zip_file:
        st.error("Please upload a ZIP file containing your images.")
    elif not flow_text.strip():
        st.error("Please paste your image sequence in Step 03.")
    else:
        parsed_sections = parse_flow_text(flow_text)
        if not parsed_sections:
            st.error("Could not parse the flow text. Please check the format.")
        else:
            st.markdown("<br>", unsafe_allow_html=True)
            status_text = st.empty()
            progress_bar = st.progress(0)

            status_text.markdown("*Building your document...*")

            def update_progress(val):
                progress_bar.progress(val)

            try:
                doc_io = build_word_doc(
                    sections=parsed_sections,
                    image_map=image_map,
                    doc_title=doc_title,
                    progress_cb=update_progress
                )
                progress_bar.progress(1.0)
                status_text.markdown("")

                st.markdown("""
                <div class="success-box" style="text-align:center">
                ✓ &nbsp; Document generated successfully
                </div>
                """, unsafe_allow_html=True)

                dl_col1, dl_col2, dl_col3 = st.columns([1.2, 1, 1.2])
                with dl_col2:
                    safe_title = re.sub(r'[^\w\s-]', '', doc_title)[:40].strip().replace(' ', '_')
                    st.download_button(
                        label="Download .docx",
                        data=doc_io,
                        file_name=f"{safe_title}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"Error generating document: {str(e)}")
                st.exception(e)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div style='text-align:center; padding: 3rem 0 1rem 0; border-top: 1px solid var(--border); margin-top: 3rem'>
    <p style='font-family: DM Sans; font-size: 0.7rem; letter-spacing: 0.15em; text-transform: uppercase; color: var(--warm-grey)'>
        DocBuilder · Hiranandani Communities · Real Estate Marketing
    </p>
</div>
""", unsafe_allow_html=True)
