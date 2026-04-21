import streamlit as st
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json
from docx import Document
import tempfile
import re

def parse_docx(file_bytes):

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name

    doc = Document(tmp_path)

    items = []
    current_section = "default"
    current_question = None

    # -------- TABLES FIRST --------

    for table in doc.tables:

        for row in table.rows:

            cells = row.cells

            if len(cells) == 0:
                continue

            txt = cells[0].text.strip()

            if not txt:
                continue

            # SECTION
            if txt.isupper() and len(txt) < 120:

                current_section = txt

                items.append({
                    "id": current_section,
                    "type": "section",
                    "sectionId": current_section,
                    "text": txt,
                    "answers": []
                })

                continue


            # QUESTION
            m = re.match(r'^([A-Z]+\d+[\.\w]*):?\s+(.*)', txt)

            if m:
                qid = m.group(1)
                qtext = m.group(2)

                current_question = {
                    "id": qid,
                    "type": "question",
                    "sectionId": current_section,
                    "text": qtext,
                    "answers": []
                }

                items.append(current_question)

                continue


            # ANSWER ROW
            m2 = re.match(r'^(\d+)\s+(.*)', txt)

            if m2 and current_question:

                current_question["answers"].append({
                    "id": f"{current_question['id']}_{m2.group(1)}",
                    "role": "row",
                    "number": m2.group(1),
                    "text": m2.group(2)
                })

                continue


    return items

def parse_file(uploaded_file):
    ext = uploaded_file.name.lower().split('.')[-1]

    if ext == "xml":
        return parse_xml(uploaded_file.read())

    if ext == "docx":
        return parse_docx(uploaded_file.read())

    raise Exception("Unsupported file type")
st.set_page_config(page_title="Survey Bilingual Generator", layout="centered")
st.title("📋 Survey Bilingual Document Generator")
st.markdown("Upload two XML survey files to generate a bilingual side-by-side DOCX.")

# ── HELPERS ──────────────────────────────────────────────────────────────────

NON_QUESTION_KEYS = {
    'id', 'title', 'current_version_id', 'current_version_number',
    'next_version_id', 'num_of_versions', 'plotly_export_format',
    'previous_version_id', 'regenerate_collection', 'section_type', 'methodology'
}

def extract_title(elem):
    """Extract question text from <title> or <textblockcontent>."""
    title = elem.find('title')
    if title is not None:
        p = title.find('paragraph')
        if p is not None:
            return get_full_text(p)
    tbc = elem.find('textblockcontent')
    if tbc is not None:
        p = tbc.find('paragraph')
        if p is not None:
            return get_full_text(p)
    return ''

def get_full_text(elem):
    """Recursively get all text from element."""
    parts = []
    if elem.text:
        parts.append(elem.text.strip())
    for child in elem:
        parts.append(get_full_text(child))
        if child.tail:
            parts.append(child.tail.strip())
    return ' '.join(p for p in parts if p)

def extract_answers(elem):
    """Extract all answer rows from any answers structure."""
    answers = []
    answers_node = elem.find('answers')
    if answers_node is None:
        # textquestion uses <row> directly
        for row in elem.findall('row'):
            answers.append({
                'id': row.get('id', ''),
                'role': 'row',
                'text': get_full_text(row)
            })
        return answers

    # answers can be repeated (grid has multiple answers blocks)
    all_answer_nodes = elem.findall('answers')
    for ans_node in all_answer_nodes:
        child_type = ans_node.get('childrenType', 'row')
        for child_tag in ['row', 'grow', 'gcol']:
            for item in ans_node.findall(child_tag):
                answers.append({
                    'id': item.get('id', ''),
                    'role': child_type,
                    'text': get_full_text(item)
                })
    return answers

def parse_xml(xml_bytes):
    """Parse XML bytes into normalized list of items."""
    root = ET.fromstring(xml_bytes)
    items = []

    for section in root.findall('.//section'):
        section_id = section.get('id', '')
        section_title = section.get('title', section_id)

        items.append({
            'id': section_id,
            'type': 'section',
            'sectionId': section_id,
            'text': section_title,
            'answers': []
        })

        # Dynamically find all question-type children
        for child in section:
            tag = child.tag
            if tag in NON_QUESTION_KEYS:
                continue
            q_id = child.get('id')
            if not q_id:
                continue

            items.append({
                'id': q_id,
                'type': tag,
                'sectionId': section_id,
                'text': extract_title(child),
                'answers': extract_answers(child)
            })

    return items

def merge(primary_items, secondary_items):
    """Merge primary and secondary by ID."""
    secondary_map = {item['id']: item for item in secondary_items}
    secondary_answer_map = {}
    for item in secondary_items:
        for ans in item.get('answers', []):
            secondary_answer_map[ans['id']] = ans['text']

    merged = []
    for p_item in primary_items:
        s_item = secondary_map.get(p_item['id'])
        merged.append({
            'id': p_item['id'],
            'type': p_item['type'],
            'sectionId': p_item['sectionId'],
            'primary': p_item['text'],
            'secondary': s_item['text'] if s_item else '',
            'answers': [
                {
                    'id': a['id'],
                    'role': a['role'],
                    'number': a.get('number',''),
                    'primary': a['text'],
                    'secondary': secondary_answer_map.get(a['id'], '')
                }
                for a in p_item.get('answers', [])
            ]
        })
    return merged

def set_cell_bg(cell, hex_color):
    """Set cell background color."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_col_width(table, col_index, width_inches):
    """Set column width."""
    for row in table.rows:
        row.cells[col_index].width = Inches(width_inches)

def generate_docx(merged, primary_label, secondary_label):
    """Generate bilingual DOCX from merged data."""
    doc = Document()

    # Page margins
    section = doc.sections[0]
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)

    # Title
    title = doc.add_heading('Bilingual Survey Document', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph('')

    # Table
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'

    # Header row
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = primary_label
    hdr_cells[1].text = secondary_label

    for cell in hdr_cells:
        set_cell_bg(cell, '2C3E50')
        for para in cell.paragraphs:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in para.runs:
                run.bold = True
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.size = Pt(11)

    # Data rows
    for item in merged:
        if item['type'] == 'section':
            row = table.add_row()
            cells = row.cells
            # Merge cells for section header
            cells[0].merge(cells[1])
            cells[0].text = f"  {item['primary'].upper()}"
            set_cell_bg(cells[0], '1ABC9C')
            for para in cells[0].paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    run.font.size = Pt(10)
            continue

        if item['type'] == 'textblock':
            if not item['primary']:
                continue
            row = table.add_row()
            cells = row.cells
            cells[0].text = item['primary']
            cells[1].text = item['secondary']
            set_cell_bg(cells[0], 'ECF0F1')
            set_cell_bg(cells[1], 'ECF0F1')
            for cell in cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.italic = True
                        run.font.size = Pt(9)
            continue

        # Question row
        if item['primary']:
            row = table.add_row()
            cells = row.cells
            q_text_en = f"[{item['id']}]  {item['primary']}"
            q_text_tr = item['secondary'] if item['secondary'] else ''
            cells[0].text = q_text_en
            cells[1].text = q_text_tr
            set_cell_bg(cells[0], 'EBF5FB')
            set_cell_bg(cells[1], 'EBF5FB')
            for cell in cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.bold = True
                        run.font.size = Pt(10)

      # Answer rows
    for ans in item['answers']:
        if not ans['primary']:
            continue
            
        row = table.add_row()
        cells = row.cells
        
        prefix = '    •  ' if ans['role'] == 'row' else '    ○  '
        num = ans.get('number', '')
        
        cells[0].text = f"{prefix}{num} {ans['primary']}".strip()
        cells[1].text = f"{prefix}{num} {ans['secondary']}".strip()
        
        for cell in cells:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    
    # Set column widths
        for row in table.rows:
            for i, cell in enumerate(row.cells):
                cell.width = Inches(3.8)
                
            buf = io.BytesIO()
            doc.save(buf)
            buf.seek(0)
            return buf

# ── UI ─────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    primary_label = st.text_input(
        "Primary language label",
        value="English"
    )

    primary_file = st.file_uploader(
        "Upload Primary file",
        type=['xml', 'docx']
    )

with col2:
    secondary_label = st.text_input(
        "Secondary language label",
        value="Translation"
    )

    secondary_file = st.file_uploader(
        "Upload Secondary file",
        type=['xml', 'docx']
    )

if primary_file and secondary_file:

    if st.button("🚀 Generate Bilingual DOCX", type="primary"):

        with st.spinner("Parsing and merging..."):

            try:
                primary_items = parse_file(primary_file)
                secondary_items = parse_file(secondary_file)

                merged = merge(primary_items, secondary_items)

                st.success(
                    f"✅ Merged {len(merged)} items successfully"
                )

                with st.expander("Preview merged data (first 10 items)"):
                    st.json(merged[:10])

                with st.spinner("Generating DOCX..."):
                    docx_buf = generate_docx(
                        merged,
                        primary_label,
                        secondary_label
                    )

                st.download_button(
                    label="📥 Download DOCX",
                    data=docx_buf,
                    file_name="survey_bilingual.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.exception(e)

else:
    st.info("👆 Upload both files to get started.")
