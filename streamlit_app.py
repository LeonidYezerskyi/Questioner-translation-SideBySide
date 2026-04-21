import streamlit as st
import xml.etree.ElementTree as ET
from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import json
import os
import tempfile
import re

def _iter_block_items(parent):
    """Walk body in document order (paragraphs and tables interleaved)."""
    if isinstance(parent, DocxDocument):
        body = parent.element.body
    elif isinstance(parent, _Cell):
        body = parent._tc
    else:
        return
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield ("p", Paragraph(child, parent))
        elif isinstance(child, CT_Tbl):
            yield ("t", Table(child, parent))


_RE_DOCX_Q = re.compile(
    r"^([A-Za-z]+\d+[a-zA-Z]*)\s*[:.]\s*(.+)$",
    re.DOTALL,
)
_META_LINE_PREFIXES = (
    "Q Type:",
    "Rotate/Randomize:",
    "Programmer notes:",
    "Тип запитання:",
    "Ротація/Рандомізація:",
    "Примітки програміста:",
)


def _docx_is_meta_line(txt):
    t = txt.strip()
    return any(t.startswith(p) for p in _META_LINE_PREFIXES)


def _docx_looks_like_section(txt):
    t = txt.strip()
    if _docx_is_meta_line(t) or _RE_DOCX_Q.match(t):
        return False
    if len(t) < 4 or len(t) > 140:
        return False
    if re.match(r"^Part\s+[A-Z]\.", t):
        return True
    letters = [c for c in t if c.isalpha()]
    if len(letters) < 4:
        return False
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    return upper_ratio >= 0.72


def _docx_table_to_answers(qid, table):
    answers = []
    for row in table.rows:
        if not row.cells:
            continue
        code = row.cells[0].text.strip()
        label = (
            row.cells[1].text.replace("\n", " ").strip()
            if len(row.cells) > 1
            else ""
        )
        if not code and not label:
            continue
        if re.match(r"^[\d\-]+$", code):
            answers.append({
                "id": f"{qid}_{code}",
                "role": "row",
                "number": code,
                "text": label,
            })
    return answers


def _docx_flush_question_buffer(current_q, buf):
    """Turn queued paragraphs/tables after a question header into stem text + answers."""
    if not current_q or not buf:
        return

    while buf and buf[0][0] == "p" and _docx_is_meta_line(buf[0][1]):
        buf.pop(0)
    while buf and buf[-1][0] == "p" and _docx_is_meta_line(buf[-1][1]):
        buf.pop()
    if not buf:
        return

    has_table = any(k == "t" for k, _ in buf)

    if has_table:
        stem_parts = []
        for k, d in buf:
            if k == "p" and not _docx_is_meta_line(d):
                stem_parts.append(d.strip())
            elif k == "t":
                break
        current_q["text"] = "\n".join(stem_parts).strip()
        for k, d in buf:
            if k == "t":
                current_q["answers"].extend(_docx_table_to_answers(current_q["id"], d))
        return

    split_at = None
    for i, (k, t) in enumerate(buf):
        if k == "p" and _docx_is_meta_line(t):
            split_at = i
            break

    texts = [d.strip() for k, d in buf if k == "p" and not _docx_is_meta_line(d)]

    if split_at is None:
        if len(buf) <= 5:
            current_q["text"] = "\n".join(texts).strip()
            return
        if texts:
            current_q["text"] = texts[0]
            for j, ot in enumerate(texts[1:], 1):
                current_q["answers"].append({
                    "id": f"{current_q['id']}_{j}",
                    "role": "row",
                    "number": str(j),
                    "text": ot,
                })
        return

    pre_texts = [
        d.strip()
        for k, d in buf[:split_at]
        if k == "p" and not _docx_is_meta_line(d)
    ]
    if not pre_texts:
        return
    current_q["text"] = pre_texts[0]
    for j, ot in enumerate(pre_texts[1:], 1):
        current_q["answers"].append({
            "id": f"{current_q['id']}_{j}",
            "role": "row",
            "number": str(j),
            "text": ot,
        })


def parse_docx(file_bytes):
    """
    Parse script-style questionnaires: question lines like ``S1: …`` or ``D4a. …``,
    tables with coded rows, and paragraph option lists (EN/UA use the same question IDs).
    """
    tmp_path = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name

        doc = Document(tmp_path)

        items = []
        sec_idx = 0
        section_id = "SEC_0"
        current_q = None
        buf = []

        for kind, item in _iter_block_items(doc):
            if kind == "p":
                t = item.text.strip()
                if not t:
                    continue

                if _docx_looks_like_section(t):
                    _docx_flush_question_buffer(current_q, buf)
                    buf = []
                    current_q = None
                    sid = f"SEC_{sec_idx}"
                    sec_idx += 1
                    section_id = sid
                    items.append({
                        "id": sid,
                        "type": "section",
                        "sectionId": sid,
                        "text": t,
                        "answers": [],
                    })
                    continue

                m = _RE_DOCX_Q.match(t)
                if m:
                    _docx_flush_question_buffer(current_q, buf)
                    buf = []
                    qid = m.group(1)
                    title = m.group(2).strip()
                    current_q = {
                        "id": qid,
                        "type": "question",
                        "sectionId": section_id,
                        "text": title,
                        "answers": [],
                    }
                    items.append(current_q)
                    continue

                if current_q:
                    buf.append(("p", t))
            else:
                if current_q:
                    buf.append(("t", item))

        _docx_flush_question_buffer(current_q, buf)

        return items
    finally:
        if tmp_path and os.path.isfile(tmp_path):
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

def parse_file(uploaded_file):
    ext = uploaded_file.name.lower().split('.')[-1]

    if ext == "xml":
        return parse_xml(uploaded_file.read())

    if ext == "docx":
        return parse_docx(uploaded_file.read())

    raise Exception("Unsupported file type")
st.set_page_config(page_title="Survey Bilingual Generator", layout="centered")
st.title("📋 Survey Bilingual Document Generator")
st.markdown(
    "Завантажте два файли опитування (**XML** або **DOCX**): перший — основна мова, "
    "другий — переклад. Для Word очікується структура як у скрипті (рядки **S1:**, **D4a.** тощо, "
    "таблиці з кодами у першій колонці або списки варіантів до наступного питання)."
)

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

    for row in table.rows:
        for cell in row.cells:
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

                if len(merged) == 0:
                    st.warning(
                        "Не вдалося розпізнати структуру DOCX/XML. "
                        "Для Word: питання як **S1: …** або **D4a. …**, секції — переважно ВЕЛИКИМИ ЛІТЕРАМИ, "
                        "варіанти — абзаци до «Q Type» / «Rotate» або рядки таблиці «код | текст». "
                        "Спробуйте експорт у XML, якщо макет інший."
                    )

                with st.expander("Preview merged data (first 10 items)"):
                    st.json(merged[:10])

                with st.spinner("Generating DOCX..."):
                    docx_buf = generate_docx(
                        merged,
                        primary_label,
                        secondary_label
                    )

                if docx_buf is not None:
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
