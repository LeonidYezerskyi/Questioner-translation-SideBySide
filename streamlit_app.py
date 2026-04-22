"""Bilingual survey generator: two XML exports (same schema) -> one DOCX.

Run: ``streamlit run survey_xml_bilingual.py``

Expects the same structure as ``parse_xml`` in ``page copy.tsx`` (``<survey>`` / ``<section>`` / questions with ``id``).
"""
import streamlit as st
import xml.etree.ElementTree as ET
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from collections import deque, defaultdict
from itertools import zip_longest
import re

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

def _answer_code_from_primary_row(ans):
    """Numeric code from table column or from id suffix (e.g. S1_37 -> 37)."""
    num = (ans.get('number') or '').strip()
    if num and re.match(r'^[\d\-]+$', num):
        return num
    aid = ans.get('id', '')
    if '_' in aid:
        suf = aid.rsplit('_', 1)[-1]
        if re.match(r'^[\d\-]+$', suf):
            return suf
    return ''

def _leading_numbered_code_in_text(text):
    """UA sometimes keeps script code as ``37. text`` — align to EN row code 37."""
    m = re.match(r'^\s*(\d+)\s*[).:]\s+', text or '')
    if m:
        return m.group(1)
    return None

def _collapse_bracket_option_runs(texts):
    """
    Merge consecutive short ``[…]`` lines only when both are brief (one logical row split
    across two paragraphs). Long bracket lines (e.g. F1 media labels) must stay separate
    or EN/UA rows shift in the bilingual table.
    """
    out = []
    i = 0
    while i < len(texts):
        t = (texts[i] or "").strip()
        if i + 1 < len(texts):
            u = (texts[i + 1] or "").strip()
            if (
                t.startswith("[")
                and u.startswith("[")
                and len(t) < 120
                and len(u) < 120
            ):
                out.append(texts[i].rstrip() + "\n" + texts[i + 1].rstrip())
                i += 2
                continue
        out.append(texts[i])
        i += 1
    return out

def _merge_answer_lists_aligned(primary_answers, secondary_answers):
    """
    Pair EN/UA answer rows for DOCX where ids are ordinal on one side and coded on the other.
    Uses: (1) explicit numbers at start of UA line -> EN code; (2) merge consecutive '[' lines;
    (3) FIFO for remaining rows in document order.
    """
    if not primary_answers:
        return []
    if not secondary_answers:
        return [
            {
                'id': a['id'],
                'role': a['role'],
                'number': a.get('number', ''),
                'primary': a['text'],
                'secondary': '',
            }
            for a in primary_answers
        ]

    sec_texts = [a.get('text', '') for a in secondary_answers]
    keyed = {}
    non_keyed = []
    for t in sec_texts:
        code = _leading_numbered_code_in_text(t)
        if code is not None:
            keyed[code] = t.strip()
        else:
            non_keyed.append(t)
    non_keyed = _collapse_bracket_option_runs(non_keyed)
    plain = deque(non_keyed)
    used_keyed = set()

    rows = []
    for pa in primary_answers:
        code = _answer_code_from_primary_row(pa)
        sec = ''
        if code and code in keyed and code not in used_keyed:
            sec = keyed[code]
            used_keyed.add(code)
        elif plain:
            sec = plain.popleft().strip()
        rows.append({
            'id': pa['id'],
            'role': pa.get('role', 'row'),
            'number': pa.get('number', ''),
            'primary': pa['text'],
            'secondary': sec,
        })
    return rows

def _answers_merge_by_ids_ok(primary_answers, secondary_answers):
    """True when ids line up (same count + almost every id has non-empty secondary text)."""
    if not primary_answers:
        return True
    if len(primary_answers) != len(secondary_answers):
        return False
    sm = {a['id']: a.get('text', '') for a in secondary_answers}
    hits = sum(1 for a in primary_answers if sm.get(a['id'], '').strip())
    return hits / len(primary_answers) >= 0.92

def _split_secondary_question_body_to_answers(s_item, pa):
    """
    If options were merged into question ``text`` (legacy flush / Word layout),
    move trailing lines into answer rows aligned with ``pa`` by count/order.
    """
    if not s_item or not pa:
        return s_item, []
    sa = list(s_item.get('answers') or [])
    if len(sa) >= len(pa):
        return s_item, sa
    if sa:
        return s_item, sa
    body = (s_item.get('text') or '').strip()
    if not body:
        return s_item, sa
    lines = [ln.strip() for ln in body.splitlines() if ln.strip()]
    if len(lines) < 2 or len(lines) - 1 < len(pa):
        return s_item, sa
    head, tail = lines[0], lines[1:]
    if len(tail) < len(pa):
        return s_item, sa
    if not head.rstrip().endswith(("?", ".", "!", "…", ":", "：")):
        return s_item, sa
    fixed = dict(s_item)
    fixed['text'] = head
    new_sa = []
    for i, a in enumerate(pa):
        new_sa.append({
            'id': a['id'],
            'role': a.get('role', 'row'),
            'number': a.get('number', ''),
            'text': tail[i] if i < len(tail) else '',
        })
    return fixed, new_sa


def merge(primary_items, secondary_items):
    """Merge primary and secondary by ID (FIFO when the same id appears more than once)."""
    secondary_queues = defaultdict(deque)
    for item in secondary_items:
        secondary_queues[item['id']].append(item)

    merged = []
    for p_item in primary_items:
        qid = p_item['id']
        qsec = secondary_queues.get(qid)
        s_item = qsec.popleft() if qsec else None

        pa = p_item.get('answers', [])
        sa = list(s_item.get('answers', [])) if s_item else []

        if s_item is not None:
            s_item, sa = _split_secondary_question_body_to_answers(s_item, pa)

        secondary_answer_map = {a['id']: (a.get('text') or '') for a in sa}

        if pa and sa and not _answers_merge_by_ids_ok(pa, sa):
            merged_ans = _merge_answer_lists_aligned(pa, sa)
        else:
            merged_ans = [
                {
                    'id': a['id'],
                    'role': a['role'],
                    'number': a.get('number', ''),
                    'primary': a['text'],
                    'secondary': secondary_answer_map.get(a['id'], ''),
                }
                for a in pa
            ]

        merged.append({
            'id': p_item['id'],
            'type': p_item['type'],
            'sectionId': p_item['sectionId'],
            'primary': p_item['text'],
            'secondary': s_item['text'] if s_item else '',
            'answers': merged_ans,
            'notes_primary': list(p_item.get('notes') or []),
            'notes_secondary': list((s_item or {}).get('notes') or []),
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


def _format_question_heading(qid, body_text):
    """Same style as reference: ``S4. Question text`` in both language columns."""
    qid = (qid or "").strip()
    body = (body_text or "").strip()
    if not qid:
        return body
    if not body:
        return f"{qid}."
    return f"{qid}. {body}"


def _set_cell_multiline_text(cell, text, font_pt=9):
    """
    Put ``text`` into ``cell``, splitting on ``\\n`` into separate paragraphs so Word
    shows stacked lines inside the cell (single-run ``cell.text`` does not reliably).
    """
    raw = (text or "").replace("\r\n", "\n").replace("\r", "\n")
    lines = [ln.strip() for ln in raw.split("\n")]
    lines = [ln for ln in lines if ln] or ([""] if not raw.strip() else [])
    cell.text = lines[0]
    for ln in lines[1:]:
        cell.add_paragraph(ln)
    for para in cell.paragraphs:
        for run in para.runs:
            run.bold = False
            run.font.size = Pt(font_pt)


def _set_table_width_percent(table, pct_fiftieths):
    """pct_fiftieths: 5000 = 100% of parent (Word ``pct`` width)."""
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        tbl.insert(0, tbl_pr)
    for child in list(tbl_pr):
        if child.tag == qn("w:tblW"):
            tbl_pr.remove(child)
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:type"), "pct")
    tbl_w.set(qn("w:w"), str(pct_fiftieths))
    tbl_pr.append(tbl_w)


def _add_nested_answer_table(cell, answers, use_primary_text, label_column_width=None):
    """
    Nested 2-column table: index | label.
    Rows follow the primary (English) answer list so left/right stay aligned.
    Table is ~full width of the cell and centered.
    """
    leaders = [a for a in answers if (a.get("primary") or "").strip()]
    if not leaders:
        cell.text = ""
        return
    nt = cell.add_table(rows=len(leaders), cols=2)
    nt.style = "Table Grid"
    nt.autofit = False
    nt.alignment = WD_TABLE_ALIGNMENT.CENTER
    _set_table_width_percent(nt, 4800)

    idx_w = Inches(0.52)
    if label_column_width is not None:
        lbl_w = Inches(max(1.8, label_column_width.inches - idx_w.inches))
    else:
        lbl_w = Inches(3.0)

    for ri, ans in enumerate(leaders):
        num = (ans.get("number") or "").strip() or str(ri + 1)
        if use_primary_text:
            label = (ans.get("primary") or "").strip()
        else:
            label = (ans.get("secondary") or "").strip()
        row = nt.rows[ri]
        row.cells[0].text = num
        for para in row.cells[0].paragraphs:
            for run in para.runs:
                run.bold = False
                run.font.size = Pt(9)
        _set_cell_multiline_text(row.cells[1], label, font_pt=9)
        for para in row.cells[1].paragraphs:
            for run in para.runs:
                run.bold = False
        row.cells[0].width = idx_w
        row.cells[1].width = lbl_w

    for nrow in nt.rows:
        for cell in nrow.cells:
            set_cell_bg(cell, "FFFFFF")


def generate_docx(merged, primary_label, secondary_label):
    """Generate bilingual DOCX from merged data."""
    doc = Document()

    # Page: landscape (wider columns for two languages)
    section = doc.sections[0]
    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)
    section.top_margin = Inches(0.7)
    section.bottom_margin = Inches(0.7)
    new_width, new_height = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = new_width
    section.page_height = new_height

    col_half = Inches(
        (section.page_width.inches - section.left_margin.inches - section.right_margin.inches)
        / 2.0
    )

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
            sec = (item.get('secondary') or '').strip()
            if sec:
                cells[0].text = f"  {item['primary'].upper()}"
                cells[1].text = f"  {sec.upper()}"
                for cell in cells:
                    set_cell_bg(cell, '1ABC9C')
                    for para in cell.paragraphs:
                        for run in para.runs:
                            run.bold = True
                            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                            run.font.size = Pt(10)
            else:
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

        # Questions: heading row, then next row = nested option tables (no spacer row)
        qid = str(item.get("id", ""))
        answers = item.get("answers") or []
        prim = (item.get("primary") or "").strip()
        sec = (item.get("secondary") or "").strip()
        has_answers = any((a.get("primary") or "").strip() for a in answers)
        notes_p = item.get("notes_primary") or []
        notes_s = item.get("notes_secondary") or []
        has_notes = bool(notes_p or notes_s)

        if not prim and not has_answers and not has_notes:
            continue

        qrow = table.add_row()
        qc = qrow.cells
        qc[0].text = _format_question_heading(qid, prim) if prim else f"{qid}."
        qc[1].text = _format_question_heading(qid, sec) if sec else f"{qid}."
        for cell in qc:
            set_cell_bg(cell, "EBF5FB")
            for para in cell.paragraphs:
                for run in para.runs:
                    run.bold = True
                    run.font.size = Pt(10)

        if has_answers:
            opt_row = table.add_row()
            oc0, oc1 = opt_row.cells[0], opt_row.cells[1]
            for c in (oc0, oc1):
                set_cell_bg(c, "FFFFFF")
            _add_nested_answer_table(oc0, answers, True, col_half)
            _add_nested_answer_table(oc1, answers, False, col_half)
            if has_notes:
                _append_question_notes_paired(oc0, oc1, notes_p, notes_s)
        elif has_notes:
            opt_row = table.add_row()
            oc0, oc1 = opt_row.cells[0], opt_row.cells[1]
            for c in (oc0, oc1):
                set_cell_bg(c, "FFFFFF")
            _append_question_notes_paired(oc0, oc1, notes_p, notes_s)

    for row in table.rows:
        for cell in row.cells:
            cell.width = col_half

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


st.set_page_config(page_title="Survey Bilingual (XML)", layout="centered")
st.title("Survey Bilingual — XML → DOCX")
st.markdown(
    "Завантажте **два XML** однакової структури (експорт опитування): перший — основна мова, "
    "другий — переклад. Вихід — двоколонковий DOCX для звірки."
)

c1, c2 = st.columns(2)
with c1:
    primary_label = st.text_input("Підпис лівої колонки", value="English")
    primary_file = st.file_uploader("XML (основна мова)", type=["xml"])
with c2:
    secondary_label = st.text_input("Підпис правої колонки", value="Translation")
    secondary_file = st.file_uploader("XML (переклад)", type=["xml"])

if primary_file and secondary_file:
    if st.button("Generate bilingual DOCX", type="primary"):
        with st.spinner("Parsing XML…"):
            try:
                primary_items = parse_xml(primary_file.read())
                secondary_items = parse_xml(secondary_file.read())
                merged = merge(primary_items, secondary_items)
                st.success(f"Merged {len(merged)} items")
                with st.expander("Preview (first 8)"):
                    st.json(merged[:8])
                with st.spinner("Building DOCX…"):
                    buf = generate_docx(merged, primary_label, secondary_label)
                st.download_button(
                    "Download DOCX",
                    data=buf,
                    file_name="survey_bilingual_from_xml.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(str(e))
                st.exception(e)
else:
    st.info("Завантажте обидва XML.")
