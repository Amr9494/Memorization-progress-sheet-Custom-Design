from datetime import date, timedelta
import glob
import io
import os
import arabic_reshaper
from bidi.algorithm import get_display
import docx
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm, Pt, RGBColor
import fitz  # PyMuPDF
from PIL import Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.flowables import HRFlowable
import streamlit as st


# ==========================================
# 1. FONT HANDLING & TEXT UTILITIES
# ==========================================
def get_available_fonts() -> dict[str, str]:
    """Scans repository for .ttf and .otf font files."""
    font_files = glob.glob("**/*.ttf", recursive=True) + glob.glob(
        "**/*.otf", recursive=True
    )
    available_fonts = {}
    for path in font_files:
        clean_name = os.path.splitext(os.path.basename(path))[0]
        available_fonts[clean_name] = path

    if not available_fonts:
        available_fonts["Helvetica (Built-in)"] = ""
    return available_fonts


def register_font_safely(font_name: str, font_path: str) -> str:
    """Registers font with ReportLab, returning active identifier."""
    if not font_path or not os.path.exists(font_path):
        return "Helvetica"
    try:
        pdfmetrics.registerFont(TTFont(font_name, font_path))
        return font_name
    except Exception:
        return "Helvetica"


def process_arabic_pdf(text: str) -> str:
    """Reshapes Arabic text and applies bidirectional reordering."""
    if not text:
        return ""
    reshaped = arabic_reshaper.reshape(text)
    return get_display(reshaped)


def get_weekday_dates(
    start_date: date, end_date: date, target_weekday: int
) -> list:
    """Returns all dates matching the target weekday (0=Monday, 6=Sunday)."""
    curr = start_date
    dates = []
    while curr <= end_date:
        if curr.weekday() == target_weekday:
            dates.append(curr)
        curr += timedelta(days=1)
    return dates


# ==========================================
# 2. BUILT-IN REPORTLAB PDF GENERATOR
# ==========================================
def build_single_day_table_pdf(
    entry_date: date, days_per_page: int, font_name: str, scale_factor: float
) -> Table:
    from reportlab.lib.styles import ParagraphStyle

    if days_per_page == 1:
        row_heights = [1.2 * cm, 4.8 * cm, 4.8 * cm, 4.2 * cm, 3.5 * cm]
        base_ar_sz = 14
        base_en_sz = 10
    elif days_per_page == 2:
        row_heights = [0.9 * cm, 2.5 * cm, 2.5 * cm, 1.8 * cm, 1.8 * cm]
        base_ar_sz = 11
        base_en_sz = 8.5
    else:
        row_heights = [0.7 * cm, 1.6 * cm, 1.6 * cm, 1.1 * cm, 1.1 * cm]
        base_ar_sz = 9.5
        base_en_sz = 7.5

    ar_sz = max(6, int(base_ar_sz * scale_factor))
    en_sz = max(5, int(base_en_sz * scale_factor))
    leading_sz = int(ar_sz * 1.3)

    date_header_str = f"{entry_date.strftime('%d %B')}"
    lbl_np = f"{process_arabic_pdf('الحفظ الجديد')}<br/><font size={en_sz}>New Practice</font>"
    lbl_rev = f"{process_arabic_pdf('الماضي - المراجعة')}<br/><font size={en_sz}>Revision</font>"
    lbl_notes = (
        f"{process_arabic_pdf('الملاحظات')}<br/><font size={en_sz}>Notes</font>"
    )
    lbl_sig = f"{process_arabic_pdf('إمضاء ولي الأمر')}<br/><font size={en_sz}>Father Signature</font>"

    table_data = [
        [date_header_str, "", ""],
        [lbl_np, "", "/5"],
        [lbl_rev, "", "/5"],
        [lbl_notes, "", ""],
        [lbl_sig, "", ""],
    ]

    col_widths = [4.8 * cm, 10.7 * cm, 2.5 * cm]

    p_style = ParagraphStyle(
        "TableCell",
        fontName=font_name,
        fontSize=ar_sz,
        leading=leading_sz,
        alignment=1,
    )

    formatted_data = []
    for row in table_data:
        formatted_row = [
            Paragraph(cell, p_style) if cell else "" for cell in row
        ]
        formatted_data.append(formatted_row)

    t = Table(
        formatted_data, colWidths=col_widths, rowHeights=row_heights, hAlign="CENTER"
    )
    t.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 1, colors.HexColor("#2C3E50")),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D8EDF8")),
                ("SPAN", (0, 0), (2, 0)),
                ("SPAN", (1, 3), (2, 3)),
                ("SPAN", (1, 4), (2, 4)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (2, 1), (2, 2), "CENTER"),
            ]
        )
    )
    return t


def build_notes_page_elements_pdf(
    font_name: str, scale_factor: float, page_num: int
) -> list:
    from reportlab.lib.styles import ParagraphStyle

    title_sz = int(14 * scale_factor)
    p_style = ParagraphStyle(
        "NotesTitle",
        fontName=font_name,
        fontSize=title_sz,
        leading=int(title_sz * 1.3),
        alignment=1,
    )

    elements = [
        Paragraph(f"{process_arabic_pdf('ملاحظات')} / Notes", p_style),
        Spacer(1, 0.3 * cm),
        HRFlowable(
            width="100%",
            thickness=0.8,
            color=colors.gray,
            spaceAfter=15,
            dash=[2, 2],
        ),
    ]

    for _ in range(16):
        elements.append(Spacer(1, 1.1 * cm))
        elements.append(
            HRFlowable(
                width="100%",
                thickness=0.4,
                color=colors.HexColor("#B0BEC5"),
                spaceAfter=0,
                dash=[1, 3],
            )
        )

    elements.append(Spacer(1, 0.6 * cm))
    elements.append(
        Paragraph(
            f"<font size=10>—— [ {page_num} ] ——</font>",
            ParagraphStyle("PageNumNotes", alignment=1, fontName=font_name),
        )
    )
    return elements


def generate_content_pdf(
    dates: list,
    days_per_page: int,
    font_name: str,
    scale_factor: float,
    has_front_cover: bool,
    has_back_cover: bool,
) -> io.BytesIO:
    page_chunks = [
        dates[i : i + days_per_page]
        for i in range(0, len(dates), days_per_page)
    ]
    raw_content_pages = len(page_chunks)

    month_break_indices = []
    for idx in range(raw_content_pages - 1):
        month_curr = page_chunks[idx][0].strftime("%m-%Y")
        month_next = page_chunks[idx + 1][0].strftime("%m-%Y")
        if month_curr != month_next:
            month_break_indices.append(idx)

    cover_count = (1 if has_front_cover else 0) + (1 if has_back_cover else 0)
    current_total = raw_content_pages + cover_count
    needed_blanks = (4 - (current_total % 4)) % 4

    blank_insert_after = []
    allocated = 0
    for b_idx in month_break_indices:
        if allocated < needed_blanks:
            blank_insert_after.append(b_idx)
            allocated += 1

    while allocated < needed_blanks:
        blank_insert_after.append(raw_content_pages - 1)
        allocated += 1

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=1.5 * cm,
        leftMargin=1.5 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
    )

    from reportlab.lib.styles import ParagraphStyle

    title_sz = int(16 * scale_factor)
    month_sz = int(12 * scale_factor)

    title_style = ParagraphStyle(
        "BookletTitle",
        fontName=font_name,
        fontSize=title_sz,
        leading=int(title_sz * 1.25),
        alignment=1,
    )
    month_style = ParagraphStyle(
        "MonthHeader",
        fontName=font_name,
        fontSize=month_sz,
        leading=int(month_sz * 1.25),
        alignment=1,
    )

    elements = []
    page_num = 1

    for idx, page_dates in enumerate(page_chunks):
        top_title = process_arabic_pdf(
            "دَفْتَرُ مُتَابَعَةِ حِفْظِ القُرْآنِ الكَرِيم"
        )
        elements.append(Paragraph(top_title, title_style))
        elements.append(Spacer(1, 0.15 * cm))
        elements.append(
            HRFlowable(
                width="100%",
                thickness=0.8,
                color=colors.gray,
                spaceAfter=4,
                dash=[2, 2],
            )
        )

        month_label = page_dates[0].strftime("%B %Y")
        month_tbl = Table(
            [[Paragraph(f"<b>{month_label}</b>", month_style)]],
            colWidths=[18.0 * cm],
            rowHeights=[0.75 * cm],
            hAlign="CENTER",
        )
        month_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#A8E6CF")),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        elements.append(month_tbl)
        elements.append(Spacer(1, 0.4 * cm))

        for d in page_dates:
            elements.append(
                build_single_day_table_pdf(
                    d, days_per_page, font_name, scale_factor
                )
            )
            elements.append(Spacer(1, 0.4 * cm))

        page_num_str = f"—— [ {page_num} ] ——"
        elements.append(Spacer(1, 0.1 * cm))
        elements.append(
            Paragraph(
                f"<font size=11>{page_num_str}</font>",
                ParagraphStyle("PageNum", alignment=1, fontName=font_name),
            )
        )
        page_num += 1

        num_blanks_here = blank_insert_after.count(idx)
        for _ in range(num_blanks_here):
            elements.append(PageBreak())
            elements.extend(
                build_notes_page_elements_pdf(font_name, scale_factor, page_num)
            )
            page_num += 1

        if (
            idx < raw_content_pages - 1
            or (idx == raw_content_pages - 1 and num_blanks_here > 0)
        ) and idx != raw_content_pages - 1:
            elements.append(PageBreak())

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ==========================================
# 3. CUSTOM PDF TEMPLATE OVERLAY PROCESSOR
# ==========================================
def generate_booklet_from_template(
    template_file,
    dates: list,
    days_per_page: int,
    font_path: str,
    font_scale: float,
    has_front_cover: bool,
    has_back_cover: bool,
) -> io.BytesIO:
    """Generates the booklet by overlaying dates directly onto an uploaded 1-page PDF template."""
    raw_template_bytes = template_file.getvalue()
    source_doc = fitz.open(stream=raw_template_bytes, filetype="pdf")

    if len(source_doc) < 1:
        raise ValueError("The uploaded template PDF is empty.")

    page_chunks = [
        dates[i : i + days_per_page]
        for i in range(0, len(dates), days_per_page)
    ]
    raw_content_pages = len(page_chunks)

    # Detect month boundaries
    month_break_indices = []
    for idx in range(raw_content_pages - 1):
        if (
            page_chunks[idx][0].strftime("%m-%Y")
            != page_chunks[idx + 1][0].strftime("%m-%Y")
        ):
            month_break_indices.append(idx)

    # Signature math (modulo 4 booklet layout)
    cover_count = (1 if has_front_cover else 0) + (1 if has_back_cover else 0)
    current_total = raw_content_pages + cover_count
    needed_blanks = (4 - (current_total % 4)) % 4

    blank_insert_after = []
    allocated = 0
    for b_idx in month_break_indices:
        if allocated < needed_blanks:
            blank_insert_after.append(b_idx)
            allocated += 1
    while allocated < needed_blanks:
        blank_insert_after.append(raw_content_pages - 1)
        allocated += 1

    final_doc = fitz.open()
    page_num = 1

    for idx, page_dates in enumerate(page_chunks):
        # Open a fresh instance of the template page
        page_doc = fitz.open(stream=raw_template_bytes, filetype="pdf")
        page = page_doc[0]

        # Register custom font if available
        font_ref = "helv"
        if font_path and os.path.exists(font_path):
            try:
                page.insert_font(fontname="CustomFont", fontfile=font_path)
                font_ref = "CustomFont"
            except Exception:
                font_ref = "helv"

        # Build replacement dictionary
        replacements = {
            "{{PAGE_NUM}}": f"—— [ {page_num} ] ——",
            "{{MONTH}}": process_arabic_pdf(page_dates[0].strftime("%B %Y")),
        }

        for slot_idx, d in enumerate(page_dates):
            tag_name = f"{{{{DATE_{slot_idx + 1}}}}}"
            replacements[tag_name] = d.strftime("%d %B")

        # Clear remaining unused slots on the template
        for unused_slot in range(len(page_dates) + 1, 4):
            replacements[f"{{{{DATE_{unused_slot}}}}}"] = ""

        # Search bounding boxes, redact placeholder text, and insert target text
        for tag, val in replacements.items():
            rects = page.search_for(tag)
            for r in rects:
                page.add_redact_annot(r)
                page.apply_redactions()

                if val:
                    font_sz = int(11 * font_scale)
                    page.insert_textbox(
                        r,
                        val,
                        fontsize=font_sz,
                        fontname=font_ref,
                        align=fitz.TEXT_ALIGN_CENTER,
                    )

        final_doc.insert_pdf(page_doc)
        page_num += 1

        # Add notes/blank pages
        num_blanks = blank_insert_after.count(idx)
        for _ in range(num_blanks):
            blank_page = final_doc.new_page(
                width=page.rect.width, height=page.rect.height
            )
            blank_page.insert_textbox(
                fitz.Rect(50, 40, page.rect.width - 50, 70),
                process_arabic_pdf("ملاحظات / Notes"),
                fontsize=14,
                align=fitz.TEXT_ALIGN_CENTER,
            )
            # Add dashed lined notes background
            y = 100
            while y < page.rect.height - 80:
                p1 = fitz.Point(50, y)
                p2 = fitz.Point(page.rect.width - 50, y)
                blank_page.draw_line(
                    p1, p2, color=(0.7, 0.7, 0.7), dashes="[2 2]"
                )
                y += 30

            blank_page.insert_text(
                fitz.Point(page.rect.width / 2 - 30, page.rect.height - 30),
                f"—— [ {page_num} ] ——",
                fontsize=10,
            )
            page_num += 1

    out_buf = io.BytesIO()
    final_doc.save(out_buf)
    out_buf.seek(0)
    return out_buf


# ==========================================
# 4. COVER MERGER & PREVIEW UTILITIES
# ==========================================
def convert_uploaded_to_fitz_doc(uploaded_file) -> fitz.Document:
    """Converts uploaded PDF or Image safely into an A4 fitz Document."""
    if uploaded_file is None:
        return None

    file_bytes = uploaded_file.getvalue()
    ext = uploaded_file.name.split(".")[-1].lower()

    if ext == "pdf":
        return fitz.open(stream=file_bytes, filetype="pdf")

    try:
        pil_img = Image.open(io.BytesIO(file_bytes))
        if pil_img.mode in ("RGBA", "P", "LA", "CMYK"):
            pil_img = pil_img.convert("RGB")

        img_buffer = io.BytesIO()
        pil_img.save(img_buffer, format="JPEG", quality=92, optimize=True)
        clean_img_bytes = img_buffer.getvalue()

        img_doc = fitz.open(stream=clean_img_bytes, filetype="jpeg")
        pdf_bytes = img_doc.convert_to_pdf()
        raw_page_doc = fitz.open("pdf", pdf_bytes)

        final_cover_doc = fitz.open()
        a4_rect = fitz.paper_rect("a4")
        page = final_cover_doc.new_page(
            width=a4_rect.width, height=a4_rect.height
        )
        page.show_pdf_page(a4_rect, raw_page_doc, 0)
        return final_cover_doc

    except Exception as e:
        st.error(f"Error processing cover '{uploaded_file.name}': {e}")
        return None


def merge_complete_booklet(
    content_pdf_bytes: io.BytesIO, front_cover_file, back_cover_file
) -> io.BytesIO:
    final_doc = fitz.open()

    front_doc = convert_uploaded_to_fitz_doc(front_cover_file)
    if front_doc:
        final_doc.insert_pdf(front_doc, from_page=0, to_page=0)

    content_pdf_bytes.seek(0)
    body_doc = fitz.open(stream=content_pdf_bytes.read(), filetype="pdf")
    final_doc.insert_pdf(body_doc)

    back_doc = convert_uploaded_to_fitz_doc(back_cover_file)
    if back_doc:
        final_doc.insert_pdf(back_doc, from_page=0, to_page=0)

    output_buffer = io.BytesIO()
    final_doc.save(output_buffer)
    output_buffer.seek(0)
    return output_buffer


def render_pdf_page(pdf_bytes: io.BytesIO, page_index: int) -> bytes:
    pdf_bytes.seek(0)
    pdf_doc = fitz.open(stream=pdf_bytes.read(), filetype="pdf")
    target_idx = max(0, min(page_index, len(pdf_doc) - 1))
    page = pdf_doc[target_idx]
    pix = page.get_pixmap(dpi=150)
    return pix.tobytes("png")


# ==========================================
# 5. EDITABLE WORD (.DOCX) BUILDER
# ==========================================
def set_cell_shading(cell, color_hex: str):
    shd = parse_xml(
        f'<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="{color_hex}"/>'
    )
    cell._tc.get_or_add_tcPr().append(shd)


def set_cell_borders(cell):
    tcBorders = parse_xml(r"""
        <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
            <w:top w:val="single" w:sz="6" w:space="0" w:color="2C3E50"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="2C3E50"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="2C3E50"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="2C3E50"/>
        </w:tcBorders>
    """)
    cell._tc.get_or_add_tcPr().append(tcBorders)


def set_cell_margins(cell, top=60, bottom=60, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f"</w:tcMar>"
    )
    tcPr.append(tcMar)


def clean_paragraph_spacing(
    paragraph, space_before=0, space_after=0, line_spacing=1.05
):
    p_format = paragraph.paragraph_format
    p_format.space_before = Pt(space_before)
    p_format.space_after = Pt(space_after)
    p_format.line_spacing = line_spacing


def add_single_day_table_docx(
    doc: docx.Document,
    entry_date: date,
    days_per_page: int,
    font_name: str,
    scale_factor: float,
):
    tbl = doc.add_table(rows=5, cols=3)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    col_widths = [Cm(5.0), Cm(10.5), Cm(2.5)]

    if days_per_page == 1:
        row_heights = [Cm(1.1), Cm(4.9), Cm(4.9), Cm(4.2), Cm(3.4)]
        base_ar_sz = 14
        base_en_sz = 10
    elif days_per_page == 2:
        row_heights = [Cm(0.85), Cm(2.4), Cm(2.4), Cm(1.75), Cm(1.75)]
        base_ar_sz = 11
        base_en_sz = 8.5
    else:
        row_heights = [Cm(0.7), Cm(1.5), Cm(1.5), Cm(1.1), Cm(1.1)]
        base_ar_sz = 9.5
        base_en_sz = 7.5

    ar_sz = max(6, int(base_ar_sz * scale_factor))
    en_sz = max(5, int(base_en_sz * scale_factor))

    # Row 0: Header
    tbl.cell(0, 0).merge(tbl.cell(0, 2))
    set_cell_shading(tbl.cell(0, 0), "D8EDF8")
    p_hdr = tbl.cell(0, 0).paragraphs[0]
    p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_hdr, space_before=1, space_after=1)
    run_hdr = p_hdr.add_run(entry_date.strftime("%d %B"))
    run_hdr.font.name = font_name
    run_hdr.font.size = Pt(ar_sz + 1)
    run_hdr.font.bold = True

    # Row 1: New Practice
    p_ar = tbl.cell(1, 0).paragraphs[0]
    p_ar.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_ar)
    run = p_ar.add_run("الحفظ الجديد\n")
    run.font.name = font_name
    run.font.size = Pt(ar_sz)
    run_en = p_ar.add_run("New Practice")
    run_en.font.name = font_name
    run_en.font.size = Pt(en_sz)

    p_score1 = tbl.cell(1, 2).paragraphs[0]
    p_score1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_score1)
    run_s1 = p_score1.add_run("/5")
    run_s1.font.size = Pt(ar_sz)

    # Row 2: Revision
    p_rev = tbl.cell(2, 0).paragraphs[0]
    p_rev.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_rev)
    run = p_rev.add_run("الماضي - المراجعة\n")
    run.font.name = font_name
    run.font.size = Pt(ar_sz)
    run_en = p_rev.add_run("Revision")
    run_en.font.name = font_name
    run_en.font.size = Pt(en_sz)

    p_score2 = tbl.cell(2, 2).paragraphs[0]
    p_score2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_score2)
    run_s2 = p_score2.add_run("/5")
    run_s2.font.size = Pt(ar_sz)

    # Row 3: Notes
    tbl.cell(3, 1).merge(tbl.cell(3, 2))
    p_notes = tbl.cell(3, 0).paragraphs[0]
    p_notes.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_notes)
    run = p_notes.add_run("الملاحظات\n")
    run.font.name = font_name
    run.font.size = Pt(ar_sz)
    run_en = p_notes.add_run("Notes")
    run_en.font.name = font_name
    run_en.font.size = Pt(en_sz)

    # Row 4: Signature
    tbl.cell(4, 1).merge(tbl.cell(4, 2))
    p_sig = tbl.cell(4, 0).paragraphs[0]
    p_sig.alignment = WD_ALIGN_PARAGRAPH.CENTER
    clean_paragraph_spacing(p_sig)
    run = p_sig.add_run("إمضاء ولي الأمر\n")
    run.font.name = font_name
    run.font.size = Pt(ar_sz)
    run_en = p_sig.add_run("Father Signature")
    run_en.font.name = font_name
    run_en.font.size = Pt(en_sz)

    for r_idx, row in enumerate(tbl.rows):
        row.height = row_heights[r_idx]
        trPr = row._tr.get_or_add_trPr()
        trHeight = parse_xml(
            f'<w:trHeight xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
            f'w:val="{int(row_heights[r_idx].pt * 20)}" w:hRule="atLeast"/>'
        )
        trPr.append(trHeight)

        for c_idx, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.width = col_widths[c_idx]
            set_cell_borders(cell)
            set_cell_margins(cell, top=60, bottom=60, left=100, right=100)


def generate_editable_docx(
    dates: list,
    days_per_page: int,
    font_name: str,
    scale_factor: float,
    front_cover_file,
    back_cover_file,
) -> io.BytesIO:
    doc = docx.Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(0.6)
    section.bottom_margin = Cm(0.6)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    if front_cover_file and front_cover_file.name.split(".")[-1].lower() in [
        "png",
        "jpg",
        "jpeg",
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_paragraph_spacing(p)
        p.add_run().add_picture(
            io.BytesIO(front_cover_file.getvalue()), width=Cm(18.0)
        )
        doc.add_page_break()

    page_chunks = [
        dates[i : i + days_per_page]
        for i in range(0, len(dates), days_per_page)
    ]
    raw_content_pages = len(page_chunks)

    month_break_indices = []
    for idx in range(raw_content_pages - 1):
        if (
            page_chunks[idx][0].strftime("%m-%Y")
            != page_chunks[idx + 1][0].strftime("%m-%Y")
        ):
            month_break_indices.append(idx)

    has_fc = front_cover_file is not None
    has_bc = back_cover_file is not None
    cover_count = (1 if has_fc else 0) + (1 if has_bc else 0)
    needed_blanks = (4 - ((raw_content_pages + cover_count) % 4)) % 4

    blank_insert_after = []
    allocated = 0
    for b_idx in month_break_indices:
        if allocated < needed_blanks:
            blank_insert_after.append(b_idx)
            allocated += 1
    while allocated < needed_blanks:
        blank_insert_after.append(raw_content_pages - 1)
        allocated += 1

    page_num = 1
    for idx, page_dates in enumerate(page_chunks):
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_paragraph_spacing(p_title, space_after=2)
        run_t = p_title.add_run(
            "دَفْتَرُ مُتَابَعَةِ حِفْظِ القُرْآنِ الكَرِيم"
        )
        run_t.font.name = font_name
        run_t.font.size = Pt(int(14 * scale_factor))
        run_t.font.bold = True

        month_tbl = doc.add_table(rows=1, cols=1)
        month_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell_m = month_tbl.cell(0, 0)
        cell_m.width = Cm(18.0)
        set_cell_shading(cell_m, "A8E6CF")
        set_cell_margins(cell_m, top=40, bottom=40)
        p_m = cell_m.paragraphs[0]
        p_m.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_paragraph_spacing(p_m)
        run_m = p_m.add_run(page_dates[0].strftime("%B %Y"))
        run_m.font.name = font_name
        run_m.font.size = Pt(int(11 * scale_factor))
        run_m.font.bold = True

        p_spacer = doc.add_paragraph()
        clean_paragraph_spacing(p_spacer, space_after=3)

        for d_idx, d in enumerate(page_dates):
            add_single_day_table_docx(
                doc, d, days_per_page, font_name, scale_factor
            )
            if d_idx < len(page_dates) - 1:
                p_gap = doc.add_paragraph()
                clean_paragraph_spacing(p_gap, space_after=4)

        p_foot = doc.add_paragraph()
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_paragraph_spacing(p_foot, space_before=3, space_after=0)
        r_f = p_foot.add_run(f"—— [ {page_num} ] ——")
        r_f.font.name = font_name
        r_f.font.size = Pt(10)
        page_num += 1

        num_blanks_here = blank_insert_after.count(idx)
        for _ in range(num_blanks_here):
            doc.add_page_break()
            p_n_title = doc.add_paragraph()
            p_n_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clean_paragraph_spacing(p_n_title, space_after=4)
            r_nt = p_n_title.add_run("ملاحظات / Notes")
            r_nt.font.name = font_name
            r_nt.font.size = Pt(int(13 * scale_factor))
            r_nt.font.bold = True

            for _ in range(16):
                p_line = doc.add_paragraph()
                clean_paragraph_spacing(p_line, space_before=11, space_after=0)
                p_line.add_run(". " * 44).font.color.rgb = RGBColor(
                    176, 190, 197
                )

            p_n_foot = doc.add_paragraph()
            p_n_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            clean_paragraph_spacing(p_n_foot, space_before=6, space_after=0)
            p_n_foot.add_run(f"—— [ {page_num} ] ——").font.size = Pt(10)
            page_num += 1

        if (
            idx < raw_content_pages - 1
            or (idx == raw_content_pages - 1 and num_blanks_here > 0)
        ) and idx != raw_content_pages - 1:
            doc.add_page_break()

    if back_cover_file and back_cover_file.name.split(".")[-1].lower() in [
        "png",
        "jpg",
        "jpeg",
    ]:
        doc.add_page_break()
        p_b = doc.add_paragraph()
        p_b.alignment = WD_ALIGN_PARAGRAPH.CENTER
        clean_paragraph_spacing(p_b)
        p_b.add_run().add_picture(
            io.BytesIO(back_cover_file.getvalue()), width=Cm(18.0)
        )

    out_buf = io.BytesIO()
    doc.save(out_buf)
    out_buf.seek(0)
    return out_buf


# ==========================================
# 6. STREAMLIT USER INTERFACE
# ==========================================
st.set_page_config(
    page_title="Quran Memorization Booklet Generator", layout="wide"
)
st.title("📖 Quran Memorization Tracker & Booklet Generator")

layout_col, preview_col = st.columns([1.1, 0.9], gap="large")

with layout_col:
    st.subheader("📑 Design & Template Selection")
    template_mode = st.radio(
        "Choose Design Mode:",
        ["Built-in Standard Layout", "Upload Custom PDF Template"],
        horizontal=True,
    )

    custom_template = None
    if template_mode == "Upload Custom PDF Template":
        custom_template = st.file_uploader(
            "Upload 1-Page PDF Template",
            type=["pdf"],
            help="Upload an A4 PDF sheet designed in Canva, Word, etc., containing your placeholder text.",
        )
        with st.expander("ℹ️ Supported Placeholders Guide"):
            st.markdown("""
            Add plain text boxes in your PDF design with these exact tags:
            * `{{MONTH}}` — Replaced by the month & year (e.g., **October 2026**).
            * `{{DATE_1}}` — Replaced by the 1st session date (e.g., **04 October**).
            * `{{DATE_2}}` — Replaced by the 2nd session date (if 2 or 3 days/page).
            * `{{DATE_3}}` — Replaced by the 3rd session date (if 3 days/page).
            * `{{PAGE_NUM}}` — Replaced by the booklet page index.
            """)

    st.markdown("---")
    st.subheader("⚙️ Session Schedule")

    c1, c2 = st.columns(2)
    with c1:
        start_dt = st.date_input("Start Date", value=date.today())
    with c2:
        end_dt = st.date_input(
            "End Date", value=date.today() + timedelta(days=90)
        )

    c3, c4 = st.columns(2)
    with c3:
        weekdays = [
            "Monday",
            "Tuesday",
            "Wednesday",
            "Thursday",
            "Friday",
            "Saturday",
            "Sunday",
        ]
        selected_day_name = st.selectbox("Day of the Week", weekdays, index=6)
        selected_weekday = weekdays.index(selected_day_name)
    with c4:
        days_per_page = st.number_input(
            "Days per Page", min_value=1, max_value=3, value=2
        )

    available_fonts = get_available_fonts()
    c5, c6 = st.columns(2)
    with c5:
        selected_font_label = st.selectbox(
            "Select Font (from repo)",
            options=list(available_fonts.keys()),
            index=0,
        )
    with c6:
        font_scale = st.slider(
            "Font Size Scale",
            min_value=1.0,
            max_value=2.0,
            value=1.3,
            step=0.05,
        )

    st.markdown("---")
    st.subheader("🎨 Optional Covers (PDF or Image)")
    col_fc, col_bc = st.columns(2)
    with col_fc:
        front_cover = st.file_uploader(
            "Front Cover (Page 1)",
            type=["pdf", "png", "jpg", "jpeg"],
            key="front_cover",
        )
    with col_bc:
        back_cover = st.file_uploader(
            "Back Cover (Final Page)",
            type=["pdf", "png", "jpg", "jpeg"],
            key="back_cover",
        )

with preview_col:
    st.subheader("📄 Live Booklet Preview & Download")

    if start_dt > end_dt:
        st.error("End date must be after start date.")
    elif template_mode == "Upload Custom PDF Template" and not custom_template:
        st.info("👈 Please upload your 1-page PDF template to generate the booklet.")
    else:
        font_path = available_fonts[selected_font_label]
        active_font = register_font_safely(selected_font_label, font_path)
        matched_dates = get_weekday_dates(start_dt, end_dt, selected_weekday)

        if not matched_dates:
            st.warning("No matching dates found in the chosen date range.")
        else:
            if template_mode == "Upload Custom PDF Template":
                content_pdf_bytes = generate_booklet_from_template(
                    template_file=custom_template,
                    dates=matched_dates,
                    days_per_page=days_per_page,
                    font_path=font_path,
                    font_scale=font_scale,
                    has_front_cover=front_cover is not None,
                    has_back_cover=back_cover is not None,
                )
            else:
                content_pdf_bytes = generate_content_pdf(
                    dates=matched_dates,
                    days_per_page=days_per_page,
                    font_name=active_font,
                    scale_factor=font_scale,
                    has_front_cover=front_cover is not None,
                    has_back_cover=back_cover is not None,
                )

            final_pdf_bytes = merge_complete_booklet(
                content_pdf_bytes, front_cover, back_cover
            )

            temp_doc = fitz.open(
                stream=final_pdf_bytes.getvalue(), filetype="pdf"
            )
            total_pages_count = len(temp_doc)

            st.success(
                f"✅ **Total Pages: {total_pages_count}** (Divisible by 4: {total_pages_count % 4 == 0}) — {len(matched_dates)} total session dates."
            )

            preview_page_num = st.number_input(
                f"Preview Page (1 to {total_pages_count})",
                min_value=1,
                max_value=total_pages_count,
                value=1,
                step=1,
            )

            preview_img = render_pdf_page(
                final_pdf_bytes, preview_page_num - 1
            )
            st.image(
                preview_img,
                caption=f"Preview of Page {preview_page_num} of {total_pages_count}",
                use_container_width=True,
            )

            dcol1, dcol2 = st.columns(2)
            with dcol1:
                st.download_button(
                    label="📥 Download Printable PDF",
                    data=final_pdf_bytes,
                    file_name="quran_memorization_booklet.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                    type="primary",
                )

            with dcol2:
                if template_mode == "Built-in Standard Layout":
                    docx_bytes = generate_editable_docx(
                        dates=matched_dates,
                        days_per_page=days_per_page,
                        font_name=selected_font_label,
                        scale_factor=font_scale,
                        front_cover_file=front_cover,
                        back_cover_file=back_cover,
                    )
                    st.download_button(
                        label="📝 Download Editable Word (.docx)",
                        data=docx_bytes,
                        file_name="quran_memorization_booklet.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                else:
                    st.caption("ℹ️ Word (.docx) export is available for Built-in Layout mode.")