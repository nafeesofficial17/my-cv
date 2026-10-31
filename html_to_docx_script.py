"""Build a Word version of the resume with the same layout as the WeasyPrint PDF.

The Word export reads docx.html, a copy of weasyprint.html with a light sidebar
(the PDF keeps the dark one). Content edits must go into both files. This script
reads the HTML structure (sidebar + main column, blocks, project cards, stats strip)
and re-creates the layout with Word-native elements:

- the sidebar is a page-anchored shape in the header, so it fills every page;
- the two columns are one borderless table row that can break across pages;
- skill chips are shaded runs, project cards and the stats strip are nested tables.

Colours and sizes below mirror the CSS in docx.html; keep them in sync.
"""

from datetime import datetime
from pathlib import Path
import re

from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Emu, Mm, Pt, RGBColor


INPUT_HTML = Path("docx.html")
OUTPUT_DOCX = Path(f"m-nafees-{datetime.now():%Y%m%d}.docx")

FONT = "Segoe UI"

# Layout (mirrors @page, .sidebar and .main)
PAGE_WIDTH = Mm(210)
PAGE_HEIGHT = Mm(297)
PAGE_MARGIN_Y = Mm(10)
SIDEBAR_WIDTH = Mm(210 * 0.32)
MAIN_WIDTH = PAGE_WIDTH - SIDEBAR_WIDTH
SIDEBAR_PAD = {"top": Mm(2), "left": Mm(8), "bottom": Mm(4), "right": Mm(6)}
MAIN_PAD = {"top": Mm(1), "left": Mm(8), "bottom": Mm(4), "right": Mm(10)}
MAIN_CONTENT_WIDTH = MAIN_WIDTH - MAIN_PAD["left"] - MAIN_PAD["right"]
CARD_PAD = {"top": Mm(1.8), "left": Mm(3), "bottom": Mm(1.8), "right": Mm(2.6)}
STAT_GAP = Mm(1.6)

# Colours (mirrors the CSS)
SIDEBAR_BG = "EAF1F8"
SIDEBAR_NAME = "0F172A"
SIDEBAR_ROLE = "0369A1"
SIDEBAR_HEADING = "075985"
SIDEBAR_SUBHEADING = "9A3412"
SIDEBAR_TEXT = "0F172A"
SIDEBAR_VALUE = "0F172A"
SIDEBAR_MUTED = "475569"
SIDEBAR_EDU = "334155"
SIDEBAR_RULE = "C3D3E4"
CHIP_BORDER = "8FA9C6"
HEADING = "0F172A"
BODY = "334155"
DATE = "64748B"
ACCENT = "0284C7"
ACCENT_DARK = "0369A1"
STACK = "0F766E"
MAIN_RULE = "E2E8F0"
STAT_FILL = "F0F9FF"
STAT_LABEL = "475569"
CARD_FILL = "F8FAFC"
PILL_FILL = "E0F2FE"
FOOTER = "94A3B8"

# Child order Word expects inside property elements; out-of-order XML can make
# Word report the file as damaged.
RPR_ORDER = [
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps", "w:smallCaps", "w:strike",
    "w:dstrike", "w:outline", "w:shadow", "w:emboss", "w:imprint", "w:noProof", "w:snapToGrid",
    "w:vanish", "w:webHidden", "w:color", "w:spacing", "w:w", "w:kern", "w:position", "w:sz",
    "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
    "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath",
]
PPR_ORDER = [
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore", "w:framePr", "w:widowControl",
    "w:numPr", "w:suppressLineNumbers", "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
    "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
    "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc", "w:textDirection",
    "w:textAlignment", "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
    "w:sectPr", "w:pPrChange",
]
TCPR_ORDER = [
    "w:cnfStyle", "w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge", "w:tcBorders", "w:shd",
    "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark",
]
TBLPR_ORDER = [
    "w:tblStyle", "w:tblpPr", "w:tblOverlap", "w:bidiVisual", "w:tblStyleRowBandSize",
    "w:tblStyleColBandSize", "w:tblW", "w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders",
    "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
]
EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")


def html_to_docx(html_file: Path = INPUT_HTML, output_file: Path = OUTPUT_DOCX) -> None:
    soup = BeautifulSoup(html_file.read_text(encoding="utf-8"), "html.parser")
    document = build_document(soup)
    saved_file = save_document(document, output_file)

    print(f"DOCX generated successfully: {saved_file}")


def build_document(soup: BeautifulSoup) -> Document:
    sidebar = soup.select_one("aside.sidebar")
    main = soup.select_one("main.main")
    if sidebar is None or main is None:
        raise ValueError("Expected sidebar and main sections in the HTML file.")

    document = Document()
    for paragraph in document.paragraphs:
        paragraph._element.getparent().remove(paragraph._element)
    configure_document(document, soup)
    add_sidebar_background(document)
    add_footer(document, soup)

    table = document.add_table(rows=1, cols=2)
    prepare_table(table, [SIDEBAR_WIDTH, MAIN_WIDTH])
    sidebar_cell, main_cell = table.rows[0].cells
    set_cell_margins(sidebar_cell, **SIDEBAR_PAD)
    set_cell_margins(main_cell, **MAIN_PAD)

    render_sidebar(sidebar_cell, sidebar)
    render_main(main_cell, main)

    # Word needs a paragraph after a table; keep it tiny so it can't spill onto a new page
    style_spacer(document.add_paragraph())
    return document


# ---------------------------------------------------------------------------
# Document setup
# ---------------------------------------------------------------------------

def configure_document(document: Document, soup: BeautifulSoup) -> None:
    section = document.sections[0]
    section.page_width = PAGE_WIDTH
    section.page_height = PAGE_HEIGHT
    section.top_margin = PAGE_MARGIN_Y
    section.bottom_margin = PAGE_MARGIN_Y
    section.left_margin = Mm(0)
    section.right_margin = Mm(0)
    section.header_distance = Mm(0)
    section.footer_distance = Mm(4)

    normal = document.styles["Normal"]
    normal.font.name = FONT
    fonts = normal.element.rPr.rFonts
    for attribute in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        fonts.set(qn(attribute), FONT)
    normal.font.size = Pt(9)
    normal.font.color.rgb = RGBColor.from_string(BODY)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(0)

    # Word 2013+ table layout, so cell padding sits inside the column widths we set
    for setting in document.settings.element.iter(qn("w:compatSetting")):
        if setting.get(qn("w:name")) == "compatibilityMode":
            setting.set(qn("w:val"), "15")

    title = soup.find("title")
    author = soup.find("meta", attrs={"name": "author"})
    keywords = soup.find("meta", attrs={"name": "keywords"})
    description = soup.find("meta", attrs={"name": "description"})
    properties = document.core_properties
    properties.title = clean_text(title.get_text()) if title else ""
    properties.author = author["content"] if author else ""
    properties.keywords = keywords["content"] if keywords else ""
    properties.subject = description["content"] if description else ""


def add_sidebar_background(document: Document) -> None:
    """Page-anchored rectangles behind the text, repeated on every page via the header:
    the sidebar fill plus a thin divider line on its right edge."""
    paragraph = document.sections[0].header.paragraphs[0]
    style_spacer(paragraph)
    for left, width, fill in ((0, SIDEBAR_WIDTH.pt, SIDEBAR_BG), (SIDEBAR_WIDTH.pt, 0.75, SIDEBAR_RULE)):
        shape = parse_xml(
            '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            ' xmlns:v="urn:schemas-microsoft-com:vml"><w:pict>'
            f'<v:rect style="position:absolute;margin-left:{left:.1f}pt;margin-top:0;width:{width:.2f}pt;'
            f'height:{PAGE_HEIGHT.pt:.1f}pt;z-index:-251658240;mso-position-horizontal-relative:page;'
            f'mso-position-vertical-relative:page" fillcolor="#{fill}" stroked="f"/>'
            "</w:pict></w:r>"
        )
        paragraph._p.append(shape)


def add_footer(document: Document, soup: BeautifulSoup) -> None:
    name = soup.select_one("aside.sidebar h1")
    paragraph = document.sections[0].footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.left_indent = Emu(SIDEBAR_WIDTH + MAIN_PAD["left"])
    paragraph.paragraph_format.right_indent = MAIN_PAD["right"]

    prefix = f"{clean_text(name.get_text())}  ·  " if name else ""
    add_run(paragraph, f"{prefix}Page ", 7.5, FOOTER)
    add_field(paragraph, "PAGE", 7.5, FOOTER)
    add_run(paragraph, " of ", 7.5, FOOTER)
    add_field(paragraph, "NUMPAGES", 7.5, FOOTER)


# ---------------------------------------------------------------------------
# Sidebar
# ---------------------------------------------------------------------------

def render_sidebar(cell, sidebar: Tag) -> None:
    for child in iter_tag_children(sidebar):
        classes = child.get("class", [])
        if child.name == "h1":
            paragraph = new_paragraph(cell, space_after=4.5, line=27)
            run = add_run(paragraph, clean_text(child.get_text()), 23, SIDEBAR_NAME, bold=True)
            set_letter_spacing(run, 4)
        elif child.name == "p" and "role_title" in classes:
            paragraph = new_paragraph(cell, space_after=3, line=13)
            add_run(paragraph, clean_text(child.get_text()), 9.5, SIDEBAR_ROLE, bold=True)
        elif child.name == "p" and "tagline" in classes:
            paragraph = new_paragraph(cell, space_after=3, line=12)
            add_run(paragraph, clean_text(child.get_text()), 8.5, SIDEBAR_MUTED)
        elif child.name == "h3":
            paragraph = new_paragraph(cell, space_before=12, space_after=6, keep_with_next=True)
            set_paragraph_border(paragraph, bottom=(SIDEBAR_RULE, 4, 3))
            run = add_run(paragraph, clean_text(child.get_text()), 8.5, SIDEBAR_HEADING, bold=True, caps=True)
            set_letter_spacing(run, 21)
        elif child.name == "strong":
            paragraph = new_paragraph(cell, space_before=6.5, space_after=3, keep_with_next=True)
            add_run(paragraph, clean_text(child.get_text()), 8.5, SIDEBAR_SUBHEADING, bold=True)
        elif child.name == "ul" and "contacts" in classes:
            render_contacts(cell, child)
        elif child.name == "ul" and "chips" in classes:
            render_chips(cell, child)
        elif child.name == "ul" and "edu" in classes:
            render_education(cell, child)


def render_contacts(cell, list_tag: Tag) -> None:
    for item in list_tag.find_all("li", recursive=False):
        label = item.find("b")
        link = item.find("a")
        value = link or item.find("span")

        if label is not None:
            paragraph = new_paragraph(cell, keep_with_next=True, line=9.5)
            run = add_run(paragraph, clean_text(label.get_text()).rstrip(":"), 6.8, SIDEBAR_MUTED, bold=True, caps=True)
            set_letter_spacing(run, 15)

        paragraph = new_paragraph(cell, space_after=4.5, line=11)
        text = clean_text(value.get_text()) if value is not None else clean_text(item.get_text())
        if link is not None and link.get("href"):
            add_hyperlink(paragraph, link["href"], text, 8.5, SIDEBAR_VALUE)
        else:
            add_run(paragraph, text, 8.5, SIDEBAR_VALUE)


def render_chips(cell, list_tag: Tag) -> None:
    paragraph = new_paragraph(cell, space_after=1, line=18)
    items = list_tag.find_all("li", recursive=False)
    for index, item in enumerate(items):
        # Thin spaces pad the outline; non-breaking spaces as padding make Word break
        # chips mid-word, so they are only used inside multi-word chips
        label = " " + clean_text(item.get_text()).replace(" ", " ") + " "
        run = add_run(paragraph, label, 7.8, SIDEBAR_TEXT)
        # Outline only: Word paints run shading taller than the border, which looks misaligned
        set_run_property(run, "w:bdr", val="single", sz=4, space=1, color=CHIP_BORDER)
        if index < len(items) - 1:
            add_run(paragraph, "  ", 7.8, SIDEBAR_TEXT)


def render_education(cell, list_tag: Tag) -> None:
    for item in list_tag.find_all("li", recursive=False):
        parts = []
        for node in item.children:
            if isinstance(node, Tag) and "degree" in node.get("class", []):
                parts.append(("degree", clean_text(node.get_text())))
            elif isinstance(node, Tag) and "score" in node.get("class", []):
                parts.append(("score", clean_text(node.get_text())))
            else:
                text = clean_text(node.get_text() if isinstance(node, Tag) else str(node))
                if text:
                    parts.append(("text", text))

        styles = {"degree": (SIDEBAR_NAME, True), "score": (SIDEBAR_SUBHEADING, True), "text": (SIDEBAR_EDU, False)}
        for index, (kind, text) in enumerate(parts):
            last = index == len(parts) - 1
            paragraph = new_paragraph(cell, space_after=6 if last else 0, line=11.3, keep_with_next=not last)
            color, bold = styles[kind]
            add_run(paragraph, text, 8.3, color, bold=bold)


# ---------------------------------------------------------------------------
# Main column
# ---------------------------------------------------------------------------

def render_main(cell, main: Tag) -> None:
    first_heading = True
    for child in iter_tag_children(main):
        classes = child.get("class", [])
        if child.name == "h2":
            add_section_heading(cell, clean_text(child.get_text()), first=first_heading)
            first_heading = False
        elif child.name == "div" and "summary" in classes:
            for paragraph_tag in child.find_all("p", recursive=False):
                paragraph = new_paragraph(cell, space_after=7, line=14.4)
                add_inline(paragraph, paragraph_tag, 9.3, BODY)
        elif child.name == "div" and "stats" in classes:
            add_stats(cell, child)
        elif child.name == "div" and "project" in classes:
            add_project_card(cell, child)
        elif child.name == "div" and "block" in classes:
            add_entry(cell, child, MAIN_CONTENT_WIDTH)


def add_section_heading(cell, text: str, first: bool) -> None:
    paragraph = new_paragraph(cell, space_before=0 if first else 10, space_after=6, keep_with_next=True)
    set_paragraph_border(paragraph, bottom=(MAIN_RULE, 12, 3))
    add_run(paragraph, "■", 7, ACCENT)
    run = add_run(paragraph, f"  {text}", 11, HEADING, bold=True, caps=True)
    set_letter_spacing(run, 21)


def add_stats(cell, stats: Tag) -> None:
    items = stats.select(".stat")
    if not items:
        return

    count = len(items)
    stat_width = int((MAIN_CONTENT_WIDTH - STAT_GAP * (count - 1)) / count)
    widths = []
    for index in range(count):
        widths.append(stat_width)
        if index < count - 1:
            widths.append(STAT_GAP)

    table = cell.add_table(rows=1, cols=len(widths))
    prepare_table(table, widths)
    for index, item in enumerate(items):
        stat_cell = table.rows[0].cells[index * 2]
        set_cell_margins(stat_cell, top=Mm(1.6), left=Mm(1), bottom=Mm(1.3), right=Mm(1))
        set_cell_shading(stat_cell, STAT_FILL)
        set_cell_borders(stat_cell, bottom=(ACCENT, 20))
        stat_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        number = item.find("b")
        label = item.find("span")
        paragraph = new_paragraph(stat_cell, line=18, align=WD_ALIGN_PARAGRAPH.CENTER)
        add_run(paragraph, clean_text(number.get_text()) if number else "", 15, ACCENT_DARK, bold=True)
        paragraph = new_paragraph(stat_cell, line=10, align=WD_ALIGN_PARAGRAPH.CENTER)
        run = add_run(paragraph, clean_text(label.get_text()) if label else "", 6.6, STAT_LABEL, bold=True, caps=True)
        set_letter_spacing(run, 4)

    for index in range(1, len(widths), 2):
        set_cell_margins(table.rows[0].cells[index], top=0, left=0, bottom=0, right=0)
        style_spacer(table.rows[0].cells[index].add_paragraph())

    style_spacer(cell.paragraphs[-1], after=4)


def add_project_card(cell, block: Tag) -> None:
    table = cell.add_table(rows=1, cols=1)
    prepare_table(table, [MAIN_CONTENT_WIDTH])
    keep_row_together(table.rows[0])

    card = table.rows[0].cells[0]
    set_cell_margins(card, **CARD_PAD)
    set_cell_shading(card, CARD_FILL)
    set_cell_borders(card, left=(ACCENT, 24))

    inner_width = MAIN_CONTENT_WIDTH - CARD_PAD["left"] - CARD_PAD["right"] - Mm(1)
    add_entry(card, block, inner_width, pill_links=True, in_card=True)
    style_spacer(cell.paragraphs[-1], after=5)


def add_entry(cell, block: Tag, width, pill_links: bool = False, in_card: bool = False) -> None:
    """One experience/project/list block; every paragraph keeps with the next so the
    block never splits across pages (mirrors page-break-inside: avoid)."""
    children = [child for child in iter_tag_children(block)]
    paragraphs = []

    for child in children:
        classes = child.get("class", [])
        if child.name == "h3":
            paragraph = new_paragraph(cell, line=14)
            paragraph.paragraph_format.tab_stops.add_tab_stop(Emu(int(width)), WD_TAB_ALIGNMENT.RIGHT)
            date = child.select_one(".date")
            title = child.get_text().replace(date.get_text(), "") if date is not None else child.get_text()
            add_run(paragraph, clean_text(title), 10 if in_card else 10.2, HEADING, bold=True)
            if date is not None:
                add_run(paragraph, "\t", 8.3, DATE)
                date_text = clean_text(date.get_text())
                if date.name == "a" and date.get("href"):
                    fill = PILL_FILL if pill_links else None
                    label = f"  {date_text}  " if pill_links else date_text
                    add_hyperlink(paragraph, date["href"], label, 7.8, ACCENT_DARK, bold=True, fill=fill)
                else:
                    add_run(paragraph, date_text, 8.3, DATE, bold=True)
            paragraphs.append(paragraph)
        elif child.name == "span" and "role" in classes:
            paragraph = new_paragraph(cell, space_before=1, space_after=1.5, line=12)
            add_run(paragraph, clean_text(child.get_text()), 9, ACCENT_DARK, bold=True)
            paragraphs.append(paragraph)
        elif child.name == "span" and "stack" in classes:
            paragraph = new_paragraph(cell, line=11)
            add_run(paragraph, clean_text(child.get_text()), 7.8, STACK, bold=True)
            paragraphs.append(paragraph)
        elif child.name == "p":
            paragraph = new_paragraph(cell, space_before=2, space_after=2, line=12.9)
            add_inline(paragraph, child, 8.9, BODY)
            paragraphs.append(paragraph)
        elif child.name == "ul":
            for item in child.find_all("li", recursive=False):
                paragraph = new_paragraph(cell, space_after=1, line=12.9)
                paragraph.paragraph_format.left_indent = Mm(3.5)
                paragraph.paragraph_format.first_line_indent = -Mm(3.5)
                add_run(paragraph, "•\t", 8.9, ACCENT)
                add_inline(paragraph, item, 8.9, BODY)
                paragraphs.append(paragraph)

    for paragraph in paragraphs[:-1]:
        paragraph.paragraph_format.keep_with_next = True
    if paragraphs and not in_card:
        paragraphs[-1].paragraph_format.space_after = Pt(7.5)


def add_inline(paragraph, tag: Tag, size: float, color: str) -> None:
    """Write a tag's mixed content, keeping <b>/<strong> bold and <a> clickable."""
    segments = []

    def walk(node, bold=False, href=None, em=False):
        if isinstance(node, NavigableString):
            segments.append([str(node), bold, href, em])
            return
        if not isinstance(node, Tag):
            return
        if node.name in ("b", "strong"):
            bold = True
        elif node.name == "a":
            href = node.get("href")
        elif node.name == "em":
            em = True
        for child in node.children:
            walk(child, bold, href, em)

    for child in tag.children:
        walk(child)

    for segment in segments:
        segment[0] = re.sub(r"\s+", " ", fix_ligatures(segment[0]))
    while segments and not segments[0][0].strip():
        segments.pop(0)
    while segments and not segments[-1][0].strip():
        segments.pop()
    if not segments:
        return
    segments[0][0] = segments[0][0].lstrip()
    segments[-1][0] = segments[-1][0].rstrip()

    for text, bold, href, em in segments:
        if not text:
            continue
        run_size = 8 if em else size
        run_color = STACK if em else (HEADING if bold else color)
        if href:
            add_hyperlink(paragraph, href, text, run_size, ACCENT_DARK, bold=bold or em)
        else:
            add_run(paragraph, text, run_size, run_color, bold=bold or em)


# ---------------------------------------------------------------------------
# Paragraph and run helpers
# ---------------------------------------------------------------------------

def new_paragraph(cell, space_before: float = 0, space_after: float = 0, line: float | None = None,
                  keep_with_next: bool = False, align=None):
    paragraph = cell.add_paragraph()
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(space_before)
    paragraph_format.space_after = Pt(space_after)
    if line is not None:
        paragraph_format.line_spacing = Pt(line)
    if keep_with_next:
        paragraph_format.keep_with_next = True
    if align is not None:
        paragraph.alignment = align
    return paragraph


def add_run(paragraph, text: str, size: float, color: str, bold: bool = False, caps: bool = False):
    run = paragraph.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    run.font.bold = bold
    if caps:
        run.font.all_caps = True
    return run


def add_hyperlink(paragraph, url: str, text: str, size: float, color: str, bold: bool = False, fill: str | None = None):
    run = add_run(paragraph, text, size, color, bold=bold)
    if fill:
        set_run_property(run, "w:shd", val="clear", color="auto", fill=fill)

    relationship_id = paragraph.part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), relationship_id)
    paragraph._p.remove(run._r)
    hyperlink.append(run._r)
    paragraph._p.append(hyperlink)
    return run


def add_field(paragraph, instruction: str, size: float, color: str) -> None:
    def field_char(kind):
        run = add_run(paragraph, "", size, color)
        element = OxmlElement("w:fldChar")
        element.set(qn("w:fldCharType"), kind)
        run._r.append(element)

    field_char("begin")
    run = add_run(paragraph, "", size, color)
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    run._r.append(instr)
    field_char("separate")
    add_run(paragraph, "1", size, color)
    field_char("end")


def style_spacer(paragraph, after: float = 0) -> None:
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(after)
    paragraph_format.line_spacing = Pt(1)
    marker = OxmlElement("w:rPr")
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "2")
    marker.append(size)
    insert_ordered(paragraph._p.get_or_add_pPr(), marker, PPR_ORDER)


def set_letter_spacing(run, twentieths_of_point: int) -> None:
    set_run_property(run, "w:spacing", val=twentieths_of_point)


def set_run_property(run, tag: str, **attributes) -> None:
    element = OxmlElement(tag)
    for name, value in attributes.items():
        element.set(qn(f"w:{name}"), str(value))
    insert_ordered(run._r.get_or_add_rPr(), element, RPR_ORDER)


def set_paragraph_border(paragraph, **edges) -> None:
    borders = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        if edge in edges:
            color, size, space = edges[edge]
            element = OxmlElement(f"w:{edge}")
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), str(size))
            element.set(qn("w:space"), str(space))
            element.set(qn("w:color"), color)
            borders.append(element)
    insert_ordered(paragraph._p.get_or_add_pPr(), borders, PPR_ORDER)


def insert_ordered(parent, element, order) -> None:
    """Replace any existing child with the same tag, keeping Word's schema order."""
    existing = parent.find(element.tag)
    if existing is not None:
        parent.remove(existing)
    tag = element.tag
    prefixed = next(name for name in order if qn(name) == tag)
    parent.insert_element_before(element, *order[order.index(prefixed) + 1:])


# ---------------------------------------------------------------------------
# Table helpers
# ---------------------------------------------------------------------------

def prepare_table(table, widths) -> None:
    """Borderless, fixed-layout table with exact column widths and no indent."""
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    table_properties = table._tbl.tblPr

    total = sum(int(width) for width in widths)
    table_width = OxmlElement("w:tblW")
    table_width.set(qn("w:w"), str(twips(total)))
    table_width.set(qn("w:type"), "dxa")
    insert_ordered(table_properties, table_width, TBLPR_ORDER)

    indent = OxmlElement("w:tblInd")
    indent.set(qn("w:w"), "0")
    indent.set(qn("w:type"), "dxa")
    insert_ordered(table_properties, indent, TBLPR_ORDER)

    borders = OxmlElement("w:tblBorders")
    for edge in EDGES:
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:val"), "nil")
        borders.append(element)
    insert_ordered(table_properties, borders, TBLPR_ORDER)

    cell_margins = OxmlElement("w:tblCellMar")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:w"), "0")
        element.set(qn("w:type"), "dxa")
        cell_margins.append(element)
    insert_ordered(table_properties, cell_margins, TBLPR_ORDER)

    for column, width in zip(table.columns, widths):
        column.width = width
    for row in table.rows:
        for cell, width in zip(row.cells, widths):
            cell.width = width
            cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            clear_cell(cell)


def keep_row_together(row) -> None:
    row_properties = row._tr.get_or_add_trPr()
    row_properties.append(OxmlElement("w:cantSplit"))


def set_cell_margins(cell, top, left, bottom, right) -> None:
    margins = OxmlElement("w:tcMar")
    for edge, value in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        element = OxmlElement(f"w:{edge}")
        element.set(qn("w:w"), str(twips(value)))
        element.set(qn("w:type"), "dxa")
        margins.append(element)
    insert_ordered(cell._tc.get_or_add_tcPr(), margins, TCPR_ORDER)


def set_cell_shading(cell, fill: str) -> None:
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    insert_ordered(cell._tc.get_or_add_tcPr(), shading, TCPR_ORDER)


def set_cell_borders(cell, **edges) -> None:
    borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        if edge in edges:
            color, size = edges[edge]
            element.set(qn("w:val"), "single")
            element.set(qn("w:sz"), str(size))
            element.set(qn("w:space"), "0")
            element.set(qn("w:color"), color)
        else:
            element.set(qn("w:val"), "nil")
        borders.append(element)
    insert_ordered(cell._tc.get_or_add_tcPr(), borders, TCPR_ORDER)


def clear_cell(cell) -> None:
    for paragraph in list(cell.paragraphs):
        element = paragraph._element
        element.getparent().remove(element)


# ---------------------------------------------------------------------------
# Text and file helpers
# ---------------------------------------------------------------------------

def twips(length) -> int:
    return int(Emu(int(length)).twips)


def iter_tag_children(parent: Tag):
    for child in parent.children:
        if isinstance(child, Tag):
            yield child


def fix_ligatures(text: str) -> str:
    return text.translate(str.maketrans({"ﬁ": "fi", "ﬂ": "fl", " ": " "}))


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", fix_ligatures(text)).strip()


def save_document(document: Document, output_file: Path) -> Path:
    try:
        document.save(output_file)
        return output_file
    except PermissionError:
        # The file is probably open in Word; write next to it instead
        fallback_file = next_available_path(output_file)
        document.save(fallback_file)
        return fallback_file


def next_available_path(path: Path) -> Path:
    for index in range(1, 100):
        candidate = path.with_name(f"{path.stem}-{index}{path.suffix}")
        if not candidate.exists():
            return candidate

    raise RuntimeError(f"Could not find a free output filename for {path.name}.")


if __name__ == "__main__":
    html_to_docx()
