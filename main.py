"""Convert text-based PDF files to editable Word documents.

The converter keeps page boundaries, approximate text order, basic paragraph
alignment, and span-level font emphasis. It is intended for digitally-created
PDFs; scanned documents need OCR before conversion.
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from io import BytesIO
from pathlib import Path
from typing import Any

import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import qn

LOGGER = logging.getLogger("pdf_to_word")
EDITABLE_TOP_MARGIN_PT = 36.0  # Leave room for a Word header while preserving PDF page coordinates.

# Common PDF font names mapped to fonts likely to exist in Microsoft Word.
FONT_MAP = {
    "simsun": "宋体",
    "nsimsun": "新宋体",
    "simhei": "黑体",
    "simkai": "楷体",
    "kaiti": "楷体",
    "fangsong": "仿宋",
    "timesnewroman": "Times New Roman",
    "arial": "Arial",
    "calibri": "Calibri",
    "courier": "Courier New",
}


def _word_font(pdf_font: str) -> str | None:
    """Return a known Word font name, or None to use the document default."""
    normalized = re.sub(r"[^a-z0-9]", "", pdf_font.lower())
    for source, target in FONT_MAP.items():
        if source in normalized:
            return target
    return None


def _alignment(line: dict[str, Any], page_width: float) -> int:
    """Infer simple left/center/right alignment from a line's bounding box."""
    x0, _, x1, _ = line["bbox"]
    left_gap = max(0.0, x0)
    right_gap = max(0.0, page_width - x1)
    if abs(left_gap - right_gap) < page_width * 0.08 and min(left_gap, right_gap) > page_width * 0.12:
        return WD_ALIGN_PARAGRAPH.CENTER
    if right_gap < page_width * 0.04 and left_gap > page_width * 0.35:
        return WD_ALIGN_PARAGRAPH.RIGHT
    return WD_ALIGN_PARAGRAPH.LEFT


def _format_run(run: Any, span: dict[str, Any]) -> None:
    run.font.size = Pt(max(5, min(float(span.get("size", 11)), 72)))
    font_name = _word_font(span.get("font", ""))
    if font_name:
        run.font.name = font_name
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), font_name)
    flags = int(span.get("flags", 0))
    run.bold = bool(flags & 16) or "bold" in span.get("font", "").lower()
    run.italic = bool(flags & 2) or "italic" in span.get("font", "").lower()
    color = span.get("color")
    if isinstance(color, int):
        run.font.color.rgb = RGBColor((color >> 16) & 255, (color >> 8) & 255, color & 255)


def _line_text(line: dict[str, Any]) -> str:
    return "".join(span.get("text", "") for span in line.get("spans", [])).strip()


def _heading_style(line: dict[str, Any]) -> str | None:
    text = re.sub(r"\s+", "", _line_text(line))
    if re.match(r"^\d+(?:\.\d+){2,}(?=\D)", text):
        return "Heading 3"
    if re.match(r"^\d+\.\d+(?=\D)", text):
        return "Heading 2"
    if re.match(r"^\d{1,2}[\u4e00-\u9fff]", text):
        return "Heading 1"
    if text in {"引言", "结论", "参考文献", "致谢", "附录"}:
        return "Heading 1"
    return None


def _extract_toc_entries(pdf: fitz.Document) -> list[tuple[int, str, str, int]]:
    entries: list[tuple[int, str, str, int]] = []
    in_toc = False
    pattern = re.compile(r"^(.*?)\s*[.．。·…]{2,}\s*(\d{1,3})\s*$")
    for page in pdf:
        lines = [
            line
            for block in page.get_text("dict", sort=True).get("blocks", [])
            for line in block.get("lines", [])
            if any(span.get("text", "").strip() for span in line.get("spans", []))
        ]
        for line in lines:
            text = _line_text(line)
            normalized = re.sub(r"\s+", "", text)
            if normalized == "目录":
                in_toc = True
                continue
            if not in_toc:
                continue
            match = pattern.match(text)
            if match:
                label, page_number = match.group(1).strip(), match.group(2)
                compact = re.sub(r"\s+", "", label)
                if re.match(r"^\d+\.\d+\.\d+", compact):
                    level = 3
                elif re.match(r"^\d+\.\d+", compact):
                    level = 2
                else:
                    level = 1
                entries.append((level, label, page_number, page.number))
                continue
            if re.fullmatch(r"\d+", normalized):
                continue
            return entries
    return entries


def _add_toc_field(document: Document, entries: list[tuple[int, str, str, int]]) -> None:
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = ' TOC \\o "1-3" \\h \\z \\u '
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")

    if not entries:
        entries = [(1, "请在 Word 中更新目录域", "", 0)]
    previous_source_page = entries[0][3]
    for index, (level, label, page_number, source_page) in enumerate(entries):
        if index and source_page != previous_source_page:
            # Start the continuation at the same vertical position as the PDF.
            spacer = document.add_paragraph()
            spacer.paragraph_format.page_break_before = True
            spacer.paragraph_format.space_before = Pt(0)
            spacer.paragraph_format.space_after = Pt(0)
            spacer.paragraph_format.line_spacing = Pt(40)
            spacer.add_run(" ").font.size = Pt(1)
        paragraph = document.add_paragraph(style=document.styles[f"TOC {level}"])
        if index and source_page != previous_source_page:
            paragraph.paragraph_format.space_before = Pt(0)
        previous_source_page = source_page
        if index == 0:
            paragraph.paragraph_format.space_before = Pt(30)
            run = paragraph.add_run()
            run._r.extend((begin, instruction, separate))
        paragraph.add_run(f"{label}\t{page_number}".rstrip())
        if index == len(entries) - 1:
            paragraph.add_run()._r.append(end)


def _configure_toc_styles(document: Document) -> None:
    """Give generated TOC entries comfortable margins and consistent spacing."""
    for level, left_indent in ((1, 74), (2, 94), (3, 114)):
        name = f"TOC {level}"
        try:
            style = document.styles[name]
        except KeyError:
            style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        # Override Word's latent built-in TOC styles; otherwise Word may ignore
        # these paragraph settings when it refreshes the TOC field.
        style._element.attrib.pop(qn("w:customStyle"), None)
        style.font.name = "宋体"
        style.font.size = Pt(12)
        style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "宋体")
        style.paragraph_format.left_indent = Pt(left_indent)
        style.paragraph_format.right_indent = Pt(105)
        style.paragraph_format.space_before = Pt(0)
        # The PDF's TOC rows are spaced about 42 pt apart. Keep that
        # vertical rhythm so the cached entries occupy the same two pages.
        style.paragraph_format.space_after = Pt(24)
        style.paragraph_format.line_spacing = 1.15
        style.paragraph_format.tab_stops.add_tab_stop(Pt(490), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)


def _configure_heading_styles(document: Document) -> None:
    """Keep Word's built-in outline IDs while showing localized heading names."""
    for level in (1, 2, 3):
        style = document.styles.get_by_id(f"Heading{level}", WD_STYLE_TYPE.PARAGRAPH)
        style.name = f"标题 {level}"
        style.font.name = "黑体"
        style.font.color.rgb = RGBColor(0, 0, 0)
        style._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "黑体")


def _add_line(
    document: Document, line: dict[str, Any], previous_bottom: float, style_name: str | None = None
) -> float:
    spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
    if not spans:
        return previous_bottom

    paragraph = document.add_paragraph()
    if style_name:
        if style_name.startswith("Heading "):
            level = style_name.rsplit(" ", 1)[1]
            paragraph.style = document.styles.get_by_id(f"Heading{level}", WD_STYLE_TYPE.PARAGRAPH)
        else:
            paragraph.style = style_name
    x0, y0, _x1, y1 = line["bbox"]
    paragraph.paragraph_format.left_indent = Inches(max(0.0, x0) / 72)
    paragraph.paragraph_format.space_before = Pt(max(0.0, y0 - previous_bottom))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(max(1.0, y1 - y0))

    for span in spans:
        run = paragraph.add_run(span.get("text", ""))
        _format_run(run, span)
    return max(previous_bottom, float(y1))


def _table_column_edges(table: Any) -> list[float]:
    x0, _y0, x1, _y1 = table.bbox
    edges = [float(x0), float(x1)]
    wide_cells = []
    for row in table.rows:
        for cell in row.cells:
            if cell is not None and cell[2] - cell[0] >= 20:
                wide_cells.append(cell)
    internal = sorted(
        edge
        for cell in wide_cells
        for edge in (float(cell[0]), float(cell[2]))
        if x0 + 12 < edge < x1 - 12
    )
    clusters: list[list[float]] = []
    for edge in internal:
        if not clusters or edge - clusters[-1][-1] > 12:
            clusters.append([edge])
        else:
            clusters[-1].append(edge)
    edges.extend(sum(cluster) / len(cluster) for cluster in clusters if len(cluster) >= 2)
    return sorted(set(round(edge, 2) for edge in edges))


def _add_pdf_table(document: Document, table_data: Any, page: fitz.Page) -> float:
    """Create an editable Word table from a detected PDF table and its text spans."""
    x0, y0, x1, y1 = table_data.bbox
    column_edges = _table_column_edges(table_data)
    column_count = len(column_edges) - 1
    if column_count < 1:
        return float(y1)

    word_table = document.add_table(rows=table_data.row_count, cols=column_count)
    word_table.style = "Table Grid"
    word_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    word_table.autofit = False
    total_width_twips = round((x1 - x0) * 20)
    table_width = word_table._tbl.tblPr.find(qn("w:tblW"))
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), str(total_width_twips))
    tbl_ind = OxmlElement("w:tblInd")
    tbl_ind.set(qn("w:w"), str(round(x0 * 20)))
    tbl_ind.set(qn("w:type"), "dxa")
    word_table._tbl.tblPr.append(tbl_ind)
    for index, cell in enumerate(word_table.columns):
        cell.width = Inches((column_edges[index + 1] - column_edges[index]) / 72)

    row_bounds = []
    for row_index, source_row in enumerate(table_data.rows):
        boxes = [cell for cell in source_row.cells if cell is not None]
        row_top = min(box[1] for box in boxes)
        row_bottom = max(box[3] for box in boxes)
        row_bounds.append((row_top, row_bottom))
        word_row = word_table.rows[row_index]
        word_row.height = Pt(max(1, row_bottom - row_top))
        word_row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY

        # Merge cells where the PDF cell spans more than one logical column.
        for box in boxes:
            if box[2] - box[0] < 20:
                continue
            covered = [
                column
                for column in range(column_count)
                if min(box[2], column_edges[column + 1]) - max(box[0], column_edges[column]) > 8
            ]
            if len(covered) > 1:
                word_row.cells[covered[0]].merge(word_row.cells[covered[-1]])

    column_widths = [column_edges[i + 1] - column_edges[i] for i in range(column_count)]
    for word_row in word_table.rows:
        column = 0
        while column < column_count:
            cell = word_row.cells[column]
            span = max(1, cell.grid_span)
            cell.width = Inches(sum(column_widths[column : column + span]) / 72)
            column += span

    # Remove Word's default cell padding so text placement tracks the PDF cell geometry.
    for row in word_table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            margins = OxmlElement("w:tcMar")
            for side in ("top", "left", "bottom", "right"):
                margin = OxmlElement(f"w:{side}")
                margin.set(qn("w:w"), "0")
                margin.set(qn("w:type"), "dxa")
                margins.append(margin)
            tc_pr.append(margins)
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0

    cell_paragraphs: dict[tuple[int, int], tuple[Any, float]] = {}
    for block in page.get_text("dict", sort=True).get("blocks", []):
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                text = span.get("text", "")
                if not text.strip():
                    continue
                sx0, sy0, sx1, sy1 = span["bbox"]
                cx, cy = (sx0 + sx1) / 2, (sy0 + sy1) / 2
                if not (x0 <= cx <= x1 and y0 <= cy <= y1):
                    continue
                row_index = next(
                    (i for i, (top, bottom) in enumerate(row_bounds) if top <= cy <= bottom), None
                )
                col_index = next(
                    (i for i in range(column_count) if column_edges[i] <= cx <= column_edges[i + 1]), None
                )
                if row_index is None or col_index is None:
                    continue
                key = (row_index, col_index)
                previous = cell_paragraphs.get(key)
                if previous and abs(previous[1] - sy0) < 2:
                    paragraph = previous[0]
                elif previous:
                    paragraph = word_table.cell(row_index, col_index).add_paragraph()
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.0
                    cell_paragraphs[key] = (paragraph, sy0)
                else:
                    paragraph = word_table.cell(row_index, col_index).paragraphs[0]
                    cell_paragraphs[key] = (paragraph, sy0)
                _format_run(paragraph.add_run(text), span)
    return float(y1)


def pdf_to_word(pdf_path: str | Path, word_path: str | Path) -> Path:
    """Convert a text-based PDF into a DOCX and return the output path."""
    source = Path(pdf_path).expanduser()
    destination = Path(word_path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"找不到 PDF 文件：{source}")
    if source.resolve() == destination.resolve():
        raise ValueError("输入文件和输出文件不能是同一个文件。")
    if destination.suffix.lower() != ".docx":
        raise ValueError("输出文件必须使用 .docx 扩展名。")

    destination.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    _configure_heading_styles(document)
    _configure_toc_styles(document)
    document.core_properties.title = source.stem
    update_fields = document.settings._element.find(qn("w:updateFields"))
    if update_fields is None:
        update_fields = OxmlElement("w:updateFields")
        document.settings._element.append(update_fields)
    # Preserve the source PDF's TOC text and page numbers when Word opens the
    # file. The TOC remains a real field and can still be updated manually.
    update_fields.set(qn("w:val"), "false")
    section = document.sections[0]
    section.top_margin = Pt(EDITABLE_TOP_MARGIN_PT)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.header_distance = Inches(0.25)

    with fitz.open(source) as pdf:
        if pdf.is_encrypted:
            raise ValueError("PDF 已加密，当前版本无法转换。请先解除密码保护。")
        if len(pdf) == 0:
            raise ValueError("PDF 不包含页面。")
        toc_entries = _extract_toc_entries(pdf)

        # Match the first PDF page's dimensions and orientation in Word.
        first_page = pdf[0]
        section.page_width = Inches(first_page.rect.width / 72)
        section.page_height = Inches(first_page.rect.height / 72)

        toc_active = False
        for page_number, page in enumerate(pdf, start=1):
            LOGGER.info("正在处理第 %d/%d 页", page_number, len(pdf))
            blocks = page.get_text("dict", sort=True).get("blocks", [])
            detected_tables = page.find_tables().tables
            table_rects = [fitz.Rect(table.bbox) for table in detected_tables]
            page_lines = [
                line
                for block in blocks
                for line in block.get("lines", [])
                if any(span.get("text", "").strip() for span in line.get("spans", []))
            ]
            items: list[tuple[float, str, Any]] = []
            page_has_toc_title = False
            page_has_toc_entries = False
            for text_line in page_lines:
                line_box = fitz.Rect(text_line["bbox"])
                normalized_text = re.sub(r"\s+", "", _line_text(text_line))
                if normalized_text == "目录":
                    toc_active = True
                    page_has_toc_title = True
                elif toc_active:
                    if "…" in normalized_text or re.fullmatch(r"\d+", normalized_text):
                        page_has_toc_entries = True
                        continue  # Replace static PDF contents entries with the Word TOC field.
                    toc_active = False
                center = fitz.Point((line_box.x0 + line_box.x1) / 2, (line_box.y0 + line_box.y1) / 2)
                if not any(rect.contains(center) for rect in table_rects):
                    items.append((line_box.y0, "line", text_line))
            items.extend((float(table.bbox[1]), "table", table) for table in detected_tables)
            items.sort(key=lambda item: item[0])

            if not items and not page_has_toc_entries:
                LOGGER.warning("第 %d 页没有可提取的文本，可能是扫描页。", page_number)
            previous_bottom = EDITABLE_TOP_MARGIN_PT
            for _top, kind, item in items:
                if kind == "line":
                    text_norm = re.sub(r"\s+", "", _line_text(item))
                    style_name = None if text_norm == "目录" else _heading_style(item)
                    if text_norm == "目录":
                        # Word suppresses space-before on the first paragraph
                        # after a page break. Use an explicit spacer to match
                        # the title's original PDF position.
                        toc_spacer = document.add_paragraph()
                        toc_spacer.paragraph_format.space_before = Pt(0)
                        toc_spacer.paragraph_format.space_after = Pt(0)
                        toc_spacer.paragraph_format.line_spacing = Pt(56)
                        toc_spacer.add_run(" ").font.size = Pt(1)
                        previous_bottom = _add_line(document, item, 92.0, style_name)
                    else:
                        previous_bottom = _add_line(document, item, previous_bottom, style_name)
                    if text_norm == "目录":
                        _add_toc_field(document, toc_entries)
                else:
                    table_top = float(item.bbox[1])
                    gap = max(0.0, table_top - previous_bottom)
                    if gap > 0.5:
                        spacer = document.add_paragraph()
                        spacer.paragraph_format.space_before = Pt(0)
                        spacer.paragraph_format.space_after = Pt(0)
                        spacer.paragraph_format.line_spacing = Pt(gap)
                        spacer.add_run(" ").font.size = Pt(1)
                    previous_bottom = _add_pdf_table(document, item, page)
            if page_number < len(pdf) and not (page_has_toc_entries and not page_has_toc_title):
                document.add_page_break()

    document.save(destination)
    LOGGER.info("转换完成：%s", destination.resolve())
    return destination


def pdf_to_word_visual(pdf_path: str | Path, word_path: str | Path, dpi: int = 300) -> Path:
    """Put a high-resolution rendering of each PDF page into a DOCX.

    This preserves the PDF's appearance; page content is not editable as Word text.
    """
    source = Path(pdf_path).expanduser()
    destination = Path(word_path).expanduser()
    if not source.is_file():
        raise FileNotFoundError(f"找不到 PDF 文件：{source}")
    if source.resolve() == destination.resolve():
        raise ValueError("输入文件和输出文件不能是同一个文件。")
    if destination.suffix.lower() != ".docx":
        raise ValueError("输出文件必须使用 .docx 扩展名。")
    if dpi < 72 or dpi > 1200:
        raise ValueError("渲染分辨率应在 72 到 1200 DPI 之间。")

    destination.parent.mkdir(parents=True, exist_ok=True)
    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0)
    section.bottom_margin = Inches(0)
    section.left_margin = Inches(0)
    section.right_margin = Inches(0)
    section.header_distance = Inches(0)
    section.footer_distance = Inches(0)

    with fitz.open(source) as pdf:
        if pdf.is_encrypted:
            raise ValueError("PDF 已加密，当前版本无法转换。请先解除密码保护。")
        if len(pdf) == 0:
            raise ValueError("PDF 不包含页面。")

        for page_number, page in enumerate(pdf):
            section.page_width = Inches(page.rect.width / 72)
            section.page_height = Inches(page.rect.height / 72)
            if page_number:
                document.add_page_break()
            image = BytesIO(page.get_pixmap(dpi=dpi, alpha=False).tobytes("png"))
            paragraph = document.add_paragraph()
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.add_run().add_picture(
                image, width=Inches(page.rect.width / 72), height=Inches(page.rect.height / 72)
            )

    document.save(destination)
    LOGGER.info("保真版转换完成：%s", destination.resolve())
    return destination


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="将可复制文字的 PDF 转换为可编辑的 Word 文档。扫描版 PDF 请先进行 OCR。"
    )
    parser.add_argument("pdf", nargs="?", default="test.pdf", help="输入 PDF 路径（默认：test.pdf）")
    parser.add_argument("-o", "--output", help="输出 DOCX 路径（默认：与 PDF 同名）")
    parser.add_argument("--visual", action="store_true", help="按页面图像保留 PDF 外观（文字不可直接编辑）")
    parser.add_argument("--dpi", type=int, default=300, help="保真版页面渲染分辨率（默认：300）")
    parser.add_argument("-v", "--verbose", action="store_true", help="显示逐页处理信息")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    logging.basicConfig(level=logging.INFO if args.verbose else logging.WARNING, format="%(message)s")
    output = Path(args.output) if args.output else Path(args.pdf).with_suffix(".docx")
    try:
        if args.visual:
            pdf_to_word_visual(args.pdf, output, args.dpi)
        else:
            pdf_to_word(args.pdf, output)
    except (OSError, ValueError, fitz.FileDataError) as exc:
        LOGGER.error("转换失败：%s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
