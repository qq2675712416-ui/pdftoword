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
from types import SimpleNamespace
from typing import Any

import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION_START
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


def _format_run(run: Any, span: dict[str, Any], underline: bool = False) -> None:
    run.font.size = Pt(max(5, min(float(span.get("size", 11)), 72)))
    font_name = _word_font(span.get("font", ""))
    if font_name:
        run.font.name = font_name
        run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), font_name)
    flags = int(span.get("flags", 0))
    run.bold = bool(flags & 16) or "bold" in span.get("font", "").lower()
    run.italic = bool(flags & 2) or "italic" in span.get("font", "").lower()
    if underline:
        run.underline = True
    color = span.get("color")
    if isinstance(color, int):
        run.font.color.rgb = RGBColor((color >> 16) & 255, (color >> 8) & 255, color & 255)


def _line_text(line: dict[str, Any]) -> str:
    return "".join(span.get("text", "") for span in line.get("spans", [])).strip()


def _horizontal_marks(page: fitz.Page) -> list[fitz.Rect]:
    """Return thin horizontal PDF marks used for rules, fraction bars and underlines."""
    marks: list[fitz.Rect] = []
    for drawing in page.get_drawings():
        rect = fitz.Rect(drawing["rect"])
        if rect.height <= 2.2 and rect.width >= 5:
            marks.append(rect)
        for item in drawing.get("items", []):
            if item[0] != "l":
                continue
            start, end = item[1], item[2]
            if abs(start.y - end.y) <= 1.2 and abs(start.x - end.x) >= 5:
                marks.append(fitz.Rect(min(start.x, end.x), start.y, max(start.x, end.x), start.y + 0.8))
    return marks


def _span_is_underlined(span: dict[str, Any], marks: list[fitz.Rect]) -> bool:
    x0, y0, x1, y1 = span["bbox"]
    text_width = max(1.0, x1 - x0)
    for mark in marks:
        distance = mark.y0 - y1
        if -1.5 <= distance <= 2.5 and mark.width <= text_width * 1.6:
            overlap = max(0.0, min(x1, mark.x1) - max(x0, mark.x0))
            if overlap >= min(text_width * 0.72, mark.width * 0.72):
                return True
    return False


def _is_equation_line(line: dict[str, Any]) -> bool:
    spans = [span for span in line.get("spans", []) if span.get("text", "").strip()]
    total = sum(len(span.get("text", "")) for span in spans)
    math_chars = sum(
        len(span.get("text", ""))
        for span in spans
        if "cambria" in span.get("font", "").lower() or "math" in span.get("font", "").lower()
    )
    text = "".join(span.get("text", "") for span in spans)
    has_operator = any(symbol in text for symbol in "=≈≠≤≥±∑∫√")
    return math_chars >= 2 and total > 0 and (math_chars / total >= 0.34 or has_operator)


def _merge_equation_lines(lines: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Join PDF's separately positioned base/superscript fragments into equation rows."""
    equation_lines = sorted((line for line in lines if _is_equation_line(line)), key=lambda line: line["bbox"][1])
    groups: list[dict[str, Any]] = []
    for line in equation_lines:
        rect = fitz.Rect(line["bbox"])
        target = None
        for group in reversed(groups):
            group_rect = fitz.Rect(group["bbox"])
            vertical_overlap = max(0.0, min(rect.y1, group_rect.y1) - max(rect.y0, group_rect.y0))
            min_height = max(1.0, min(rect.height, group_rect.height))
            horizontal_gap = max(0.0, max(rect.x0, group_rect.x0) - min(rect.x1, group_rect.x1))
            if vertical_overlap / min_height >= 0.22 and horizontal_gap <= 18:
                target = group
                break
        if target is None:
            groups.append({"bbox": tuple(rect), "spans": list(line["spans"]), "_lines": [line]})
        else:
            target["_lines"].append(line)
            target["spans"].extend(line["spans"])
            target["bbox"] = tuple(fitz.Rect(target["bbox"]).include_rect(rect))
    for group in groups:
        group["spans"].sort(key=lambda span: (span["bbox"][0], span["bbox"][1]))
    return groups


def _math_run(text: str) -> Any:
    run = OxmlElement("m:r")
    run_props = OxmlElement("m:rPr")
    style = OxmlElement("m:sty")
    style.set(qn("m:val"), "p")
    run_props.append(style)
    run.append(run_props)
    node = OxmlElement("m:t")
    node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    return run


def _script_math(base: str, subscript: str = "", superscript: str = "") -> Any:
    if subscript and superscript:
        node = OxmlElement("m:sSubSup")
        props = OxmlElement("m:sSubSupPr")
        props.append(OxmlElement("m:ctrlPr"))
        node.append(props)
        base_node, sub_node, sup_node = OxmlElement("m:e"), OxmlElement("m:sub"), OxmlElement("m:sup")
        base_node.append(_math_run(base))
        sub_node.append(_math_run(subscript))
        sup_node.append(_math_run(superscript))
        node.extend((base_node, sub_node, sup_node))
        return node
    if subscript:
        node = OxmlElement("m:sSub")
        props = OxmlElement("m:sSubPr")
        props.append(OxmlElement("m:ctrlPr"))
        node.append(props)
        base_node, script_node = OxmlElement("m:e"), OxmlElement("m:sub")
    else:
        node = OxmlElement("m:sSup")
        props = OxmlElement("m:sSupPr")
        props.append(OxmlElement("m:ctrlPr"))
        node.append(props)
        base_node, script_node = OxmlElement("m:e"), OxmlElement("m:sup")
    base_node.append(_math_run(base))
    script_node.append(_math_run(subscript or superscript))
    node.extend((base_node, script_node))
    return node


def _append_math_content(parent: Any, spans: list[dict[str, Any]]) -> None:
    """Write editable OMML runs and map small positioned PDF glyphs to scripts."""
    nonempty = [span for span in spans if span.get("text", "").strip()]
    if not nonempty:
        return
    base_sizes = [float(span.get("size", 11)) for span in nonempty]
    base_size = sorted(base_sizes)[len(base_sizes) // 2]
    consumed: set[int] = set()
    for index, span in enumerate(nonempty):
        if index in consumed:
            continue
        text = span.get("text", "")
        x0, y0, x1, y1 = span["bbox"]
        subscript = ""
        superscript = ""
        for script_index in range(index + 1, len(nonempty)):
            candidate = nonempty[script_index]
            cx0, cy0, _cx1, _cy1 = candidate["bbox"]
            gap = cx0 - x1
            size = float(candidate.get("size", base_size))
            if gap > 10 or gap < -3 or size > base_size * 0.88:
                if gap > 10:
                    break
                continue
            script_text = candidate.get("text", "").strip()
            if not script_text:
                continue
            if cy0 < y0 - 1.0:
                superscript += script_text
                consumed.add(script_index)
            elif cy0 > y0 + 1.0:
                subscript += script_text
                consumed.add(script_index)
        if subscript or superscript:
            match = re.search(r"([\w𝑎-𝑧𝒜-𝓏𝛼-𝜔∗*]+)(\s*)$", text, flags=re.UNICODE)
            if match:
                prefix = text[:match.start(1)]
                if prefix:
                    parent.append(_math_run(prefix))
                parent.append(_script_math(match.group(1), subscript, superscript))
                if match.group(2):
                    parent.append(_math_run(match.group(2)))
                continue
        parent.append(_math_run(text))


def _add_math_line(
    document: Document, line: dict[str, Any], previous_bottom: float, horizontal_marks: list[fitz.Rect]
) -> float:
    paragraph = document.add_paragraph()
    x0, y0, _x1, y1 = line["bbox"]
    alignment = _alignment(line, document.sections[-1].page_width.pt)
    paragraph.alignment = alignment
    paragraph.paragraph_format.left_indent = Inches(
        max(0.0, x0) / 72 if alignment == WD_ALIGN_PARAGRAPH.LEFT else 0
    )
    paragraph.paragraph_format.space_before = Pt(max(0.0, y0 - previous_bottom))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(max(1.0, y1 - y0))
    equation = OxmlElement("m:oMath")
    spans = line["spans"]
    rect = fitz.Rect(line["bbox"])
    fraction = next(
        (
            mark
            for mark in horizontal_marks
            if 7 <= mark.width <= min(90, rect.width * 0.28)
            and rect.y0 - 1 <= mark.y0 <= rect.y1 + 1
        ),
        None,
    )
    if fraction is not None:
        numerator = [
            span for span in spans
            if (span["bbox"][0] + span["bbox"][2]) / 2 >= fraction.x0 - 4
            and (span["bbox"][0] + span["bbox"][2]) / 2 <= fraction.x1 + 4
            and (span["bbox"][1] + span["bbox"][3]) / 2 < fraction.y0
        ]
        denominator = [
            span for span in spans
            if (span["bbox"][0] + span["bbox"][2]) / 2 >= fraction.x0 - 4
            and (span["bbox"][0] + span["bbox"][2]) / 2 <= fraction.x1 + 4
            and (span["bbox"][1] + span["bbox"][3]) / 2 > fraction.y1
        ]
        if numerator and denominator:
            prefix = [span for span in spans if span["bbox"][2] <= fraction.x0 + 0.5]
            suffix = [span for span in spans if span["bbox"][0] >= fraction.x1 - 0.5]
            used = {id(span) for span in prefix + numerator + denominator + suffix}
            remainder = [span for span in spans if id(span) not in used]
            prefix.extend(span for span in remainder if span["bbox"][0] < fraction.x0)
            suffix.extend(span for span in remainder if span["bbox"][0] >= fraction.x0)
            _append_math_content(equation, sorted(prefix, key=lambda span: span["bbox"][0]))
            fraction_node = OxmlElement("m:f")
            fraction_props = OxmlElement("m:fPr")
            fraction_type = OxmlElement("m:type")
            fraction_type.set(qn("m:val"), "bar")
            fraction_props.append(fraction_type)
            fraction_node.append(fraction_props)
            numerator_node, denominator_node = OxmlElement("m:num"), OxmlElement("m:den")
            _append_math_content(numerator_node, sorted(numerator, key=lambda span: span["bbox"][0]))
            _append_math_content(denominator_node, sorted(denominator, key=lambda span: span["bbox"][0]))
            fraction_node.extend((numerator_node, denominator_node))
            equation.append(fraction_node)
            _append_math_content(equation, sorted(suffix, key=lambda span: span["bbox"][0]))
        else:
            _append_math_content(equation, spans)
    else:
        _append_math_content(equation, spans)
    paragraph._p.append(equation)
    return max(previous_bottom, float(y1))


def _rule_bands(page: fitz.Page) -> list[dict[str, Any]]:
    marks = _horizontal_marks(page)
    bands: list[dict[str, Any]] = []
    for mark in sorted(marks, key=lambda rect: (rect.y0, rect.x0)):
        band = next((item for item in reversed(bands) if abs(item["y"] - mark.y0) <= 1.0), None)
        if band is None:
            bands.append({"y": mark.y0, "segments": [mark]})
        else:
            band["segments"].append(mark)
            band["y"] = sum(rect.y0 for rect in band["segments"]) / len(band["segments"])
    for band in bands:
        band["segments"].sort(key=lambda rect: rect.x0)
        band["x0"] = min(rect.x0 for rect in band["segments"])
        band["x1"] = max(rect.x1 for rect in band["segments"])
    return bands


def _detect_three_line_tables(
    page: fitz.Page, lines: list[dict[str, Any]], existing_rects: list[fitz.Rect]
) -> list[tuple[Any, fitz.Rect]]:
    """Recognize text tables drawn with only top, header and bottom rules."""
    bands = [
        band for band in _rule_bands(page)
        if band["y"] > 70 and band["y"] < page.rect.height - 55
        and band["x1"] - band["x0"] > page.rect.width * 0.52
    ]
    candidates = []
    for top_index, top in enumerate(bands):
        if any(rect.y0 - 2 <= top["y"] <= rect.y1 + 2 for rect in existing_rects):
            continue
        for middle_index in range(top_index + 1, len(bands)):
            middle = bands[middle_index]
            if not 12 <= middle["y"] - top["y"] <= 42:
                continue
            if abs(middle["x0"] - top["x0"]) > 14 or abs(middle["x1"] - top["x1"]) > 14:
                continue
            for bottom in bands[middle_index + 1:]:
                if not 38 <= bottom["y"] - middle["y"] <= 330:
                    continue
                if abs(bottom["x0"] - top["x0"]) > 14 or abs(bottom["x1"] - top["x1"]) > 14:
                    continue
                rect = fitz.Rect(top["x0"], top["y"], top["x1"], bottom["y"])
                if any(rect.intersects(existing) for existing in existing_rects):
                    continue
                row_lines = [
                    line for line in lines
                    if rect.y0 < (line["bbox"][1] + line["bbox"][3]) / 2 < rect.y1
                    and rect.x0 <= (line["bbox"][0] + line["bbox"][2]) / 2 <= rect.x1
                ]
                row_centers: list[list[float]] = []
                for line in sorted(row_lines, key=lambda entry: entry["bbox"][1]):
                    center = (line["bbox"][1] + line["bbox"][3]) / 2
                    group = next((values for values in reversed(row_centers) if abs(values[-1] - center) <= 5), None)
                    if group is None:
                        row_centers.append([center])
                    else:
                        group.append(center)
                if len(row_centers) >= 3:
                    candidates.append((rect, top, middle, bottom, row_lines, row_centers))
    selected: list[tuple[Any, fitz.Rect]] = []
    occupied: list[fitz.Rect] = list(existing_rects)
    for rect, top, middle, bottom, row_lines, row_centers in sorted(candidates, key=lambda item: (item[0].y0, item[0].x0)):
        if any(rect.intersects(other) for other in occupied):
            continue
        gaps = []
        for band in (top, middle, bottom):
            parts = band["segments"]
            for part in parts:
                if 0.45 <= part.width <= 3 and rect.x0 + 8 < (part.x0 + part.x1) / 2 < rect.x1 - 8:
                    gaps.append((part.width, (part.x0 + part.x1) / 2))
            for left, right in zip(parts, parts[1:]):
                gap = right.x0 - left.x1
                if 0.45 <= gap <= 8:
                    gaps.append((gap, (left.x1 + right.x0) / 2))
        edges = [rect.x0, rect.x1]
        if gaps:
            clusters: list[list[float]] = []
            for _gap, edge in sorted(gaps, key=lambda item: item[1]):
                if not clusters or edge - clusters[-1][-1] > 8:
                    clusters.append([edge])
                else:
                    clusters[-1].append(edge)
            edges.extend(sum(cluster) / len(cluster) for cluster in clusters)
        edges = sorted(set(round(edge, 2) for edge in edges))
        if len(edges) < 3:
            continue
        centers = [sum(values) / len(values) for values in row_centers]
        row_bounds = [rect.y0]
        row_bounds.extend((a + b) / 2 for a, b in zip(centers, centers[1:]))
        row_bounds.append(rect.y1)
        row_data = []
        for row_index in range(len(centers)):
            cells = [
                (edges[column], row_bounds[row_index], edges[column + 1], row_bounds[row_index + 1])
                for column in range(len(edges) - 1)
            ]
            row_data.append(SimpleNamespace(cells=cells))
        table = SimpleNamespace(bbox=tuple(rect), row_count=len(row_data), rows=row_data, _three_line=True)
        selected.append((table, rect))
        occupied.append(rect)
    return selected


def _add_pdf_image(document: Document, block: dict[str, Any], previous_bottom: float, page_width: float) -> float:
    x0, y0, x1, y1 = block["bbox"]
    width = min(float(x1 - x0), max(8.0, page_width - float(x0)))
    height = float(y1 - y0)
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.left_indent = Inches(max(0.0, float(x0)) / 72)
    paragraph.paragraph_format.space_before = Pt(max(0.0, float(y0) - previous_bottom))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(max(1.0, height))
    inline = paragraph.add_run().add_picture(BytesIO(block["image"]), width=Inches(width / 72))
    inline.height = Inches(height / 72)
    return max(previous_bottom, float(y1))


def _vector_figure(page: fitz.Page) -> tuple[fitz.Rect, dict[str, Any]] | None:
    """Rasterize compact groups of filled vector shapes such as process diagrams."""
    shapes = [
        drawing for drawing in page.get_drawings()
        if drawing.get("fill") is not None
        and drawing["rect"].width > 8
        and drawing["rect"].height > 6
    ]
    if len(shapes) < 3:
        return None
    rect = fitz.Rect(
        min(drawing["rect"].x0 for drawing in shapes),
        min(drawing["rect"].y0 for drawing in shapes),
        max(drawing["rect"].x1 for drawing in shapes),
        max(drawing["rect"].y1 for drawing in shapes),
    )
    if rect.width < page.rect.width * 0.35 or rect.height < 24:
        return None
    pixmap = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=rect, alpha=False)
    block = {"type": 1, "bbox": tuple(rect), "image": pixmap.tobytes("png")}
    return rect, block


def _add_pdf_rule(document: Document, rule: fitz.Rect, previous_bottom: float, page_width: float) -> float:
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.left_indent = Inches(max(0.0, rule.x0) / 72)
    paragraph.paragraph_format.right_indent = Inches(max(0.0, page_width - rule.x1) / 72)
    paragraph.paragraph_format.space_before = Pt(max(0.0, rule.y0 - previous_bottom))
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = Pt(max(1.0, rule.height))
    paragraph.add_run(" ").font.size = Pt(1)
    p_pr = paragraph._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    border = OxmlElement("w:top")
    border.set(qn("w:val"), "single")
    border.set(qn("w:sz"), "6")
    border.set(qn("w:space"), "0")
    border.set(qn("w:color"), "000000")
    borders.append(border)
    p_pr.append(borders)
    return max(previous_bottom, float(rule.y1))


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
    document: Document, line: dict[str, Any], previous_bottom: float, style_name: str | None = None,
    horizontal_marks: list[fitz.Rect] | None = None,
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
        _format_run(run, span, _span_is_underlined(span, horizontal_marks or []))
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


def _set_table_border(table: Any, edge: str, value: str, size: int = 6) -> None:
    properties = table._tbl.tblPr
    borders = properties.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        properties.append(borders)
    node = borders.find(qn(f"w:{edge}"))
    if node is None:
        node = OxmlElement(f"w:{edge}")
        borders.append(node)
    node.set(qn("w:val"), value)
    if value != "nil":
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")


def _set_cell_border(cell: Any, edge: str, value: str, size: int = 6) -> None:
    properties = cell._tc.get_or_add_tcPr()
    borders = properties.find(qn("w:tcBorders"))
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        properties.append(borders)
    node = borders.find(qn(f"w:{edge}"))
    if node is None:
        node = OxmlElement(f"w:{edge}")
        borders.append(node)
    node.set(qn("w:val"), value)
    if value != "nil":
        node.set(qn("w:sz"), str(size))
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), "000000")


def _apply_three_line_style(word_table: Any) -> None:
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _set_table_border(word_table, edge, "nil")
    for row_index, row in enumerate(word_table.rows):
        for cell in row.cells:
            for edge in ("top", "left", "bottom", "right"):
                _set_cell_border(cell, edge, "nil")
            if row_index == 0:
                _set_cell_border(cell, "top", "single", 12)
                _set_cell_border(cell, "bottom", "single", 6)
            if row_index == len(word_table.rows) - 1:
                _set_cell_border(cell, "bottom", "single", 12)


def _add_pdf_table(
    document: Document, table_data: Any, page: fitz.Page, horizontal_marks: list[fitz.Rect]
) -> float:
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
        word_row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

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
                _format_run(paragraph.add_run(text), span, _span_is_underlined(span, horizontal_marks))
    if getattr(table_data, "_three_line", False):
        _apply_three_line_style(word_table)
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
            three_line_tables = _detect_three_line_tables(page, page_lines, table_rects)
            table_rects.extend(rect for _table, rect in three_line_tables)
            vector_figure = _vector_figure(page)
            vector_rects = [vector_figure[0]] if vector_figure else []
            horizontal_marks = _horizontal_marks(page)
            equation_groups = _merge_equation_lines(
                [
                    line for line in page_lines
                    if not any(rect.contains(fitz.Point(
                        (line["bbox"][0] + line["bbox"][2]) / 2,
                        (line["bbox"][1] + line["bbox"][3]) / 2,
                    )) for rect in table_rects + vector_rects)
                ]
            )
            equation_source_lines = {
                id(source_line)
                for group in equation_groups
                for source_line in group["_lines"]
            }
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
                if not any(rect.contains(center) for rect in table_rects + vector_rects):
                    if id(text_line) not in equation_source_lines:
                        items.append((line_box.y0, "line", text_line))
            items.extend((float(table.bbox[1]), "table", table) for table in detected_tables)
            items.extend((float(table.bbox[1]), "table", table) for table, _rect in three_line_tables)
            for group in equation_groups:
                group["spans"].sort(key=lambda span: (span["bbox"][0], span["bbox"][1]))
                items.append((float(group["bbox"][1]), "math", group))
            for block in blocks:
                if block.get("type") == 1 and block.get("image"):
                    items.append((float(block["bbox"][1]), "image", block))
            if vector_figure:
                items.append((float(vector_figure[1]["bbox"][1]), "image", vector_figure[1]))
            for mark in horizontal_marks:
                if mark.width < page.rect.width * 0.4 or mark.y0 < 35:
                    continue
                if any(rect.x0 - 2 <= mark.x0 and mark.x1 <= rect.x1 + 2
                       and rect.y0 - 2 <= mark.y0 <= rect.y1 + 2 for rect in table_rects):
                    continue
                items.append((float(mark.y0), "rule", mark))
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
                        previous_bottom = _add_line(document, item, 92.0, style_name, horizontal_marks)
                    else:
                        previous_bottom = _add_line(document, item, previous_bottom, style_name, horizontal_marks)
                    if text_norm == "目录":
                        _add_toc_field(document, toc_entries)
                elif kind == "math":
                    previous_bottom = _add_math_line(document, item, previous_bottom, horizontal_marks)
                elif kind == "image":
                    previous_bottom = _add_pdf_image(document, item, previous_bottom, page.rect.width)
                elif kind == "rule":
                    previous_bottom = _add_pdf_rule(document, item, previous_bottom, page.rect.width)
                else:
                    table_top = float(item.bbox[1])
                    gap = max(0.0, table_top - previous_bottom)
                    if gap > 0.5:
                        spacer = document.add_paragraph()
                        spacer.paragraph_format.space_before = Pt(0)
                        spacer.paragraph_format.space_after = Pt(0)
                        spacer.paragraph_format.line_spacing = Pt(gap)
                        spacer.add_run(" ").font.size = Pt(1)
                    previous_bottom = _add_pdf_table(document, item, page, horizontal_marks)
            if page_number < len(pdf):
                next_page = pdf[page_number]
                size_changes = (
                    abs(page.rect.width - next_page.rect.width) > 1
                    or abs(page.rect.height - next_page.rect.height) > 1
                )
                preserve_toc_flow = page_has_toc_entries and not page_has_toc_title
                if size_changes:
                    section = document.add_section(WD_SECTION_START.NEW_PAGE)
                    section.page_width = Inches(next_page.rect.width / 72)
                    section.page_height = Inches(next_page.rect.height / 72)
                    section.top_margin = Pt(EDITABLE_TOP_MARGIN_PT)
                    section.bottom_margin = Inches(0)
                    section.left_margin = Inches(0)
                    section.right_margin = Inches(0)
                    section.header_distance = Inches(0.25)
                elif not preserve_toc_flow:
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
