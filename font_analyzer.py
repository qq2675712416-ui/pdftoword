"""Report the fonts and approximate text roles found in a PDF."""

import argparse
from pathlib import Path

import fitz

from main import _word_font


def judge_type(text: str, size: float, font: str) -> str:
    stripped = text.strip()
    if stripped.isdigit() and size < 10:
        return "页码"
    if size >= 20:
        return "大标题"
    if size >= 14 or "bold" in font.lower() or "simhei" in font.lower():
        return "标题"
    return "正文"


def analyze_pdf(pdf_path: str | Path) -> None:
    source = Path(pdf_path)
    if not source.is_file():
        raise FileNotFoundError(f"找不到 PDF 文件：{source}")

    with fitz.open(source) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            print(f"\n{'=' * 16}\n第 {page_number} 页\n{'=' * 16}")
            for block in page.get_text("dict", sort=True).get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "").strip()
                        if not text:
                            continue
                        font = span.get("font", "")
                        size = float(span.get("size", 0))
                        print(
                            f"文字: {text[:80]}\n"
                            f"PDF 字体: {font}\n"
                            f"Word 字体: {_word_font(font) or '默认字体'}\n"
                            f"字号: {size:.2f}\n"
                            f"类型: {judge_type(text, size, font)}\n"
                        )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="分析 PDF 的字体、字号和文字类型")
    parser.add_argument("pdf", nargs="?", default="test.pdf", help="PDF 路径（默认：test.pdf）")
    analyze_pdf(parser.parse_args().pdf)
