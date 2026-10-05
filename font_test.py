"""Print a compact, span-by-span PDF text and font diagnostic."""

import argparse
from pathlib import Path

import fitz


def inspect_pdf(pdf_path: str | Path) -> None:
    source = Path(pdf_path)
    if not source.is_file():
        raise FileNotFoundError(f"找不到 PDF 文件：{source}")
    with fitz.open(source) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            print(f"\n{'=' * 16}\n第 {page_number} 页\n{'=' * 16}")
            for block in page.get_text("dict", sort=True).get("blocks", []):
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        text = span.get("text", "")
                        if text.strip():
                            print(
                                f"文字: {text!r} | 字体: {span.get('font', '')} | "
                                f"字号: {float(span.get('size', 0)):.2f} | "
                                f"颜色: #{int(span.get('color', 0)):06X}"
                            )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="逐段查看 PDF 文本与字体信息")
    parser.add_argument("pdf", nargs="?", default="test.pdf", help="PDF 路径（默认：test.pdf）")
    inspect_pdf(parser.parse_args().pdf)
