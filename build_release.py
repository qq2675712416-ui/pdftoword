"""Build a self-contained, no-install Windows ZIP from the active Python runtime."""

from __future__ import annotations

import shutil
import os
import struct
import subprocess
import sys
from pathlib import Path
from sysconfig import get_paths


ROOT = Path(__file__).resolve().parent
RELEASE_ROOT = ROOT / "release"
PACKAGE = RELEASE_ROOT / "PDF转Word"
ARCHIVE = ROOT / "PDF转Word-免安装版.zip"
RUNTIME = PACKAGE / "runtime"


def copy_tree(source: Path, target: Path, ignore: set[str] | None = None) -> None:
    shutil.copytree(source, target, ignore=shutil.ignore_patterns(*(ignore or set())))


def create_icon(path: Path) -> None:
    size = 32
    pixels: list[list[tuple[int, int, int, int]]] = []

    def inside_polygon(x: int, y: int, points: list[tuple[int, int]]) -> bool:
        inside = False
        j = len(points) - 1
        for i, (xi, yi) in enumerate(points):
            xj, yj = points[j]
            if (yi > y) != (yj > y) and x < (xj - xi) * (y - yi) / (yj - yi) + xi:
                inside = not inside
            j = i
        return inside

    page = [(8, 4), (20, 4), (25, 9), (25, 28), (8, 28)]
    fold = [(20, 4), (20, 9), (25, 9)]
    for y in range(size):
        row = []
        for x in range(size):
            dx, dy = max(7 - x, 0, x - 24), max(7 - y, 0, y - 24)
            if dx * dx + dy * dy > 25:
                color = (0, 0, 0, 0)
            else:
                color = (235, 243, 255, 255)  # light blue tile, BGRA below
                if inside_polygon(x, y, page):
                    color = (255, 255, 255, 255)
                if inside_polygon(x, y, fold):
                    color = (201, 220, 255, 255)
                if 11 <= x <= 21 and y in (13, 17, 21):
                    color = (37, 99, 235, 255)
                if 21 <= x <= 28 and 21 <= y <= 28:
                    color = (37, 99, 235, 255)
                if (x in (21, 28) and 21 <= y <= 28) or (y in (21, 28) and 21 <= x <= 28):
                    color = (255, 255, 255, 255)
            r, g, b, a = color
            row.append((b, g, r, a))
        pixels.append(row)

    xor = b"".join(bytes(pixel) for row in reversed(pixels) for pixel in row)
    mask = bytearray()
    for row in reversed(pixels):
        bits = bytearray(4)
        for x, (_b, _g, _r, alpha) in enumerate(row):
            if alpha == 0:
                bits[x // 8] |= 1 << (7 - x % 8)
        mask.extend(bits)
    info = struct.pack("<IiiHHIIiiII", 40, size, size * 2, 1, 32, 0, len(xor), 0, 0, 0, 0)
    image = info + xor + bytes(mask)
    directory = struct.pack("<HHH", 0, 1, 1) + struct.pack(
        "<BBBBHHII", size, size, 0, 0, 1, 32, len(image), 22
    )
    path.write_bytes(directory + image)


def main() -> int:
    if sys.platform != "win32":
        print("请在 Windows 上运行这个打包脚本。")
        return 1
    try:
        import docx
        import fitz
        import pymupdf
        import lxml
    except ImportError as error:
        print(f"缺少程序依赖：{error}\n请先运行 python -m pip install -r requirements.txt")
        return 1

    base = Path(sys.base_prefix)
    if not (base / "pythonw.exe").is_file():
        print("找不到当前 Python 的 pythonw.exe，无法制作无控制台启动包。")
        return 1

    if PACKAGE.exists():
        shutil.rmtree(PACKAGE)
    RUNTIME.mkdir(parents=True)

    # Copy the interpreter and standard library, excluding optional test and
    # development modules. This avoids requiring Python on the recipient PC.
    for pattern in ("pythonw.exe", "python.exe", "python*.dll", "vcruntime*.dll", "python*.zip"):
        for source in base.glob(pattern):
            if source.is_file():
                shutil.copy2(source, RUNTIME / source.name)
    for name in ("DLLs", "tcl"):
        source = base / name
        if source.is_dir():
            copy_tree(source, RUNTIME / name)
    standard_library = base / "Lib"
    copy_tree(
        standard_library,
        RUNTIME / "Lib",
        {"site-packages", "test", "tests", "__pycache__", "idlelib", "ensurepip", "tkinter\test"},
    )

    site_packages = Path(get_paths()["purelib"])
    embedded_packages = RUNTIME / "Lib" / "site-packages"
    embedded_packages.mkdir(parents=True)
    package_dirs = {
        Path(docx.__file__).parent,
        Path(fitz.__file__).parent,
        Path(pymupdf.__file__).parent,
        Path(lxml.__file__).parent,
    }
    for source in package_dirs:
        copy_tree(source, embedded_packages / source.name, {"__pycache__", "tests", "test"})
    typing_file = site_packages / "typing_extensions.py"
    if typing_file.is_file():
        shutil.copy2(typing_file, embedded_packages / typing_file.name)

    shutil.copy2(ROOT / "main.py", PACKAGE / "main.py")
    icon = ROOT / "app.ico"
    create_icon(icon)
    shutil.copy2(icon, PACKAGE / icon.name)
    shutil.copy2(base / "LICENSE.txt", PACKAGE / "Python-LICENSE.txt")
    framework = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Microsoft.NET" / "Framework64" / "v4.0.30319" / "csc.exe"
    if not framework.is_file():
        framework = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Microsoft.NET" / "Framework" / "v4.0.30319" / "csc.exe"
    if not framework.is_file():
        print("找不到 Windows .NET Framework C# 编译器，无法生成 PDF转Word.exe 启动程序。")
        return 1
    subprocess.run(
        [str(framework), "/nologo", "/target:winexe", "/out:" + str(PACKAGE / "PDF转Word.exe"),
         "/win32icon:" + str(icon), "/reference:System.Windows.Forms.dll", "/reference:System.Drawing.dll",
         "/reference:System.Core.dll",
         str(ROOT / "launcher.cs")],
        check=True,
    )
    (PACKAGE / "使用说明.txt").write_text(
        "PDF 转 Word（Windows 免安装版）\n\n"
        "1. 解压整个压缩包，不要直接在压缩包里运行。\n"
        "2. 双击“PDF转Word.exe”。\n"
        "3. 把 PDF 拖进窗口，或点击“选择 PDF”。\n"
        "4. 选择输出文件夹；默认生成的 Word 保存在 PDF 所在文件夹。\n\n"
        "也可以把 PDF 文件拖到“PDF转Word.exe”上直接转换。\n"
        "程序已包含所需 Python 运行环境，不需要另装 Python 或依赖。\n"
        "外观保真模式会把每页存成图片，文字不能直接编辑。\n",
        encoding="utf-8",
    )

    if ARCHIVE.exists():
        ARCHIVE.unlink()
    subprocess.run(
        [sys.executable, "-c", "import shutil,sys; shutil.make_archive(sys.argv[1], 'zip', sys.argv[2], 'PDF转Word')", str(ARCHIVE.with_suffix("")), str(RELEASE_ROOT)],
        check=True,
    )
    print(f"已生成：{ARCHIVE}")
    print(f"压缩包大小：{ARCHIVE.stat().st_size / 1024 / 1024:.1f} MB")
    shutil.rmtree(RELEASE_ROOT)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
