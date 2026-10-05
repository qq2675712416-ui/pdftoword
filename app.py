"""Desktop interface for converting PDF files to editable Word documents."""

from __future__ import annotations

import ctypes
import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

from main import pdf_to_word, pdf_to_word_visual


class PdfWordApp:
    BG = "#F3F6FB"
    INK = "#182230"
    MUTED = "#667085"
    BLUE = "#2563EB"

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("PDF 转 Word")
        self.root.geometry("760x690")
        self.root.minsize(680, 630)
        self.root.configure(bg=self.BG)
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.busy = False
        self._drop_callbacks = []
        self._old_wndprocs = []
        self.output_folder = ""
        self.output_dir = tk.StringVar(value="与 PDF 同文件夹（默认）")
        self.mode = tk.StringVar(value="可编辑文字")
        self.dpi = tk.StringVar(value="300")
        self.status = tk.StringVar(value="就绪 · 选择 PDF 或拖放文件开始转换")

        icon = Path(__file__).with_name("app.ico")
        if icon.is_file():
            try:
                self.root.iconbitmap(str(icon))
            except tk.TclError:
                pass

        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("TCombobox", padding=(8, 7), font=("Microsoft YaHei UI", 10))
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(10, 7))
        style.configure("Horizontal.TProgressbar", troughcolor="#E6ECF4", background=self.BLUE)

        self._build_ui()
        self.root.after(100, self._check_events)
        self.root.after(300, self._enable_native_drop)

    def _build_ui(self) -> None:
        page = tk.Frame(self.root, bg=self.BG)
        page.pack(fill="both", expand=True, padx=36, pady=(26, 22))

        header = tk.Frame(page, bg=self.BG)
        header.pack(fill="x", pady=(0, 20))
        tk.Label(
            header, text="PDF 转 Word", bg=self.BG, fg=self.INK,
            font=("Microsoft YaHei UI", 24, "bold"),
        ).pack(anchor="w")
        tk.Label(
            header, text="轻松转换文档，文字和表格尽量保持可编辑", bg=self.BG, fg=self.MUTED,
            font=("Microsoft YaHei UI", 10),
        ).pack(anchor="w", pady=(3, 0))

        self.drop_area = tk.Frame(
            page, bg="white", highlightbackground="#C8D6EA", highlightthickness=1,
        )
        self.drop_area.pack(fill="x", ipady=22)
        self.drop_area.bind("<Button-1>", lambda _event: self.choose_files())
        self.drop_title = tk.Label(
            self.drop_area, text="拖放 PDF 文件到这里", bg="white", fg=self.INK,
            font=("Microsoft YaHei UI", 16, "bold"), cursor="hand2",
        )
        self.drop_title.pack(pady=(10, 4))
        self.drop_title.bind("<Button-1>", lambda _event: self.choose_files())
        self.drop_hint = tk.Label(
            self.drop_area, text="支持一次添加多个文件 · 也可以点击下方按钮浏览", bg="white", fg=self.MUTED,
            font=("Microsoft YaHei UI", 9), cursor="hand2",
        )
        self.drop_hint.pack(pady=(0, 13))
        self.drop_hint.bind("<Button-1>", lambda _event: self.choose_files())
        self.choose_button = tk.Button(
            self.drop_area, text="选择 PDF 文件", command=self.choose_files,
            bg=self.BLUE, fg="white", activebackground="#1D4ED8", activeforeground="white",
            relief="flat", borderwidth=0, padx=22, pady=10,
            font=("Microsoft YaHei UI", 10, "bold"), cursor="hand2",
        )
        self.choose_button.pack(pady=(0, 6))

        settings = tk.Frame(page, bg="white", highlightbackground="#E0E6EF", highlightthickness=1)
        settings.pack(fill="x", pady=(18, 0))
        tk.Label(settings, text="转换设置", bg="white", fg=self.INK,
                 font=("Microsoft YaHei UI", 12, "bold")).pack(anchor="w", padx=20, pady=(17, 12))

        mode_row = tk.Frame(settings, bg="white")
        mode_row.pack(fill="x", padx=20)
        tk.Label(mode_row, text="转换模式", width=12, anchor="w", bg="white", fg="#344054",
                 font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.mode_box = ttk.Combobox(
            mode_row, textvariable=self.mode, values=("可编辑文字", "外观保真（图片）"),
            state="readonly", width=28,
        )
        self.mode_box.pack(side="left", fill="x", expand=True)
        self.mode_box.bind("<<ComboboxSelected>>", self._mode_changed)
        self.dpi_label = tk.Label(mode_row, text="图片清晰度", bg="white", fg=self.MUTED,
                                  font=("Microsoft YaHei UI", 9))
        self.dpi_label.pack(side="left", padx=(14, 6))
        self.dpi_box = ttk.Combobox(mode_row, textvariable=self.dpi, values=("300", "600"),
                                    state="disabled", width=5)
        self.dpi_box.pack(side="left")
        tk.Label(mode_row, text="DPI", bg="white", fg=self.MUTED,
                 font=("Microsoft YaHei UI", 9)).pack(side="left", padx=(5, 0))

        output_row = tk.Frame(settings, bg="white")
        output_row.pack(fill="x", padx=20, pady=(16, 6))
        tk.Label(output_row, text="输出位置", width=12, anchor="w", bg="white", fg="#344054",
                 font=("Microsoft YaHei UI", 10)).pack(side="left")
        self.output_entry = ttk.Entry(output_row, textvariable=self.output_dir, state="readonly")
        self.output_entry.pack(side="left", fill="x", expand=True)
        self.browse_output_button = ttk.Button(output_row, text="选择文件夹…", command=self.choose_output_folder)
        self.browse_output_button.pack(side="left", padx=(9, 0))
        self.reset_output_button = ttk.Button(output_row, text="恢复默认", command=self.reset_output_folder)
        self.reset_output_button.pack(side="left", padx=(7, 0))
        self.output_hint = tk.Label(
            settings, text="留空时，每个 Word 文件会保存到对应 PDF 所在文件夹。",
            bg="white", fg=self.MUTED, font=("Microsoft YaHei UI", 9),
        )
        self.output_hint.pack(anchor="w", padx=(116, 20), pady=(0, 15))

        progress_card = tk.Frame(page, bg=self.BG)
        progress_card.pack(fill="x", pady=(18, 0))
        self.progress = ttk.Progressbar(progress_card, mode="determinate", maximum=1, value=0)
        self.progress.pack(fill="x")
        tk.Label(progress_card, textvariable=self.status, bg=self.BG, fg=self.MUTED,
                 font=("Microsoft YaHei UI", 9), anchor="w").pack(fill="x", pady=(9, 0))

        tk.Label(
            page, text="可编辑模式适合继续修改；复杂版式和扫描件可能需要后续整理。",
            bg=self.BG, fg="#98A2B3", font=("Microsoft YaHei UI", 9),
        ).pack(anchor="w", side="bottom", pady=(18, 0))

    def _mode_changed(self, _event: object = None) -> None:
        state = "readonly" if self.mode.get().startswith("外观保真") else "disabled"
        self.dpi_box.configure(state=state)
        self.dpi_label.configure(fg="#344054" if state == "readonly" else "#B0B8C4")

    def choose_output_folder(self) -> None:
        initial = self.output_folder or str(Path.home())
        folder = filedialog.askdirectory(title="选择 Word 文件保存位置", initialdir=initial, mustexist=True)
        if folder:
            self.output_folder = folder
            self.output_dir.set(folder)
            self.status.set("Word 文件将保存到所选文件夹。")

    def reset_output_folder(self) -> None:
        self.output_folder = ""
        self.output_dir.set("与 PDF 同文件夹（默认）")
        self.status.set("已恢复默认：每个 Word 文件保存到对应 PDF 所在文件夹。")

    def _enable_native_drop(self) -> None:
        """Register the top-level window and its child widgets for Explorer drops."""
        if os.name != "nt":
            return
        try:
            from ctypes import wintypes

            hwnd = self.root.winfo_id()
            shell32, user32 = ctypes.windll.shell32, ctypes.windll.user32
            shell32.DragAcceptFiles.argtypes = [wintypes.HWND, wintypes.BOOL]
            shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
            shell32.DragQueryFileW.restype = wintypes.UINT
            shell32.DragFinish.argtypes = [wintypes.HANDLE]
            self._wndproc_type = ctypes.WINFUNCTYPE(
                ctypes.c_ssize_t, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM
            )
            get_long = getattr(user32, "GetWindowLongPtrW", user32.GetWindowLongW)
            set_long = getattr(user32, "SetWindowLongPtrW", user32.SetWindowLongW)
            get_long.argtypes = [wintypes.HWND, ctypes.c_int]
            get_long.restype = ctypes.c_ssize_t
            set_long.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_ssize_t]
            set_long.restype = ctypes.c_ssize_t
            user32.CallWindowProcW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT,
                                               wintypes.WPARAM, wintypes.LPARAM]
            user32.CallWindowProcW.restype = ctypes.c_ssize_t
            user32.GetAncestor.argtypes = [wintypes.HWND, wintypes.UINT]
            user32.GetAncestor.restype = wintypes.HWND

            targets = {int(hwnd), int(user32.GetAncestor(hwnd, 2))}
            enum_type = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

            def collect_child(child, _data):
                targets.add(int(child))
                return True

            enum_callback = enum_type(collect_child)
            user32.EnumChildWindows.argtypes = [wintypes.HWND, enum_type, wintypes.LPARAM]
            user32.EnumChildWindows(hwnd, enum_callback, 0)

            for target in targets:
                if not target:
                    continue
                shell32.DragAcceptFiles(target, True)
                old_proc = get_long(target, -4)

                def wndproc(window, message, wparam, lparam, previous=old_proc):
                    if message == 0x0233:
                        count = shell32.DragQueryFileW(wparam, 0xFFFFFFFF, None, 0)
                        paths = []
                        for index in range(count):
                            length = shell32.DragQueryFileW(wparam, index, None, 0) + 1
                            buffer = ctypes.create_unicode_buffer(length)
                            shell32.DragQueryFileW(wparam, index, buffer, length)
                            paths.append(buffer.value)
                        shell32.DragFinish(wparam)
                        self.root.after(0, self.handle_files, paths)
                        return 0
                    return user32.CallWindowProcW(previous, window, message, wparam, lparam)

                callback = self._wndproc_type(wndproc)
                self._old_wndprocs.append(old_proc)
                self._drop_callbacks.append(callback)
                set_long(target, -4, ctypes.cast(callback, ctypes.c_void_p).value)
        except (AttributeError, OSError) as error:
            self.status.set(f"拖放初始化失败，可点击“选择 PDF 文件”导入。({error})")

    def choose_files(self) -> None:
        if self.busy:
            return
        paths = filedialog.askopenfilenames(
            title="选择 PDF 文件", filetypes=(("PDF 文件", "*.pdf"), ("所有文件", "*.*"))
        )
        if paths:
            self.handle_files(list(paths))

    def handle_files(self, paths: list[str]) -> None:
        pdfs = [Path(path) for path in paths if Path(path).suffix.lower() == ".pdf"]
        if not pdfs:
            messagebox.showinfo("请选择 PDF", "没有找到 PDF 文件，请重新选择。", parent=self.root)
            return
        if self.busy:
            messagebox.showinfo("正在转换", "当前转换完成后再添加其他文件。", parent=self.root)
            return
        self.busy = True
        self.choose_button.configure(state="disabled", text="正在转换…")
        self.progress.configure(maximum=len(pdfs), value=0)
        self.status.set(f"准备转换 {len(pdfs)} 个文件…")
        visual = self.mode.get().startswith("外观保真")
        threading.Thread(
            target=self._convert_files,
            args=(pdfs, int(self.dpi.get()), visual, self.output_folder),
            daemon=True,
        ).start()

    @staticmethod
    def _unique_destination(path: Path) -> Path:
        if not path.exists():
            return path
        suffix = 2
        while True:
            candidate = path.with_name(f"{path.stem}({suffix}){path.suffix}")
            if not candidate.exists():
                return candidate
            suffix += 1

    def _convert_files(self, pdfs: list[Path], dpi: int, visual: bool, output_dir: str) -> None:
        completed: list[Path] = []
        failures: list[str] = []
        for index, source in enumerate(pdfs, start=1):
            self.events.put(("status", f"正在转换 {index}/{len(pdfs)}：{source.name}"))
            label = "保真版" if visual else "可编辑版"
            folder = Path(output_dir) if output_dir else source.parent
            destination = self._unique_destination(folder / f"{source.stem}-{label}.docx")
            try:
                if visual:
                    pdf_to_word_visual(source, destination, dpi=dpi)
                else:
                    pdf_to_word(source, destination)
                completed.append(destination)
            except Exception as error:
                failures.append(f"{source.name}：{error}")
            self.events.put(("progress", index))
        self.events.put(("done", (completed, failures)))

    def _check_events(self) -> None:
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == "status":
                    self.status.set(data)
                elif kind == "progress":
                    self.progress.configure(value=data)
                else:
                    completed, failures = data
                    self.busy = False
                    self.choose_button.configure(state="normal", text="选择 PDF 文件")
                    if completed and not failures:
                        self.status.set(f"转换完成 · 共生成 {len(completed)} 个 Word 文件")
                        messagebox.showinfo("转换完成", "已生成文件：\n\n" + "\n".join(map(str, completed)), parent=self.root)
                    elif completed:
                        self.status.set(f"部分完成 · 成功 {len(completed)} 个，失败 {len(failures)} 个")
                        messagebox.showwarning(
                            "部分完成", "已生成：\n" + "\n".join(map(str, completed))
                            + "\n\n未能转换：\n" + "\n".join(failures), parent=self.root,
                        )
                    else:
                        self.status.set("转换失败，请检查 PDF 文件后重试。")
                        messagebox.showerror("转换失败", "\n".join(failures) or "请检查 PDF 文件。", parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self._check_events)

    def run(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.mainloop()


if __name__ == "__main__":
    app = PdfWordApp()
    if len(sys.argv) > 1:
        app.root.after(400, app.handle_files, sys.argv[1:])
    app.run()
