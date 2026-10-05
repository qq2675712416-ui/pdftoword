# PDF 转 Word

一个 Windows 桌面工具，可把 PDF 转为可编辑的 Word 文档，也提供外观保真模式。支持拖入 PDF、选择输出目录和在转换过程中查看进度。

## 直接使用

从 GitHub 项目的 **Releases** 页面下载 `PDF转Word-免安装版.zip`，解压后双击 `PDF转Word.exe`。无需安装 Python 或打开命令行。

如暂时没有发布包，可在 Windows 上按下文从源码构建。

## 从源码运行

需要 Windows 和 Python。安装依赖后运行：

```powershell
python -m pip install -r requirements.txt
python start_app.py
```

也可以使用命令行转换：

```powershell
python main.py "输入.pdf" --output "输出.docx"
```

## 构建免安装版

在 Windows 的 Python 环境中安装项目依赖，再运行：

```powershell
python build_release.py
```

生成的 ZIP 位于项目根目录。分发时只需发送 ZIP，不要发送整个开发目录。更多说明见 [使用说明](使用说明.md) 和 [分发说明](分发说明.md)。

## 转换模式

- **可编辑文字**：尽量保留文字位置和格式；识别到的表格可作为 Word 表格编辑，标题会使用 Word 标题样式。复杂版式和扫描页可能需要手工整理；扫描 PDF 请先做 OCR。
- **外观保真（图片）**：将页面作为图片放入 Word，适合保留视觉外观，页面文字不能直接编辑。

## 隐私

转换在本机完成。输入 PDF 和输出 Word 不会由程序上传。
