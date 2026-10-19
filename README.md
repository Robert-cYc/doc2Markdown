# Document to Markdown Pro (doc2Markdown)

A powerful, all-in-one document-to-Markdown converter with integrated OCR, image extraction, Desktop GUI (drag-and-drop), Web UI, and Windows standalone executable packaging.

---

## 🌟 Key Features

1. **Multi-Format Document Conversion**:
   - **PDF (`.pdf`)**: Extracts structured text, tabular data, embedded images, and automatically detects/OCRs scanned pages.
   - **Excel (`.xlsx`, `.xls`, `.csv`)**: Converts all worksheets into clean, standardized Markdown pipe tables.
   - **Word (`.docx`, `.doc`)**: Preserves heading hierarchy (`H1`–`H6`), bold/italic formatting, lists, tables, and extracts embedded figures.
   - **PowerPoint (`.pptx`, `.ppt`)**: Extracts slide titles, bullet points, speaker notes, tables, and presentation graphics.
   - **Plain Text & Media (`.txt`, `.json`, `.png`, `.jpg`, `.tiff`, `.webp`)**: Auto-detects text encodings (UTF-8, CP950/Big5, GB18030, etc.) or performs OCR recognition.

2. **Built-in OCR Engines**:
   - **RapidOCR (`rapidocr_onnxruntime`)**: Completely offline ONNX models for Chinese and English. Zero external dependencies needed.
   - **Tesseract OCR (`pytesseract`)**: Automatically locates portable or installed `Tesseract-OCR` on Windows (supports `chi_tra+eng`).
   - **EasyOCR (`easyocr`)**: Supported as an automated fallback.

3. **Automated Asset & Image Extraction**:
   - Saves embedded images to `<filename>_assets/`.
   - Links exported images directly in the Markdown text (`![Image Description](<filename>_assets/xxx.png)`).

4. **Multiple Operating Modes**:
   - **Desktop GUI**: Native drag-and-drop file window (`tkinterdnd2` / `tkinter`).
   - **Web UI**: In-browser drag-and-drop converter with live preview, markdown copy, and ZIP download (`http://127.0.0.1:8000`).
   - **CLI / Batch Mode**: Terminal commands or direct drag-and-drop onto the `.exe` icon.
   - **PyInstaller Bundler**: One-click build script to generate standalone Windows executables (`--onedir` or `--onefile`).

---

## 🚀 Installation & Requirements

### Prerequisites
- Python 3.8 ~ 3.14+ (Windows / macOS / Linux)

### Install Dependencies
```bash
pip install --upgrade pip
pip install pdfplumber pypdf pypdfium2 pillow openpyxl pandas python-docx python-pptx tkinterdnd2 pytesseract rapidocr_onnxruntime markitdown
```

---

## 🖥️ Usage

### 1. Desktop GUI Mode
Launch the graphical interface (default when running with no arguments or double-clicking the compiled `.exe`):
```bash
python convert_to_markdown.py
# or
python convert_to_markdown.py --gui
```
- Drag and drop files directly into the window.
- Customize OCR language, image extraction options, and output destination.
- Real-time progress bar and markdown preview.

### 2. Web UI Mode
Launch the local web server:
```bash
python convert_to_markdown.py --web --port 8000
```
Open [http://127.0.0.1:8000](http://127.0.0.1:8000) to upload, preview Markdown, and download results or bundled ZIP packages.

### 3. Command Line Interface (CLI)

#### Convert Single File
```bash
python convert_to_markdown.py sample.pdf -o output/
```

#### Batch Convert Entire Directory (Recursive)
```bash
python convert_to_markdown.py "C:\path\to\documents" -r -o "C:\path\to\markdown_out"
```

#### Disable Image Extraction
```bash
python convert_to_markdown.py report.docx --no-images
```

#### Disable OCR for Scanned Pages
```bash
python convert_to_markdown.py scan.pdf --no-ocr
```

#### Specify OCR Language
```bash
python convert_to_markdown.py scanned.png --ocr-lang "chi_tra+eng"
```

---

## 📦 Packaging Standalone Executables (PyInstaller)

### Option A: One-Folder Mode (Recommended)
Fastest startup time (loads instantly without uncompressing to `%TEMP%`):
```bash
python convert_to_markdown.py --build-exe --onedir
```
Output directory: `dist/DocumentToMarkdownPro/`

### Option B: Single File Mode (`--onefile`)
Bundles everything into a single `.exe`:
```bash
python convert_to_markdown.py --build-exe --onefile
```
Output file: `dist/DocumentToMarkdownPro.exe`

### Option C: Windows Batch Script
Run or double-click:
```cmd
build_windows_exe.bat
```

### Option D: Native PyInstaller Command
```powershell
python -m PyInstaller --onedir --windowed --name DocumentToMarkdownPro `
  --collect-all tkinterdnd2 `
  --collect-all pypdfium2 `
  --collect-all pypdfium2_raw `
  --collect-all rapidocr_onnxruntime `
  --collect-all pdfplumber `
  --collect-all pdfminer `
  --collect-all markitdown `
  --hidden-import openpyxl `
  --hidden-import docx `
  --hidden-import pptx `
  --hidden-import pypdf `
  --hidden-import PIL `
  --hidden-import pytesseract `
  --hidden-import pandas `
  --distpath dist --workpath $env:TEMP/pyi_build `
  convert_to_markdown.py
```

---

## ⚙️ CLI Reference

| Parameter | Type | Default | Description |
| :--- | :--- | :--- | :--- |
| `inputs` | `path ...` | `None` | Files or directories to convert (launches GUI if omitted). |
| `-o`, `--output` | `path` | Same directory | Output `.md` file or destination directory. |
| `-r`, `--recursive` | flag | `False` | Recursively process subdirectories when input is a folder. |
| `--no-images` | flag | `False` | Disable embedded image extraction. |
| `--no-ocr` | flag | `False` | Disable OCR for scanned PDFs and image files. |
| `--ocr-images` | flag | `False` | Run OCR on embedded figures in Word/PPT/PDF. |
| `--ocr-lang` | string | `chi_tra+eng` | Tesseract OCR language pack. |
| `--engine` | choice | `custom` | Engine selection: `custom`, `auto`, `markitdown`. |
| `--gui` | flag | `False` | Force launch Desktop GUI window. |
| `--web` | flag | `False` | Launch local Web UI server. |
| `--host` | string | `127.0.0.1` | Web server host address. |
| `--port` | int | `8000` | Web server listening port. |
| `--build-exe` | flag | `False` | Trigger PyInstaller packaging. |
| `--onedir`, `--onefolder` | flag | `False` | Package as a directory bundle (fastest startup). |
| `--onefile` | flag | `False` | Package as a single executable file. |
| `--console` | flag | `False` | Enable console terminal window in compiled `.exe`. |
| `--exe-name` | string | `DocumentToMarkdownPro`| Output executable name. |
| `--bundle-tesseract` | path | `None` | Bundle portable Tesseract-OCR directory. |
| `--generate-bat` | flag | `False` | Generate `build_windows_exe.bat` in current directory. |

---

## 📄 License
MIT License.
