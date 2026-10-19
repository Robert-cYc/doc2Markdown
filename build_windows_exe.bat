@echo off
chcp 65001 >nul
echo ====================================================================
echo   多格式文件轉 Markdown 工具 - Windows 打包腳本 (One-Folder / One-Dir)
echo ====================================================================
echo.

echo [1/3] 正在檢查並安裝必要 Python 套件與 PyInstaller...
pip install --upgrade pip
pip install pyinstaller pdfplumber pypdf pypdfium2 pillow openpyxl pandas python-docx python-pptx tkinterdnd2 pytesseract rapidocr_onnxruntime markitdown

echo.
echo [2/3] 正在透過 PyInstaller 封裝資料夾執行檔 (dist\DocumentToMarkdownPro\)...
python "%~dp0convert_to_markdown.py" --build-exe --onedir

echo.
echo [3/3] 打包完成！請至 dist\DocumentToMarkdownPro\DocumentToMarkdownPro.exe 啟動程式。
pause
