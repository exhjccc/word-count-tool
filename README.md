# 全能字数统计器

## 概述
全能字数统计器是一款功能强大的软件，能对多种格式的文件进行分析和字数统计。该软件用 Python 开发，接口亦配处理 Tkinter 带来的图形用户界面(GUI)。

## 功能特点
- **支持多种文件格式**:
    - 文本文件 (`.txt`)
    - Word 文档 (`.docx`)
    - PDF 文件 (`.pdf`)
    - HTML 文件 (`.html`)
    - JSON 文件 (`.json`)
    - Jupyter Notebook 文件 (`.ipynb`)
    - Excel 文件 (`.xlsx`)
    - Markdown 文件 (`.md`)
    - CSV 文件 (`.csv`)
    - RTF 文件 (`.rtf`)
    - LaTeX 文件 (`.tex`)
    - PowerPoint 文件 (`.pptx`)
    - EPUB 文件 (`.epub`)
    - XML 文件 (`.xml`)
    - YAML 文件 (`.yaml` / `.yml`)
    - 日志文件 (`.log`)
    - 压缩文件 (`.zip`)
    - 音频文件试词 (`.mp3`, `.wav`)
    - R 文件 (`.R`, `.RData`, `.rds`)
    - Stata 文件 (`.dta`)

- **文件读取**: 允许用户通过 GUI 选择并读入文件。
- **字数和字符数统计**: 文件读入后自动计算。
- **错误处理**: 对不支持的格式或处理问题提供有效的错误信息。

## 系统需求
- Python 3.8 或更高版本
- 所需 Python 库：
  - `tkinter`
  - `pandas`
  - `docx` (`python-docx`)
  - `pdfplumber`
  - `openpyxl`
  - `bs4` (`beautifulsoup4`)
  - `striprtf`
  - `markdown`
  - `pyyaml`
  - `python-pptx`
  - `speechrecognition`
  - `pyreadstat`（用于 Stata 文件）
  - `rpy2`（可选，用于 R 文件）

## 使用方法
1. 安装必要的库：
   ```bash
   pip install pandas python-docx pdfplumber openpyxl beautifulsoup4 striprtf markdown pyyaml python-pptx SpeechRecognition pyreadstat
   ```
    若需要支持 R 文件格式，安装 rpy2
    ```bash
    pip install rpy2
    ```
2. 运行程序：
    ```bash
    python wordcount.py
    ```
3. 点击 GUI 界面中的 "打开文件" 按钮选择文件。
4. 查看文件字数和字符数，并显示于界面中。

---

# Universal Word Counter

## Overview
The Universal Word Counter is a comprehensive tool designed to analyze and count words from various file formats. This versatile application is developed using Python and features a graphical user interface (GUI) created with Tkinter.

## Features
- **Support for Multiple File Formats**:
    - Text files (`.txt`)
    - Word documents (`.docx`)
    - PDF files (`.pdf`)
    - HTML files (`.html`)
    - JSON files (`.json`)
    - Jupyter Notebook files (`.ipynb`)
    - Excel files (`.xlsx`)
    - Markdown files (`.md`)
    - CSV files (`.csv`)
    - RTF files (`.rtf`)
    - LaTeX files (`.tex`)
    - PowerPoint files (`.pptx`)
    - EPUB files (`.epub`)
    - XML files (`.xml`)
    - YAML files (`.yaml` / `.yml`)
    - Log files (`.log`)
    - Archive files (`.zip`)
    - Audio transcription files (`.mp3`, `.wav`)
    - R files (`.R`, `.RData`, `.rds`)
    - Stata files (`.dta`)

- **File Loading**: Allows users to browse and load files of supported formats directly through the GUI.
- **Word and Character Count**: Automatically calculates the word and character counts upon loading the file.
- **Error Handling**: Provides meaningful error messages for unsupported formats or other processing issues.

## System Requirements
- Python 3.8 or higher

以下是英文版本的 README 文件，包含所有最新的功能描述和依赖项：

markdown
复制代码
# Universal Word Count Program

## Features
This program is a versatile word count tool built with Python and Tkinter, supporting multiple file formats including but not limited to:
- Text files (`*.txt`)
- Word documents (`*.docx`)
- PDF files (`*.pdf`)
- HTML files (`*.html`)
- JSON files (`*.json`)
- Jupyter Notebook files (`*.ipynb`)
- Excel spreadsheets (`*.xlsx`)
- Markdown files (`*.md`)
- R files (`*.R`, `*.RData`, `*.rds`) [Optional: Requires `rpy2`]
- Stata data files (`*.dta`)
- E-book files (`*.epub`)
- Compressed files (`*.zip`, with recursive extraction)
- Other common formats (`.csv`, `.rtf`, `.tex`, `.pptx`, etc.)

## Highlights
1. **Multi-format Support:** Handles word counting across various file formats in one tool.
2. **User-friendly Interface:** Built with Tkinter, featuring a simple GUI.
3. **Intelligent Parsing:** Utilizes appropriate libraries for each file type.
4. **Recursive Extraction:** Can parse ZIP archives and EPUB files for content.
5. **Audio Transcription:** Extracts text from MP3/WAV files via speech recognition for word counting.
6. **Extensibility:** Easily extendable to support additional file formats.

## How to Use
1. Launch the program and click the **Load File** button.
2. In the file selection dialog, choose the target file.
3. The program will automatically parse the file and count the words.
4. The results will be displayed in the main interface, along with a preview of the content.

## Dependencies
- Python 3.8+
- Required Python libraries:
  - `tkinter`
  - `pandas`
  - `docx` (`python-docx`)
  - `pdfplumber`
  - `openpyxl`
  - `bs4` (`beautifulsoup4`)
  - `striprtf`
  - `markdown`
  - `pyyaml`
  - `python-pptx`
  - `speechrecognition`
  - `pyreadstat` (for Stata files)
  - `rpy2` (optional, for R files)

## How to Use
1. Install the required libraries by running:
    ```bash
    pip install pandas python-docx pdfplumber openpyxl beautifulsoup4 striprtf markdown pyyaml python-pptx SpeechRecognition pyreadstat
    ```
    If R file support is needed, install rpy2
    ```bash
    pip install rpy2
    ```
2. Run the program:
    ```bash
    python wordcount.py
    ```
3. Use the "Open File" button in the GUI to select a file.
4. View the word and character counts displayed in the application.



