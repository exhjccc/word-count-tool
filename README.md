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
- Required libraries:
    - `tkinter`
    - `pandas`
    - `numpy`
    - `pyreadstat`
    - `PyPDF2`
    - `python-docx`
    - `openpyxl`
    - `markdown`
    - `lxml`
    - `pyyaml`
    - `rpy2`
    - `zipfile`
    - `pytube` (for extracting audio transcription)

## How to Use
1. Install the required libraries by running:
    ```bash
    pip install pandas numpy pyreadstat PyPDF2 python-docx openpyxl markdown lxml pyyaml rpy2 pytube
    ```
2. Run the program:
    ```bash
    python universal_word_counter.py
    ```
3. Use the "Open File" button in the GUI to select a file.
4. View the word and character counts displayed in the application.

---

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
- 必要的库：
    - `tkinter`
    - `pandas`
    - `numpy`
    - `pyreadstat`
    - `PyPDF2`
    - `python-docx`
    - `openpyxl`
    - `markdown`
    - `lxml`
    - `pyyaml`
    - `rpy2`
    - `zipfile`
    - `pytube`

## 使用方法
1. 安装必要的库：
    ```bash
    pip install pandas numpy pyreadstat PyPDF2 python-docx openpyxl markdown lxml pyyaml rpy2 pytube
    ```
2. 运行程序：
    ```bash
    python universal_word_counter.py
    ```
3. 点击 GUI 界面中的 "打开文件" 按钮选择文件。
4. 查看文件字数和字符数，并显示于界面中。

