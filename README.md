# PersianDictParser

PersianDictParser is a Python tool for parsing **pre-processed Persian dictionary data**
from `.docx` files and exporting the extracted entries into structured `.xlsx` files.

The script reads dictionary entries from a Word document and writes them into
user-specified columns in an Excel sheet, preserving a customizable order.

## Overview

This project is designed for linguistic and NLP-related workflows where dictionary
data is available in DOCX format but needs to be converted into a machine-
readable spreadsheet for further processing or analysis.

## Usage
python parser.py <docx_file> <xlsx_template> <sheet_name> <output_xlsx> <entry_column> <meaning_column> <synonym_column> <examples_column>

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/AmirhosseinVafasefat/PersianDictParser.git
cd PersianDictParser
pip install -r requirements.txt
