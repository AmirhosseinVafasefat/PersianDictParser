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
```bash
python parser.py [docx file] [xlsx template] [sheet name] [output xlsx] [entry column] [meaning column] [synonym column] [examples column]
```

## Installation

Clone the repository and install dependencies:

```bash
git clone https://github.com/AmirhosseinVafasefat/PersianDictParser.git
cd PersianDictParser
pip install -r requirements.txt
