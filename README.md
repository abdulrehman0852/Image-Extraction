# 🖼️ Excel Image Extractor

This script extracts all embedded images from an Excel workbook (`.xlsx`) and renames them based on corresponding **tags** found in a specific column of the Excel sheet. It is useful for organizing large datasets where images are associated with rows (e.g., tagged products, users, or items).

---

## ✅ Features

- 🔍 Extracts images embedded in `.xlsx` files
- 🗂 Saves extracted images to a directory
- 🏷️ Renames images based on tags from **Column B**
- 🔁 Avoids filename conflicts by auto-incrementing duplicates
- 📁 Automatically creates the output directory

---

## 📦 Requirements

- Python 3.x  
- Libraries:
  - `openpyxl`
  - `nltk` *(installed but not used in this script)*
  - Standard Python modules: `zipfile`, `os`

Install the required third-party library:
```bash
pip install openpyxl
