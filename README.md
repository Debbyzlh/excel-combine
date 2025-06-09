# excel-combine
extract key-value pairs from child files uploaded, check keys within parent file, output warning if needed, then output a excel file ready for vlookup with a suggested excel formula.

# 🔗 Excel Parent-Child Merger App

This Streamlit app lets you merge values from one or more child Excel files into a parent Excel file based on matching key-value pairs.

## 🚀 Features

- Upload a parent Excel file and preview its sheets
- Upload one or multiple child Excel files
- Select key and value columns using Excel-style letters (A, B, C...)
- Automatically extract and preview key-value pairs from child files
- Fill missing values in the parent file where keys match
- Warn on:
  - Duplicate keys with different values
  - New keys not present in the parent file
- Download the extracted key-value pairs as a new Excel file
- VLOOKUP formula suggestion for Excel-based merging

## 🧠 How It Works

1. Upload the parent and child files
2. Select sheets and corresponding key/value columns
3. App extracts key-value pairs from the child files
4. Matches keys to the parent and fills empty values
5. Displays summary, conflict warnings, and allows download

## 📥 Installation

```bash
pip install -r requirements.txt
