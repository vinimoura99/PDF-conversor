Got it. Here is the full version in **English**, maintaining your exact layout, structure, and formatting style.

---

# 🤖 PDF Data Automation System (Pro Version)

This system was developed to convert standardized PDF documents into dynamic Excel reports, with support for cumulative history and automatic organization.

---

## 🚀 How to Use (For the End-User)

1. **Preparation**: Make sure the PDF files you want to process are saved on your computer.
2. **Execution**: Double-click the `Universal_Converter.exe` file.
3. **Selection**: A window will open. Select all the PDFs you want to convert (you can select multiple at once by holding the **CTRL** key).
4. **Result**: The program will create a folder named `Generated_Reports`. The Excel file `PROCESSING_REPORT.xlsx` will be inside with all data extracted, summed, and stylized.

---

## 🛠️ Developer's Guide (Creating the .exe)

To create the executable that the users will use, follow these steps in your terminal:

### 1. Install necessary libraries
```bash
pip install pdfplumber openpyxl pyinstaller pywin32
```

### 2. Generate the Executable
Run the following command to bundle the script into a single file without a console window:
```bash
pyinstaller --noconfirm --onefile --windowed --name "Universal_Converter" main.py
```

### 3. Requirements
* **Python 3.8+** installed.
* Standardized PDF input (digital text-based, not scanned images).

---

Would you like me to add a **"Common Errors"** section to this README to help users who might have the Excel file open during processing?