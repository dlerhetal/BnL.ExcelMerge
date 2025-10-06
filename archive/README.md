# Excel Merger Application

## 1. Overview

This application provides a simple graphical user interface (GUI) to merge two Excel files (`.xlsx` or `.xls`) into a single file. It is designed to combine an "expenses" file and a "revenue" file based on common columns.

## 2. How to Use

1.  Double-click the `merger_app.exe` file to run the application.
2.  Click the "Browse..." button for the "First File" to select your first Excel file.
3.  Click the "Browse..." button for the "Second File" to select your second Excel file.
4.  Once both files are selected, the "Combine and Save" button will be enabled. Click it.
5.  A "Save As" dialog will appear. Choose a name and location for your new merged Excel file and click "Save".
6.  A success message will appear once the file has been saved.

**Important:** The application identifies the "expense" and "revenue" files by looking for specific column headers.
*   The **expense file** must contain `'Est. Expenses'` and `'Act. Expenses'` columns.
*   The **revenue file** must contain `'Est. Revenue'` and `'Act. Revenue'` columns.

The two files are merged using the following common columns: `'Job ID'`, `'Cost Code ID'`, and `'Phase ID'`. These columns must exist in both files.

## 3. How it Was Built

The application is written in Python and uses the following libraries:

*   **Tkinter:** For the graphical user interface.
*   **Pandas:** For reading, merging, and writing Excel files.
*   **PyInstaller:** To package the application into a single executable file.

The main application logic is in the `merger_app.py` file.

## 4. Future Improvements

To modify the application, you will need to have Python installed on your system, along with the `pandas` and `pyinstaller` libraries. You can install them using pip:
```bash
pip install pandas pyinstaller
```

### 4.1. Modifying the Merging Logic

The core merging logic is in the `combine_files` method within the `ExcelMergerApp` class in `merger_app.py`. You can edit this function to change:
*   The key columns used for merging.
*   The logic for identifying the expense and revenue files.
*   The way data is cleaned or processed before merging.

### 4.2. Rebuilding the Executable

After you have modified the `merger_app.py` script, you can rebuild the executable by running the following command in your terminal from the project's root directory:

```bash
pyinstaller --onefile --windowed merger_app.py
```

This will create an updated `merger_app.exe` in the `dist` folder.
