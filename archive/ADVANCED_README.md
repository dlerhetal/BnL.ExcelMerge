# Advanced Excel Merger Application

## 1. Overview

This application allows you to merge an "expense" and a "revenue" Excel file, with the added ability to select which columns to include in the final output.

## 2. How to Use

1.  Run `AdvancedMerger.exe`.
2.  Click "Browse..." to select the **Expense File**.
3.  Click "Browse..." to select the **Revenue File**.
4.  Click the "Browse..." button next to **Output File Name** to choose a location and name for your merged file.
5.  Click the **Preview Columns** button.
6.  A new window will appear with a list of all columns from both files. By default, all columns are selected. Uncheck the columns you don't want in the final report.
7.  Click the **Combine** button.
8.  A success message will confirm that the file has been saved.

**Important:** The application merges the files based on the following key columns: `'Job ID'`, `'Cost Code ID'`, and `'Phase ID'`. These columns must exist in both files and will be included in the final output even if not selected.

## 3. How it Was Built

*   **Language:** Python
*   **GUI:** Tkinter
*   **Data Handling:** Pandas
*   **Packaging:** PyInstaller

The application logic is in `advanced_merger_app.py`.

## 4. Future Improvements

To modify the application, you'll need Python, `pandas`, and `pyinstaller`.

### 4.1. Modifying the Logic

The core logic is in the `AdvancedExcelMergerApp` class in `advanced_merger_app.py`.
*   The main window is set up in the `__init__` method.
*   The column selection window is created in the `preview_columns` method.
*   The merging and saving logic is in the `combine_and_save` method.

### 4.2. Rebuilding the Executable

After modifying `advanced_merger_app.py`, you can rebuild the executable with this command:

```bash
pyinstaller --onefile --windowed --name AdvancedMerger advanced_merger_app.py
```

This will update `AdvancedMerger.exe` in the `dist` folder.
