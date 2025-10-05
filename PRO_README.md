# Pro Excel Merger Application

## 1. Overview

This professional version of the Excel Merger application allows you to merge an "expense" and a "revenue" Excel file, with advanced control over the output. You can rename, reorder, and remove columns before merging.

## 2. How to Use

1.  Run `ProMerger.exe`.
2.  Select the **Expense File**, **Revenue File**, and the desired **Output File** path using the "Browse..." buttons.
3.  Click the **Configure Columns** button.
4.  A new window will appear with a list of all columns from both files.
5.  In this window, you can:
    *   **Reorder:** Select a column and use the "Up" and "Down" buttons to change its position.
    *   **Rename:** Select a column and click "Rename" to enter a new name.
    *   **Remove:** Select a column and click "Remove" to exclude it from the final report.
6.  Once you are satisfied with your column configuration, click the **Combine and Save** button.
7.  A success message will confirm that the file has been saved.

**Important:** The application merges the files based on the following key columns: `'Job ID'`, `'Cost Code ID'`, and `'Phase ID'`. These columns must exist in both files.

## 3. How it Was Built

*   **Language:** Python
*   **GUI:** Tkinter
*   **Data Handling:** Pandas
*   **Packaging:** PyInstaller

The application logic is in `pro_merger_app.py`.

## 4. Future Improvements

To modify the application, you'll need Python, `pandas`, and `pyinstaller`.

### 4.1. Modifying the Logic

The core logic is in the `ProExcelMergerApp` class in `pro_merger_app.py`.
*   The main window is set up in the `__init__` method.
*   The column configuration window and its logic are in the `open_column_config` method and the associated `move_up`, `move_down`, `rename_column`, and `remove_column` methods.
*   The final merging and saving logic is in the `combine_and_save` method.

### 4.2. Rebuilding the Executable

After modifying `pro_merger_app.py`, you can rebuild the executable with this command:

```bash
pyinstaller --onefile --windowed --name ProMerger pro_merger_app.py
```

This will update `ProMerger.exe` in the `dist` folder.
