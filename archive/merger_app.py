
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel File Merger")

        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()

        # Frame for file selection
        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

        # File 1 selection
        tk.Label(frame, text="First File:").grid(row=0, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.file1_path, width=50, state='readonly').grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_file1).grid(row=0, column=2, padx=5, pady=5)

        # File 2 selection
        tk.Label(frame, text="Second File:").grid(row=1, column=0, padx=5, pady=5)
        tk.Entry(frame, textvariable=self.file2_path, width=50, state='readonly').grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_file2).grid(row=1, column=2, padx=5, pady=5)

        # Combine button
        self.combine_button = tk.Button(root, text="Combine and Save", command=self.combine_files, state=tk.DISABLED)
        self.combine_button.pack(pady=10)

    def browse_file1(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file1_path.set(filename)
            self.check_files_selected()

    def browse_file2(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.file2_path.set(filename)
            self.check_files_selected()

    def check_files_selected(self):
        if self.file1_path.get() and self.file2_path.get():
            self.combine_button.config(state=tk.NORMAL)

    def combine_files(self):
        file1 = self.file1_path.get()
        file2 = self.file2_path.get()

        try:
            # Read the Excel files
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)

            # Clean column names
            df1.columns = df1.columns.str.strip()
            df2.columns = df2.columns.str.strip()

            # Identify which file is which based on columns
            expense_cols = ['Est. Expenses', 'Act. Expenses']
            revenue_cols = ['Est. Revenue', 'Act. Revenue']

            if all(col in df1.columns for col in expense_cols):
                expenses_df = df1
                revenue_df = df2
            elif all(col in df2.columns for col in expense_cols):
                expenses_df = df2
                revenue_df = df1
            else:
                messagebox.showerror("Error", "Could not identify the expense file. Make sure it has 'Est. Expenses' and 'Act. Expenses' columns.")
                return

            if not all(col in revenue_df.columns for col in revenue_cols):
                messagebox.showerror("Error", "Could not identify the revenue file. Make sure it has 'Est. Revenue' and 'Act. Revenue' columns.")
                return

            # Define key columns for merging
            key_cols = ['Job ID', 'Cost Code ID', 'Phase ID']

            # Check for key columns
            if not all(col in expenses_df.columns for col in key_cols):
                messagebox.showerror("Error", f"Expense file is missing one of the key columns: {key_cols}")
                return
            if not all(col in revenue_df.columns for col in key_cols):
                messagebox.showerror("Error", f"Revenue file is missing one of the key columns: {key_cols}")
                return

            # Merge the dataframes
            combined_df = pd.merge(expenses_df, revenue_df, on=key_cols, how='outer')

            # Ask for a location to save the file
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if not save_path:
                return # User cancelled the save dialog

            # Save the combined dataframe to an Excel file
            combined_df.to_excel(save_path, index=False)

            messagebox.showinfo("Success", f"Files combined and saved to {save_path}")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
