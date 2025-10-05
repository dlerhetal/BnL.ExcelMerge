
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

class AdvancedExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Advanced Excel Merger")

        self.expense_file_path = tk.StringVar()
        self.revenue_file_path = tk.StringVar()
        self.output_file_name = tk.StringVar()

        # Main frame
        frame = tk.Frame(root, padx=10, pady=10)
        frame.pack(padx=10, pady=10)

        # Expense file selection
        tk.Label(frame, text="Expense File:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(frame, textvariable=self.expense_file_path, width=50, state='readonly').grid(row=0, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_expense_file).grid(row=0, column=2, padx=5, pady=5)

        # Revenue file selection
        tk.Label(frame, text="Revenue File:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(frame, textvariable=self.revenue_file_path, width=50, state='readonly').grid(row=1, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_revenue_file).grid(row=1, column=2, padx=5, pady=5)

        # Output file name
        tk.Label(frame, text="Output File Name:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(frame, textvariable=self.output_file_name, width=50, state='readonly').grid(row=2, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_output_file).grid(row=2, column=2, padx=5, pady=5)


        # Preview button
        self.preview_button = tk.Button(root, text="Preview Columns", command=self.preview_columns, state=tk.DISABLED)
        self.preview_button.pack(pady=10)

    def browse_expense_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.expense_file_path.set(filename)
            self.check_inputs()

    def browse_revenue_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filename:
            self.revenue_file_path.set(filename)
            self.check_inputs()

    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.output_file_name.set(filename)
            self.check_inputs()

    def check_inputs(self, *args):
        if self.expense_file_path.get() and self.revenue_file_path.get() and self.output_file_name.get():
            self.preview_button.config(state=tk.NORMAL)
        else:
            self.preview_button.config(state=tk.DISABLED)

    def preview_columns(self):
        try:
            self.df_expense = pd.read_excel(self.expense_file_path.get())
            self.df_revenue = pd.read_excel(self.revenue_file_path.get())

            # Get all unique column names
            all_columns = sorted(list(set(self.df_expense.columns) | set(self.df_revenue.columns)))

            # Create a new window for column selection
            self.col_window = tk.Toplevel(self.root)
            self.col_window.title("Select Columns")

            self.col_vars = {}
            for i, col in enumerate(all_columns):
                var = tk.BooleanVar(value=True)
                cb = tk.Checkbutton(self.col_window, text=col, variable=var)
                cb.grid(row=i, column=0, sticky='w', padx=10, pady=2)
                self.col_vars[col] = var

            combine_button = tk.Button(self.col_window, text="Combine", command=self.combine_and_save)
            combine_button.grid(row=len(all_columns), column=0, pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Could not read files: {e}")

    def combine_and_save(self):
        try:
            selected_columns = [col for col, var in self.col_vars.items() if var.get()]

            if not selected_columns:
                messagebox.showwarning("Warning", "No columns selected.")
                return

            key_cols = ['Job ID', 'Cost Code ID', 'Phase ID']
            if not all(col in self.df_expense.columns and col in self.df_revenue.columns for col in key_cols):
                 messagebox.showerror("Error", f"Both files must contain the key columns: {key_cols}")
                 return
            
            # Ensure key columns are in selected columns
            for col in key_cols:
                if col not in selected_columns:
                    selected_columns.insert(0, col)


            merged_df = pd.merge(self.df_expense, self.df_revenue, on=key_cols, how='outer')

            # Filter for selected columns
            final_df = merged_df[[col for col in selected_columns if col in merged_df.columns]]

            output_path = self.output_file_name.get()

            final_df.to_excel(output_path, index=False)

            messagebox.showinfo("Success", f"Combined file saved to {output_path}")
            self.col_window.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during merging: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = AdvancedExcelMergerApp(root)
    root.mainloop()
