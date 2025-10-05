
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd

class ProExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pro Excel Merger")

        self.expense_file_path = tk.StringVar()
        self.revenue_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.key_columns = ['Job ID', 'Cost Code ID', 'Phase ID']

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

        # Output file path
        tk.Label(frame, text="Output File:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(frame, textvariable=self.output_file_path, width=50, state='readonly').grid(row=2, column=1, padx=5, pady=5)
        tk.Button(frame, text="Browse...", command=self.browse_output_file).grid(row=2, column=2, padx=5, pady=5)

        # Next button
        self.next_button = tk.Button(root, text="Configure Columns", command=self.open_column_config, state=tk.DISABLED)
        self.next_button.pack(pady=10)

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
            self.output_file_path.set(filename)
            self.check_inputs()

    def check_inputs(self, *args):
        if self.expense_file_path.get() and self.revenue_file_path.get() and self.output_file_path.get():
            self.next_button.config(state=tk.NORMAL)
        else:
            self.next_button.config(state=tk.DISABLED)

    def open_column_config(self):
        try:
            self.df_expense = pd.read_excel(self.expense_file_path.get())
            self.df_revenue = pd.read_excel(self.revenue_file_path.get())

            self.original_columns = sorted(list(set(self.df_expense.columns) | set(self.df_revenue.columns)))
            self.rename_map = {col: col for col in self.original_columns}

            self.col_config_window = tk.Toplevel(self.root)
            self.col_config_window.title("Configure Columns")

            list_frame = tk.Frame(self.col_config_window)
            list_frame.grid(row=0, column=0, padx=10, pady=10, rowspan=5)

            self.listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, width=40, height=15)
            self.update_listbox()
            self.listbox.pack(side=tk.LEFT, fill=tk.BOTH)

            scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)
            scrollbar.config(command=self.listbox.yview)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            self.listbox.config(yscrollcommand=scrollbar.set)

            tk.Button(self.col_config_window, text="Up", command=self.move_up).grid(row=0, column=1, padx=5, pady=2, sticky='ew')
            tk.Button(self.col_config_window, text="Down", command=self.move_down).grid(row=1, column=1, padx=5, pady=2, sticky='ew')
            tk.Button(self.col_config_window, text="Rename", command=self.rename_column).grid(row=2, column=1, padx=5, pady=2, sticky='ew')
            tk.Button(self.col_config_window, text="Remove", command=self.remove_column).grid(row=3, column=1, padx=5, pady=2, sticky='ew')

            self.key_columns_label = tk.Label(self.col_config_window, text=f"Keys: {self.key_columns}")
            self.key_columns_label.grid(row=4, column=0, padx=5, pady=5, sticky='w')

            key_col_button = tk.Button(self.col_config_window, text="Select Key Columns", command=self.select_key_columns)
            key_col_button.grid(row=5, column=0, padx=5, pady=5, sticky='ew')

            reset_keys_button = tk.Button(self.col_config_window, text="Reset Keys", command=self.reset_key_columns)
            reset_keys_button.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

            tk.Button(self.col_config_window, text="Combine and Save", command=self.combine_and_save).grid(row=6, column=0, columnspan=2, pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Could not read files: {e}")

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for col in self.original_columns:
            self.listbox.insert(tk.END, self.rename_map[col])

    def move_up(self):
        try:
            idx = self.listbox.curselection()[0]
            if idx > 0:
                self.original_columns.insert(idx - 1, self.original_columns.pop(idx))
                self.update_listbox()
                self.listbox.selection_set(idx - 1)
        except IndexError:
            pass

    def move_down(self):
        try:
            idx = self.listbox.curselection()[0]
            if idx < len(self.original_columns) - 1:
                self.original_columns.insert(idx + 1, self.original_columns.pop(idx))
                self.update_listbox()
                self.listbox.selection_set(idx + 1)
        except IndexError:
            pass

    def rename_column(self):
        try:
            idx = self.listbox.curselection()[0]
            original_col_name = self.original_columns[idx]
            old_display_name = self.rename_map[original_col_name]
            
            new_name = simpledialog.askstring("Rename Column", f"Enter new name for '{old_display_name}':", parent=self.col_config_window)
            if new_name:
                self.rename_map[original_col_name] = new_name
                self.update_listbox()
                self.listbox.selection_set(idx)
        except IndexError:
            messagebox.showwarning("Warning", "Please select a column to rename.", parent=self.col_config_window)

    def remove_column(self):
        try:
            idx = self.listbox.curselection()[0]
            self.original_columns.pop(idx)
            self.update_listbox()
        except IndexError:
            messagebox.showwarning("Warning", "Please select a column to remove.", parent=self.col_config_window)

    def select_key_columns(self):
        common_columns = sorted(list(set(self.df_expense.columns) & set(self.df_revenue.columns)))
        
        select_window = tk.Toplevel(self.col_config_window)
        select_window.title("Select Key Columns")
        
        listbox = tk.Listbox(select_window, selectmode=tk.MULTIPLE, width=40, height=15)
        for col in common_columns:
            listbox.insert(tk.END, col)
            if col in self.key_columns:
                listbox.selection_set(tk.END)
        listbox.pack(padx=10, pady=10)

        def on_ok():
            selected_indices = listbox.curselection()
            self.key_columns = [listbox.get(i) for i in selected_indices]
            self.key_columns_label.config(text=f"Keys: {self.key_columns}")
            select_window.destroy()

        tk.Button(select_window, text="OK", command=on_ok).pack(pady=5)

    def reset_key_columns(self):
        self.key_columns = ['Job ID', 'Cost Code ID', 'Phase ID']
        self.key_columns_label.config(text=f"Keys: {self.key_columns}")

    def combine_and_save(self):
        try:
            if not self.key_columns:
                messagebox.showerror("Error", "Please select at least one key column.", parent=self.col_config_window)
                return

            if not all(col in self.df_expense.columns and col in self.df_revenue.columns for col in self.key_columns):
                 messagebox.showerror("Error", f"Both files must contain the key columns: {self.key_columns}", parent=self.col_config_window)
                 return

            merged_df = pd.merge(self.df_expense, self.df_revenue, on=self.key_columns, how='outer')

            final_df = merged_df[self.original_columns]
            
            final_df = final_df.rename(columns=self.rename_map)

            final_df.to_excel(self.output_file_path.get(), index=False)

            messagebox.showinfo("Success", f"Combined file saved to {self.output_file_path.get()}")
            self.col_config_window.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during merging: {e}", parent=self.col_config_window)

if __name__ == "__main__":
    root = tk.Tk()
    app = ProExcelMergerApp(root)
    root.mainloop()
