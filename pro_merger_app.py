
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from PIL import Image, ImageTk
import webbrowser
import json

class ProExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Pro Excel Merger")

        self.expense_file_path = tk.StringVar()
        self.revenue_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        self.key_columns = ['Job ID', 'Cost Code ID', 'Phase ID']
        self.required_column = None
        self.original_columns = []
        self.rename_map = {}
        self.config_file = 'merger_config.json'
        self.version = "1.1.1"

        self.create_menu()

        # Main frame
        self.main_frame = tk.Frame(root, padx=10, pady=10)
        self.main_frame.pack(padx=10, pady=10)

        # Add logo
        self.set_logo("C:/Users/dale/OneDrive/Documents/ISI/BnL/Sage/Images/BnL.Logo.jpg", self.main_frame)

        # Expense file selection
        tk.Label(self.main_frame, text="Expense File:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(self.main_frame, textvariable=self.expense_file_path, width=50, state='readonly').grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.main_frame, text="Browse...", command=self.browse_expense_file).grid(row=1, column=2, padx=5, pady=5)

        # Revenue file selection
        tk.Label(self.main_frame, text="Revenue File:").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(self.main_frame, textvariable=self.revenue_file_path, width=50, state='readonly').grid(row=2, column=1, padx=5, pady=5)
        tk.Button(self.main_frame, text="Browse...", command=self.browse_revenue_file).grid(row=2, column=2, padx=5, pady=5)

        # Output file path
        tk.Label(self.main_frame, text="Output File:").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        tk.Entry(self.main_frame, textvariable=self.output_file_path, width=50, state='readonly').grid(row=3, column=1, padx=5, pady=5)
        tk.Button(self.main_frame, text="Browse...", command=self.browse_output_file).grid(row=3, column=2, padx=5, pady=5)

        # Next button
        self.next_button = tk.Button(root, text="Configure Columns", command=self.open_column_config, state=tk.DISABLED)
        self.next_button.pack(pady=10)

    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        filemenu = tk.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Select Logo", command=self.select_logo_config)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=filemenu)

        helpmenu = tk.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="About", command=self.show_about)
        menubar.add_cascade(label="Help", menu=helpmenu)

        self.root.config(menu=menubar)

    def show_about(self):
        about_window = tk.Toplevel(self.root)
        about_window.title("About Pro Excel Merger")
        
        logo_path = "C:/Users/dale/OneDrive/Documents/ISI/BnL/Sage/Images/IDS.Logo.png"
        logo_image = Image.open(logo_path)
        logo_image = logo_image.resize((100, 100), Image.LANCZOS)
        self.about_logo_photo = ImageTk.PhotoImage(logo_image)
        
        tk.Label(about_window, image=self.about_logo_photo).pack(pady=10)
        tk.Label(about_window, text=f"Version: {self.version}").pack(pady=5)
        
        tk.Label(about_window, text="Contact: Dale Linn").pack(pady=5)
        email = tk.Label(about_window, text="richard.linn@biz-solve.com", fg="blue", cursor="hand2")
        email.pack()
        email.bind("<Button-1>", lambda e: webbrowser.open_new("mailto:richard.linn@biz-solve.com"))

        tk.Label(about_window, text="For new versions, please visit:").pack(pady=5)
        link = tk.Label(about_window, text="https://github.com/dlerhetal/BnL.ExcelMerge/releases", fg="blue", cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/dlerhetal/BnL.ExcelMerge/releases"))

    def set_logo(self, logo_path, frame):
        try:
            logo_image = Image.open(logo_path)
            aspect_ratio = logo_image.width / logo_image.height
            logo_image = logo_image.resize((int(50 * aspect_ratio), 50), Image.LANCZOS)
            self.logo_photo = ImageTk.PhotoImage(logo_image)
            
            if hasattr(self, 'logo_label'):
                self.logo_label.config(image=self.logo_photo)
            else:
                self.logo_label = tk.Label(frame, image=self.logo_photo)
                self.logo_label.grid(row=0, column=0, columnspan=3, pady=10)

        except Exception as e:
            messagebox.showerror("Error", f"Could not load logo: {e}")

    def select_logo_config(self):
        filename = filedialog.askopenfilename(filetypes=[
            ("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
            ("All files", "*.*",)
        ])
        if filename:
            self.set_logo(filename, self.main_frame)

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

            if not self.original_columns:
                self.original_columns = sorted(list(set(self.df_expense.columns) | set(self.df_revenue.columns)))
                self.rename_map = {col: col for col in self.original_columns}
                self.default_original_columns = self.original_columns.copy()
                self.default_rename_map = self.rename_map.copy()

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
            self.key_columns_label.grid(row=5, column=0, padx=5, pady=5, sticky='w')

            key_col_button = tk.Button(self.col_config_window, text="Select Key Columns", command=self.select_key_columns)
            key_col_button.grid(row=6, column=0, padx=5, pady=5, sticky='ew')

            reset_keys_button = tk.Button(self.col_config_window, text="Reset Keys", command=self.reset_key_columns)
            reset_keys_button.grid(row=6, column=1, padx=5, pady=5, sticky='ew')

            self.required_column_label = tk.Label(self.col_config_window, text=f"Required Column: {self.required_column if self.required_column else 'None'}")
            self.required_column_label.grid(row=7, column=0, padx=5, pady=5, sticky='w')

            required_col_button = tk.Button(self.col_config_window, text="Select Required Column", command=self.select_required_column)
            required_col_button.grid(row=8, column=0, padx=5, pady=5, sticky='ew')

            reset_required_col_button = tk.Button(self.col_config_window, text="Reset Required Column", command=self.reset_required_column)
            reset_required_col_button.grid(row=8, column=1, padx=5, pady=5, sticky='ew')

            tk.Button(self.col_config_window, text="Save Configuration", command=self.save_app_configuration).grid(row=9, column=0, padx=5, pady=5, sticky='ew')
            tk.Button(self.col_config_window, text="Load Saved Configuration", command=self.load_app_configuration).grid(row=9, column=1, padx=5, pady=5, sticky='ew')
            tk.Button(self.col_config_window, text="Reset to Default", command=self.reset_to_default).grid(row=10, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

            tk.Button(self.col_config_window, text="Combine and Save", command=self.combine_and_save).grid(row=11, column=0, columnspan=2, pady=10)

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

    def select_required_column(self):
        all_columns = sorted(list(set(self.df_expense.columns) | set(self.df_revenue.columns)))
        
        select_window = tk.Toplevel(self.col_config_window)
        select_window.title("Select Required Column")
        
        listbox = tk.Listbox(select_window, selectmode=tk.SINGLE, width=40, height=15)
        for col in all_columns:
            listbox.insert(tk.END, col)
            if col == self.required_column:
                listbox.selection_set(tk.END)
        listbox.pack(padx=10, pady=10)

        def on_ok():
            selected_indices = listbox.curselection()
            if selected_indices:
                self.required_column = listbox.get(selected_indices[0])
            else:
                self.required_column = None
            self.required_column_label.config(text=f"Required Column: {self.required_column if self.required_column else 'None'}")
            select_window.destroy()

        tk.Button(select_window, text="OK", command=on_ok).pack(pady=5)

    def reset_required_column(self):
        self.required_column = None
        self.required_column_label.config(text=f"Required Column: None")

    def save_app_configuration(self):
        config_data = {
            'original_columns': self.original_columns,
            'rename_map': self.rename_map,
            'key_columns': self.key_columns,
            'required_column': self.required_column
        }
        with open(self.config_file, 'w') as f:
            json.dump(config_data, f, indent=4)
        messagebox.showinfo("Success", "Configuration saved.")

    def load_app_configuration(self):
        try:
            with open(self.config_file, 'r') as f:
                config_data = json.load(f)
                self.original_columns = config_data.get('original_columns', [])
                self.rename_map = config_data.get('rename_map', {})
                self.key_columns = config_data.get('key_columns', ['Job ID', 'Cost Code ID', 'Phase ID'])
                self.required_column = config_data.get('required_column', None)
            self.update_listbox()
            self.key_columns_label.config(text=f"Keys: {self.key_columns}")
            self.required_column_label.config(text=f"Required Column: {self.required_column if self.required_column else 'None'}")
            messagebox.showinfo("Success", "Configuration loaded.")
        except FileNotFoundError:
            messagebox.showerror("Error", "No saved configuration file found.")

    def reset_to_default(self):
        self.original_columns = self.default_original_columns.copy()
        self.rename_map = self.default_rename_map.copy()
        self.key_columns = ['Job ID', 'Cost Code ID', 'Phase ID']
        self.required_column = None
        self.update_listbox()
        self.key_columns_label.config(text=f"Keys: {self.key_columns}")
        self.required_column_label.config(text=f"Required Column: {self.required_column if self.required_column else 'None'}")
        messagebox.showinfo("Success", "Configuration reset to default.")
    def combine_and_save(self):
        try:
            if not self.key_columns:
                messagebox.showerror("Error", "Please select at least one key column.", parent=self.col_config_window)
                return

            if self.required_column:
                self.df_expense.dropna(subset=[self.required_column], inplace=True)
                self.df_revenue.dropna(subset=[self.required_column], inplace=True)

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
