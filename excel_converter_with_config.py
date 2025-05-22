import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
import json
from datetime import datetime
from tkcalendar import DateEntry

class ExcelConverterWithConfig:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to CSV Converter with Config")
        self.root.geometry("700x600")
        
        # Store selected files
        self.selected_files = []
        self.selected_date = datetime.now()
        self.selected_format = None
        self.selected_file_name = None
        
        # Load configuration
        self.load_config()
        
        # Configure the main window
        self.root.configure(bg='#f0f0f0')
        
        # Create and configure the main frame
        self.main_frame = tk.Frame(root, bg='#f0f0f0', padx=20, pady=20)
        self.main_frame.pack(expand=True, fill='both')
        
        # Title Label
        self.title_label = tk.Label(
            self.main_frame,
            text="Excel to CSV Converter",
            font=("Helvetica", 16, "bold"),
            bg='#f0f0f0'
        )
        self.title_label.pack(pady=10)
        
        # Configuration Frame
        self.config_frame = tk.LabelFrame(
            self.main_frame,
            text="Output Configuration",
            font=("Helvetica", 10, "bold"),
            bg='#f0f0f0',
            padx=10,
            pady=10
        )
        self.config_frame.pack(fill='x', pady=10)
        
        # Format Selection
        self.format_label = tk.Label(
            self.config_frame,
            text="Output Format:",
            font=("Helvetica", 10),
            bg='#f0f0f0'
        )
        self.format_label.grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        # Create a dictionary to store format mappings
        self.format_mapping = {format['display_name']: format['file_name'] 
                             for format in self.config['output_formats']}
        
        self.format_combo = ttk.Combobox(
            self.config_frame,
            values=list(self.format_mapping.keys()),
            state='readonly',
            width=30
        )
        self.format_combo.grid(row=0, column=1, padx=5, pady=5)
        self.format_combo.bind('<<ComboboxSelected>>', self.on_format_select)
        
        # Date Selection
        self.date_label = tk.Label(
            self.config_frame,
            text="Select Date:",
            font=("Helvetica", 10),
            bg='#f0f0f0'
        )
        self.date_label.grid(row=1, column=0, padx=5, pady=5, sticky='w')
        
        self.date_picker = DateEntry(
            self.config_frame,
            width=30,
            background='#4CAF50',
            foreground='white',
            borderwidth=2,
            date_pattern='yyyy-mm-dd'
        )
        self.date_picker.grid(row=1, column=1, padx=5, pady=5)
        self.date_picker.bind('<<DateEntrySelected>>', self.on_date_select)
        
        # Select Files Button
        self.select_button = tk.Button(
            self.main_frame,
            text="Select Excel Files",
            command=self.select_files,
            font=("Helvetica", 12),
            bg='#4CAF50',
            fg='white',
            padx=20,
            pady=10
        )
        self.select_button.pack(pady=10)
        
        # Files List Frame
        self.list_frame = tk.Frame(self.main_frame, bg='#f0f0f0')
        self.list_frame.pack(fill='both', expand=True, pady=10)
        
        # Files List Label
        self.list_label = tk.Label(
            self.list_frame,
            text="Selected Files:",
            font=("Helvetica", 10, "bold"),
            bg='#f0f0f0'
        )
        self.list_label.pack(anchor='w')
        
        # Files Listbox with Scrollbar
        self.listbox_frame = tk.Frame(self.list_frame, bg='#f0f0f0')
        self.listbox_frame.pack(fill='both', expand=True)
        
        self.scrollbar = ttk.Scrollbar(self.listbox_frame)
        self.scrollbar.pack(side='right', fill='y')
        
        self.files_listbox = tk.Listbox(
            self.listbox_frame,
            yscrollcommand=self.scrollbar.set,
            font=("Helvetica", 10),
            bg='white',
            selectmode='extended'
        )
        self.files_listbox.pack(side='left', fill='both', expand=True)
        self.scrollbar.config(command=self.files_listbox.yview)
        
        # Remove Selected Button
        self.remove_button = tk.Button(
            self.list_frame,
            text="Remove Selected",
            command=self.remove_selected,
            font=("Helvetica", 10),
            bg='#ff4444',
            fg='white',
            padx=10,
            pady=5
        )
        self.remove_button.pack(pady=5)
        
        # Convert Button
        self.convert_button = tk.Button(
            self.main_frame,
            text="Convert to CSV",
            command=self.convert_files,
            font=("Helvetica", 12),
            bg='#2196F3',
            fg='white',
            padx=20,
            pady=10,
            state='disabled'
        )
        self.convert_button.pack(pady=10)
        
        # Status Label
        self.status_label = tk.Label(
            self.main_frame,
            text="No files selected",
            font=("Helvetica", 10),
            bg='#f0f0f0'
        )
        self.status_label.pack(pady=10)
    
    def load_config(self):
        """Load configuration from JSON file"""
        try:
            with open('config.json', 'r') as f:
                self.config = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load configuration: {str(e)}")
            self.config = {"output_formats": []}
    
    def on_format_select(self, event):
        """Handle format selection"""
        self.selected_format = self.format_combo.get()
        self.selected_file_name = self.format_mapping[self.selected_format]
        self.update_status()
    
    def on_date_select(self, event):
        """Handle date selection"""
        self.selected_date = self.date_picker.get_date()
        self.update_status()
    
    def clean_text(self, text):
        """Clean text by removing newlines and page breaks"""
        if pd.isna(text):
            return text
        text = str(text)
        text = re.sub(r'[\n\r]+', ' ', text)
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def format_date(self, value):
        """Format date to YYYY-MM-DD format"""
        if pd.isna(value):
            return value
        
        try:
            if isinstance(value, datetime):
                return value.strftime('%Y-%m-%d')
            elif isinstance(value, str):
                for fmt in ['%Y-%m-%d', '%d/%m/%Y', '%m/%d/%Y', '%d-%m-%Y', '%m-%d-%Y']:
                    try:
                        return datetime.strptime(value, fmt).strftime('%Y-%m-%d')
                    except ValueError:
                        continue
            return value
        except:
            return value
    
    def process_dataframe(self, df):
        """Process dataframe to clean text and format dates"""
        processed_df = df.copy()
        
        for column in processed_df.columns:
            if processed_df[column].dtype == 'datetime64[ns]':
                processed_df[column] = processed_df[column].apply(self.format_date)
            else:
                processed_df[column] = processed_df[column].apply(self.clean_text)
        
        return processed_df
    
    def get_output_filename(self, original_path):
        """Generate output filename based on configuration and date"""
        if not self.selected_file_name:
            return os.path.splitext(original_path)[0] + '.csv'
        
        directory = os.path.dirname(original_path)
        date_str = self.selected_date.strftime('%Y%m%d')
        return os.path.join(directory, f"{self.selected_file_name}_{date_str}.csv")
    
    def select_files(self):
        file_paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_paths:
            for file_path in file_paths:
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.files_listbox.insert(tk.END, os.path.basename(file_path))
            
            self.update_status()
            self.convert_button.config(state='normal')
    
    def remove_selected(self):
        selected_indices = self.files_listbox.curselection()
        
        for index in reversed(selected_indices):
            self.files_listbox.delete(index)
            self.selected_files.pop(index)
        
        self.update_status()
        if not self.selected_files:
            self.convert_button.config(state='disabled')
    
    def update_status(self):
        count = len(self.selected_files)
        format_text = f"Format: {self.selected_format}" if self.selected_format else "No format selected"
        date_text = f"Date: {self.selected_date.strftime('%Y-%m-%d')}"
        self.status_label.config(
            text=f"{count} file(s) selected | {format_text} | {date_text}",
            fg='black'
        )
    
    def convert_files(self):
        if not self.selected_files:
            return
        
        if not self.selected_format:
            messagebox.showwarning("Warning", "Please select an output format")
            return
        
        success_count = 0
        error_count = 0
        error_messages = []
        
        for file_path in self.selected_files:
            try:
                # Read the Excel file
                df = pd.read_excel(file_path)
                
                # Process the dataframe
                processed_df = self.process_dataframe(df)
                
                # Generate output path with configured format and date
                output_path = self.get_output_filename(file_path)
                
                # Save as CSV with UTF-8 encoding
                processed_df.to_csv(output_path, index=False, encoding='utf-8')
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_messages.append(f"Error converting {os.path.basename(file_path)}: {str(e)}")
        
        if success_count > 0:
            messagebox.showinfo(
                "Conversion Complete",
                f"Successfully converted {success_count} file(s).\n"
                f"Failed to convert {error_count} file(s)."
            )
        
        if error_messages:
            error_text = "\n\n".join(error_messages)
            messagebox.showerror("Conversion Errors", error_text)
        
        self.status_label.config(
            text=f"Conversion complete: {success_count} succeeded, {error_count} failed",
            fg='green' if error_count == 0 else 'orange'
        )

def main():
    root = tk.Tk()
    app = ExcelConverterWithConfig(root)
    root.mainloop()

if __name__ == "__main__":
    main() 