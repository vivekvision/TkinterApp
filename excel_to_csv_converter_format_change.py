import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import re
from datetime import datetime

class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to CSV Converter")
        self.root.geometry("600x500")
        
        # Store selected files
        self.selected_files = []
        
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
    
    def clean_text(self, text):
        """Clean text by removing newlines and page breaks"""
        if pd.isna(text):
            return text
        # Convert to string if not already
        text = str(text)
        # Remove newlines and page breaks
        text = re.sub(r'[\n\r]+', ' ', text)
        # Remove multiple spaces
        text = re.sub(r'\s+', ' ', text)
        return text.strip()
    
    def format_date(self, value):
        """Format date to YYYY-MM-DD format"""
        if pd.isna(value):
            return value
        
        try:
            # Try to parse the date
            if isinstance(value, datetime):
                return value.strftime('%Y-%m-%d')
            elif isinstance(value, str):
                # Try different date formats
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
        # Create a copy of the dataframe
        processed_df = df.copy()
        
        # Process each column
        for column in processed_df.columns:
            # Check if column might contain dates
            if processed_df[column].dtype == 'datetime64[ns]':
                processed_df[column] = processed_df[column].apply(self.format_date)
            else:
                # Clean text in non-date columns
                processed_df[column] = processed_df[column].apply(self.clean_text)
        
        return processed_df
        
    def select_files(self):
        # Open file dialog to select multiple Excel files
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
        # Get selected indices
        selected_indices = self.files_listbox.curselection()
        
        # Remove from both listbox and selected_files list
        for index in reversed(selected_indices):
            self.files_listbox.delete(index)
            self.selected_files.pop(index)
        
        self.update_status()
        if not self.selected_files:
            self.convert_button.config(state='disabled')
    
    def update_status(self):
        count = len(self.selected_files)
        self.status_label.config(
            text=f"{count} file(s) selected",
            fg='black'
        )
    
    def convert_files(self):
        if not self.selected_files:
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
                
                # Generate output CSV path
                output_path = os.path.splitext(file_path)[0] + '.csv'
                
                # Save as CSV with UTF-8 encoding
                processed_df.to_csv(output_path, index=False, encoding='utf-8')
                success_count += 1
                
            except Exception as e:
                error_count += 1
                error_messages.append(f"Error converting {os.path.basename(file_path)}: {str(e)}")
        
        # Show results
        if success_count > 0:
            messagebox.showinfo(
                "Conversion Complete",
                f"Successfully converted {success_count} file(s).\n"
                f"Failed to convert {error_count} file(s)."
            )
        
        if error_messages:
            error_text = "\n\n".join(error_messages)
            messagebox.showerror("Conversion Errors", error_text)
        
        # Update status
        self.status_label.config(
            text=f"Conversion complete: {success_count} succeeded, {error_count} failed",
            fg='green' if error_count == 0 else 'orange'
        )

def main():
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main() 