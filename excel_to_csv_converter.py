import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

class ExcelToCSVConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel to CSV Converter")
        self.root.geometry("500x300")
        
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
        self.title_label.pack(pady=20)
        
        # Select File Button
        self.select_button = tk.Button(
            self.main_frame,
            text="Select Excel File",
            command=self.select_file,
            font=("Helvetica", 12),
            bg='#4CAF50',
            fg='white',
            padx=20,
            pady=10
        )
        self.select_button.pack(pady=20)
        
        # Status Label
        self.status_label = tk.Label(
            self.main_frame,
            text="No file selected",
            font=("Helvetica", 10),
            bg='#f0f0f0'
        )
        self.status_label.pack(pady=10)
        
    def select_file(self):
        # Open file dialog to select Excel file
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            try:
                # Read the Excel file
                df = pd.read_excel(file_path)
                
                # Generate output CSV path (same folder, same name, different extension)
                output_path = os.path.splitext(file_path)[0] + '.csv'
                
                # Save as CSV
                df.to_csv(output_path, index=False)
                
                # Update status
                self.status_label.config(
                    text=f"Successfully converted to:\n{output_path}",
                    fg='green'
                )
                
                messagebox.showinfo(
                    "Success",
                    f"File has been converted successfully!\nSaved as: {output_path}"
                )
                
            except Exception as e:
                self.status_label.config(
                    text=f"Error: {str(e)}",
                    fg='red'
                )
                messagebox.showerror("Error", f"An error occurred:\n{str(e)}")

def main():
    root = tk.Tk()
    app = ExcelToCSVConverter(root)
    root.mainloop()

if __name__ == "__main__":
    main() 