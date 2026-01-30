"""
CSV Catalog Creator
Scans a folder for CSV files and creates an Excel catalog with extracted date/time information
"""
#TODO: J'ajouterai la fonction de copie des fichiers .csv sur le serveur
#TODO: il y a l'autre script Matlab de Camille qu'on a pas de version définitive en .py
#TODO: Ce serait bien de pouvoir faire 'append' le catalogue si il existe déjà
#TODO: Un petit fichier avec les sexes ou autres infos pourrait automatique s'Ajouter au catalogue


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from pathlib import Path
import re
import pandas as pd


class CSVCatalogCreator:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Catalog Creator")
        self.root.geometry("600x500")
        
        self.selected_folder = None
        
        self.create_gui()
    
    def create_gui(self):
        """Create the main GUI layout"""
        # Title
        title_frame = ttk.Frame(self.root)
        title_frame.pack(fill='x', padx=20, pady=20)
        
        ttk.Label(title_frame, text="CSV Catalog Creator", 
                  font=('Arial', 16, 'bold')).pack()
        ttk.Label(title_frame, text="Select a folder to create a catalog of CSV files", 
                  font=('Arial', 10)).pack(pady=5)
        
        # Folder selection
        folder_frame = ttk.LabelFrame(self.root, text="Folder Selection", padding=20)
        folder_frame.pack(fill='x', padx=20, pady=10)
        
        self.folder_label = ttk.Label(folder_frame, text="No folder selected", 
                                      foreground='gray', wraplength=500)
        self.folder_label.pack(side='left', fill='x', expand=True)
        
        ttk.Button(folder_frame, text="Browse...", 
                  command=self.browse_folder).pack(side='right', padx=5)
        
        # Results area
        results_frame = ttk.LabelFrame(self.root, text="Results", padding=20)
        results_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        self.results_text = tk.Text(results_frame, height=10, width=60, wrap='word')
        self.results_text.pack(fill='both', expand=True)
        
        # Buttons
        button_frame = ttk.Frame(self.root)
        button_frame.pack(fill='x', padx=20, pady=10)
        
        self.create_btn = ttk.Button(button_frame, text="Create Catalog", 
                                     command=self.create_catalog, state='disabled')
        self.create_btn.pack(side='right')
    
    def browse_folder(self):
        """Browse for folder"""
        folder = filedialog.askdirectory(title="Select Folder with CSV Files")
        
        if folder:
            # Check if it's actually a folder
            if not os.path.isdir(folder):
                messagebox.showerror("Error", "Please select a folder, not a file")
                return
            
            self.selected_folder = folder
            self.folder_label.config(text=folder, foreground='black')
            self.create_btn['state'] = 'normal'
            
            # Preview CSV files
            self.preview_files()
    
    def preview_files(self):
        """Preview CSV files in the selected folder"""
        if not self.selected_folder:
            return
        
        # Find all CSV files
        csv_files = [f for f in os.listdir(self.selected_folder) 
                     if f.lower().endswith('.csv')]
        
        self.results_text.delete('1.0', tk.END)
        
        if not csv_files:
            self.results_text.insert('1.0', "No CSV files found in this folder.")
            return
        
        preview = f"Found {len(csv_files)} CSV files\n\n"
        preview += "Sample files:\n"
        for f in csv_files[:10]:
            preview += f"  - {f}\n"
        
        if len(csv_files) > 10:
            preview += f"\n... and {len(csv_files) - 10} more files"
        
        self.results_text.insert('1.0', preview)
    
    def create_catalog(self):
        """Create Excel catalog from CSV files"""
        if not self.selected_folder:
            messagebox.showerror("Error", "Please select a folder first")
            return
        
        try:
            # Find all CSV files
            csv_files = [f for f in os.listdir(self.selected_folder) 
                         if f.lower().endswith('.csv')]
            
            if not csv_files:
                messagebox.showerror("Error", "No CSV files found in the selected folder")
                return
            
            # Parse each filename and extract information
            catalog_data = []
            
            # Pattern: YYYY_MM_DD__HH_MM_SS_animalID.csv
            # Animal ID should be 3-4 digits
            pattern = r'^(\d{4})_(\d{2})_(\d{2})__(\d{2})_(\d{2})_(\d{2})_(\d{3,4})\.csv$'
            
            for filename in csv_files:
                match = re.match(pattern, filename)
                
                if match:
                    year, month, day, hour, minute, second, animal_id = match.groups()
                    
                    catalog_data.append({
                        'Filename': filename,
                        'Year': year,
                        'Month': month,
                        'Day': day,
                        'Hour': hour,
                        'Minute': minute,
                        'Second': second,
                        'Animal_ID': animal_id
                    })
                else:
                    # File doesn't match pattern - add with empty fields
                    catalog_data.append({
                        'Filename': filename,
                        'Year': '',
                        'Month': '',
                        'Day': '',
                        'Hour': '',
                        'Minute': '',
                        'Second': '',
                        'Animal_ID': ''
                    })
            
            # Create DataFrame
            df = pd.DataFrame(catalog_data)
            
            # Create Excel file in the same folder
            catalog_filename = 'CSV_Catalog.xlsx'
            catalog_path = os.path.join(self.selected_folder, catalog_filename)
            
            # Save to Excel
            df.to_excel(catalog_path, index=False, engine='openpyxl')
            
            # Show results
            result_msg = f"Catalog created successfully!\n\n"
            result_msg += f"File: {catalog_filename}\n"
            result_msg += f"Location: {self.selected_folder}\n\n"
            result_msg += f"Total files: {len(csv_files)}\n"
            result_msg += f"Files matching pattern: {len([d for d in catalog_data if d['Year']])}\n"
            result_msg += f"Files not matching pattern: {len([d for d in catalog_data if not d['Year']])}\n"
            
            self.results_text.delete('1.0', tk.END)
            self.results_text.insert('1.0', result_msg)
            
            messagebox.showinfo("Success", f"Catalog created successfully!\n\n{catalog_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create catalog:\n{str(e)}")


def main():
    root = tk.Tk()
    app = CSVCatalogCreator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
