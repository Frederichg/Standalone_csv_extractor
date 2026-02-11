"""
CSV Trial Extractor - Behavior Analysis Data Processing
Extracts trial-based data from CSV files using a catalog-based approach
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
from pathlib import Path
import traceback


class CSVTrialExtractor:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV Trial Extractor")
        self.root.geometry("1000x800")
        
        # Data storage
        self.catalog_path = None
        self.catalog_df = None
        self.catalog_dir = None
        self.sheet_name = None
        self.sample_csv_df = None
        self.trial_separator = None
        
        # Marker storage (list of dicts)
        self.markers = []
        
        self.create_gui()
    
    def create_gui(self):
        """Create the main GUI layout"""
        # Create notebook for tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Tab 1: File Selection & Configuration
        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="1. File Selection")
        self.create_file_selection_tab(tab1)
        
        # Tab 2: Trial Configuration
        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="2. Trial Configuration")
        self.create_trial_config_tab(tab2)
        
        # Tab 3: Marker Configuration
        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text="3. Marker Configuration")
        self.create_marker_config_tab(tab3)
        
        # Bottom: Progress and Execute
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill='x', padx=5, pady=5)
        
        self.progress_label = ttk.Label(bottom_frame, text="Ready", font=('Arial', 10))
        self.progress_label.pack(side='left', padx=5)
        
        self.execute_btn = ttk.Button(bottom_frame, text="Start Data Crunching", 
                                      command=self.execute_extraction, state='disabled')
        self.execute_btn.pack(side='right', padx=5)
    
    def create_file_selection_tab(self, parent):
        """Create file selection and configuration UI"""
        # Catalog file selection
        catalog_frame = ttk.LabelFrame(parent, text="Catalog File", padding=10)
        catalog_frame.pack(fill='x', padx=10, pady=5)
        
        self.catalog_label = ttk.Label(catalog_frame, text="No file selected", foreground='gray')
        self.catalog_label.pack(side='left', fill='x', expand=True)
        
        ttk.Button(catalog_frame, text="Browse...", command=self.browse_catalog).pack(side='right')
        
        # Sheet name input
        sheet_frame = ttk.LabelFrame(parent, text="Excel Sheet Configuration", padding=10)
        sheet_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(sheet_frame, text="Sheet Name:").grid(row=0, column=0, sticky='w', pady=5)
        self.sheet_entry = ttk.Entry(sheet_frame, width=30)
        self.sheet_entry.grid(row=0, column=1, padx=5, pady=5)
        self.sheet_entry.insert(0, "data")
        
        ttk.Button(sheet_frame, text="Load Sheet", command=self.load_sheet).grid(row=0, column=2, padx=5)
        
        # Column selection
        col_frame = ttk.LabelFrame(parent, text="Column Selection", padding=10)
        col_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(col_frame, text="File Name Column:").grid(row=0, column=0, sticky='w', pady=5)
        self.filename_col_combo = ttk.Combobox(col_frame, state='readonly', width=25)
        self.filename_col_combo.grid(row=0, column=1, padx=5, pady=5)
        self.filename_col_combo.bind('<<ComboboxSelected>>', self.update_suggested_csv)
        
        ttk.Label(col_frame, text="Experiment Type Column:").grid(row=1, column=0, sticky='w', pady=5)
        self.exptype_col_combo = ttk.Combobox(col_frame, state='readonly', width=25)
        self.exptype_col_combo.grid(row=1, column=1, padx=5, pady=5)
        self.exptype_col_combo.bind('<<ComboboxSelected>>', self.update_experiment_types)
        
        ttk.Label(col_frame, text="Select Experiment Type:").grid(row=2, column=0, sticky='w', pady=5)
        self.exptype_filter_combo = ttk.Combobox(col_frame, state='readonly', width=25)
        self.exptype_filter_combo.grid(row=2, column=1, padx=5, pady=5)
        
        # Preview area
        preview_frame = ttk.LabelFrame(parent, text="Catalog Preview", padding=10)
        preview_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.preview_text = scrolledtext.ScrolledText(preview_frame, height=10, width=80)
        self.preview_text.pack(fill='both', expand=True)
    
    def create_trial_config_tab(self, parent):
        """Create trial configuration UI"""
        # Suggested CSV filename display
        suggest_frame = ttk.LabelFrame(parent, text="Suggested CSV File", padding=10)
        suggest_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(suggest_frame, text="Suggested CSV file:").pack(side='left')
        self.suggested_csv_label = ttk.Label(suggest_frame, text="No experiment type selected", foreground='gray')
        self.suggested_csv_label.pack(side='left', padx=10, fill='x', expand=True)
        
        self.browse_csv_btn = ttk.Button(suggest_frame, text="Browse CSV...", 
                                        command=self.browse_csv_file, state='disabled')
        self.browse_csv_btn.pack(side='right', padx=5)
        
        # Load sample button
        load_frame = ttk.Frame(parent, padding=10)
        load_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(load_frame, text="Load Sample CSV", command=self.load_sample_csv_manual).pack(side='left')
        self.sample_status_label = ttk.Label(load_frame, text="No sample loaded", foreground='gray')
        self.sample_status_label.pack(side='left', padx=10)
        
        # Trial separator selection
        sep_frame = ttk.LabelFrame(parent, text="Trial Identification", padding=10)
        sep_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Label(sep_frame, text="Trial Separator (from 'state' column):").grid(row=0, column=0, sticky='w', pady=5)
        self.trial_sep_combo = ttk.Combobox(sep_frame, state='readonly', width=30)
        self.trial_sep_combo.grid(row=0, column=1, padx=5, pady=5)
        self.trial_sep_combo.bind('<<ComboboxSelected>>', self.update_numcat_values)
        
        ttk.Label(sep_frame, text="Unique values for this separator:").grid(row=1, column=0, sticky='w', pady=5)
        self.numcat_label = ttk.Label(sep_frame, text="", foreground='blue', wraplength=600, justify='left')
        self.numcat_label.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(sep_frame, text="Select Cat value (Entry/Exit):").grid(row=2, column=0, sticky='w', pady=5)
        self.cat_combo = ttk.Combobox(sep_frame, state='readonly', width=30)
        self.cat_combo.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Button(sep_frame, text="Find Trials", command=self.find_trials).grid(row=3, column=0, columnspan=2, pady=10)
        
        # Trial preview
        trial_frame = ttk.LabelFrame(parent, text="Trial Detection Results", padding=10)
        trial_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.trial_text = scrolledtext.ScrolledText(trial_frame, height=15, width=80)
        self.trial_text.pack(fill='both', expand=True)
    
    def create_marker_config_tab(self, parent):
        """Create marker configuration UI with scrollable canvas"""
        # Instructions
        info_frame = ttk.Frame(parent)
        info_frame.pack(fill='x', padx=10, pady=5)
        ttk.Label(info_frame, text="Configure up to 20 markers to track in trials. Leave unused markers unconfigured.", 
                  font=('Arial', 9, 'italic')).pack()
        
        # Create canvas with scrollbar
        canvas_frame = ttk.Frame(parent)
        canvas_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        canvas = tk.Canvas(canvas_frame, height=500)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=canvas.yview)
        self.marker_frame = ttk.Frame(canvas)
        
        self.marker_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.marker_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Create 20 marker rows
        headers = ['Marker', 'State', 'Reward?', 'Reward State']
        for col, header in enumerate(headers):
            ttk.Label(self.marker_frame, text=header, font=('Arial', 9, 'bold')).grid(
                row=0, column=col, padx=5, pady=5, sticky='w')
        
        for i in range(20):
            self.create_marker_row(i)
    
    def create_marker_row(self, index):
        """Create a single marker configuration row"""
        row = index + 1
        
        marker_dict = {
            'label': ttk.Label(self.marker_frame, text=f"Marker {index + 1}"),
            'state_combo': ttk.Combobox(self.marker_frame, state='readonly', width=20),
            'reward_var': tk.BooleanVar(),
            'reward_check': None,
            'reward_combo': ttk.Combobox(self.marker_frame, state='disabled', width=20)
        }
        
        marker_dict['label'].grid(row=row, column=0, padx=5, pady=2, sticky='w')
        marker_dict['state_combo'].grid(row=row, column=1, padx=5, pady=2)
        
        marker_dict['reward_check'] = ttk.Checkbutton(
            self.marker_frame, 
            variable=marker_dict['reward_var'],
            command=lambda idx=index: self.toggle_reward_combo(idx)
        )
        marker_dict['reward_check'].grid(row=row, column=2, padx=5, pady=2)
        marker_dict['reward_combo'].grid(row=row, column=3, padx=5, pady=2)
        
        self.markers.append(marker_dict)
    
    def toggle_reward_combo(self, index):
        """Enable/disable reward combo based on checkbox"""
        if self.markers[index]['reward_var'].get():
            self.markers[index]['reward_combo']['state'] = 'readonly'
        else:
            self.markers[index]['reward_combo']['state'] = 'disabled'
    
    def browse_catalog(self):
        """Browse for catalog Excel file"""
        filepath = filedialog.askopenfilename(
            title="Select Catalog File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filepath:
            self.catalog_path = filepath
            self.catalog_dir = os.path.dirname(filepath)
            self.catalog_label.config(text=filepath, foreground='black')
    
    def load_sheet(self):
        """Load the Excel sheet and populate column dropdowns"""
        if not self.catalog_path:
            messagebox.showerror("Error", "Please select a catalog file first")
            return
        
        try:
            sheet_name = self.sheet_entry.get().strip()
            if not sheet_name:
                messagebox.showerror("Error", "Please enter a sheet name")
                return
            
            self.sheet_name = sheet_name
            self.catalog_df = pd.read_excel(self.catalog_path, sheet_name=sheet_name, engine='openpyxl')
            
            # Populate column dropdowns
            columns = list(self.catalog_df.columns)
            self.filename_col_combo['values'] = columns
            self.exptype_col_combo['values'] = columns
            
            # Set default selections if possible
            if len(columns) > 0:
                self.filename_col_combo.current(0)
            if len(columns) > 5:
                self.exptype_col_combo.current(5)  # Column F is index 5
                # Trigger experiment type update to populate the dropdown
                self.update_experiment_types()
            
            # Show preview
            preview = f"Loaded sheet '{sheet_name}' with {len(self.catalog_df)} rows and {len(columns)} columns\n\n"
            preview += "First 5 rows:\n"
            preview += self.catalog_df.head().to_string()
            self.preview_text.delete('1.0', tk.END)
            self.preview_text.insert('1.0', preview)
            
            messagebox.showinfo("Success", f"Sheet loaded successfully!\n{len(self.catalog_df)} rows found.")
            
            # Automatically load sample CSV
            self.load_sample_csv_manual()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sheet:\n{str(e)}")
    
    def update_experiment_types(self, event=None):
        """Update experiment type filter dropdown based on selected column"""
        if self.catalog_df is None:
            return
        
        try:
            exp_col = self.exptype_col_combo.get()
            if exp_col:
                unique_types = self.catalog_df[exp_col].dropna().unique().tolist()
                self.exptype_filter_combo['values'] = unique_types
                if unique_types:
                    self.exptype_filter_combo.current(0)
                    # Bind event to update suggested CSV when experiment type changes
                    self.exptype_filter_combo.bind('<<ComboboxSelected>>', self.update_suggested_csv)
                    # Update suggested CSV immediately
                    self.update_suggested_csv()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update experiment types:\n{str(e)}")    
    def update_suggested_csv(self, event=None):
        """Update the suggested CSV filename display based on current experiment type selection"""
        if self.catalog_df is None:
            self.suggested_csv_label.config(text="No catalog loaded", foreground='gray')
            self.browse_csv_btn['state'] = 'disabled'
            return
        
        try:
            filename_col = self.filename_col_combo.get()
            exptype_col = self.exptype_col_combo.get()
            exp_filter = self.exptype_filter_combo.get()
            
            if not filename_col or not exptype_col or not exp_filter:
                self.suggested_csv_label.config(text="Please complete file selection first", foreground='gray')
                self.browse_csv_btn['state'] = 'disabled'
                return
            
            # Filter by experiment type
            filtered_df = self.catalog_df[self.catalog_df[exptype_col] == exp_filter]
            
            # Get first CSV filename from filtered catalog
            suggested_file = None
            for idx, row in filtered_df.iterrows():
                fname = str(row[filename_col]).strip()
                if fname and fname.lower() != 'nan':
                    if not fname.endswith('.csv'):
                        fname = fname + '.csv'
                    suggested_file = fname
                    break
            
            if not suggested_file:
                self.suggested_csv_label.config(text="No CSV files found for this experiment type", foreground='red')
                self.browse_csv_btn['state'] = 'normal'
                return
            
            # Check if the suggested file exists
            csv_path = os.path.join(self.catalog_dir, 'data', suggested_file)
            if not os.path.exists(csv_path):
                csv_path = os.path.join(self.catalog_dir, suggested_file)
            
            if os.path.exists(csv_path):
                self.suggested_csv_label.config(text=suggested_file, foreground='green')
                self.browse_csv_btn['state'] = 'disabled'
            else:
                self.suggested_csv_label.config(text=f"{suggested_file} (NOT FOUND)", foreground='red')
                self.browse_csv_btn['state'] = 'normal'
                
        except Exception as e:
            self.suggested_csv_label.config(text="Error determining suggested file", foreground='red')
            self.browse_csv_btn['state'] = 'normal'
    
    def browse_csv_file(self):
        """Browse for CSV file when suggested file is not found"""
        initialdir = self.catalog_dir if self.catalog_dir else None
        filepath = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialdir=initialdir
        )
        
        if filepath:
            try:
                # Load the selected CSV file
                df = pd.read_csv(
                    filepath,
                    skiprows=11,
                    usecols=range(8),
                    encoding='latin-1',
                    header=None,
                    names=['Num_line', 'S', 'MS', 'Cat', 'Num_cat', 'state', 'Display', 'null']
                )
                
                self.sample_csv_df = df
                
                # Update display
                filename = os.path.basename(filepath)
                self.suggested_csv_label.config(text=f"{filename} (manually selected)", foreground='blue')
                self.sample_status_label.config(text="Manual CSV loaded successfully", foreground='green')
                self.browse_csv_btn['state'] = 'disabled'
                
                # Populate dropdowns
                unique_states = sorted(df['state'].dropna().unique().tolist())
                self.trial_sep_combo['values'] = unique_states
                
                # Populate marker state dropdowns
                for marker in self.markers:
                    marker['state_combo']['values'] = unique_states
                    marker['reward_combo']['values'] = unique_states
                
                messagebox.showinfo("Success", f"CSV file loaded successfully!\\n{len(df)} rows found.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV file:\\n{str(e)}")
        
    def load_sample_csv_manual(self):
        """Manual trigger to load sample CSV"""
        # Update the suggested CSV display first
        self.update_suggested_csv()
        
        result = self.load_sample_csv()
        if result is not None:
            self.sample_status_label.config(text="Sample loaded successfully", foreground='green')
    
    def load_sample_csv(self):
        """Load a sample CSV file to populate trial separator dropdown"""
        if self.catalog_df is None:
            if hasattr(self, 'sample_status_label'):
                self.sample_status_label.config(text="Please load catalog first", foreground='red')
            return None
        
        try:
            filename_col = self.filename_col_combo.get()
            if not filename_col:
                messagebox.showerror("Error", "Please select file name column")
                return None
            
            # Get experiment type filter if selected
            exptype_col = self.exptype_col_combo.get()
            exp_filter = self.exptype_filter_combo.get()
            
            # Filter by experiment type if both column and filter are selected
            if exptype_col and exp_filter:
                filtered_df = self.catalog_df[self.catalog_df[exptype_col] == exp_filter]
            else:
                filtered_df = self.catalog_df
            
            # Get first CSV filename from filtered catalog
            first_file = None
            for idx, row in filtered_df.iterrows():
                fname = str(row[filename_col]).strip()
                if fname and fname.lower() != 'nan':
                    if not fname.endswith('.csv'):
                        fname = fname + '.csv'
                    first_file = fname
                    break
            
            if not first_file:
                messagebox.showerror("Error", "No valid CSV filenames found in catalog")
                if hasattr(self, 'browse_csv_btn'):
                    self.browse_csv_btn['state'] = 'normal'
                return None
            
            # Try to load the CSV file
            csv_path = os.path.join(self.catalog_dir, 'data', first_file)
            if not os.path.exists(csv_path):
                csv_path = os.path.join(self.catalog_dir, first_file)
            
            if not os.path.exists(csv_path):
                if hasattr(self, 'sample_status_label'):
                    self.sample_status_label.config(text=f"CSV not found: {first_file}", foreground='red')
                if hasattr(self, 'browse_csv_btn'):
                    self.browse_csv_btn['state'] = 'normal'
                else:
                    messagebox.showerror("Error", f"Could not find sample CSV file:\n{first_file}")
                return None
            
            # Load CSV with proper encoding and structure
            df = pd.read_csv(
                csv_path,
                skiprows=11,
                usecols=range(8),
                encoding='latin-1',
                header=None,
                names=['Num_line', 'S', 'MS', 'Cat', 'Num_cat', 'state', 'Display', 'null']
            )
            
            self.sample_csv_df = df
            
            # Populate trial separator dropdown with unique 'state' values
            unique_states = sorted(df['state'].dropna().unique().tolist())
            self.trial_sep_combo['values'] = unique_states
            
            # Also get unique Cat values (Entry, Exit, etc.)
            unique_cats = sorted(df['Cat'].dropna().unique().tolist())
            
            # Populate marker state dropdowns
            for marker in self.markers:
                marker['state_combo']['values'] = unique_states
                marker['reward_combo']['values'] = unique_states
            
            if hasattr(self, 'sample_status_label'):
                self.sample_status_label.config(
                    text=f"Loaded: {len(df)} rows, {len(unique_states)} states, {len(unique_cats)} categories",
                    foreground='green'
                )
            
            # Disable browse button since we successfully loaded a CSV
            if hasattr(self, 'browse_csv_btn'):
                self.browse_csv_btn['state'] = 'disabled'
            
            return df
            
        except Exception as e:
            if hasattr(self, 'sample_status_label'):
                self.sample_status_label.config(text=f"Error loading sample", foreground='red')
            if hasattr(self, 'browse_csv_btn'):
                self.browse_csv_btn['state'] = 'normal'
            messagebox.showerror("Error", f"Failed to load sample CSV:\\n{str(e)}\\n\\n{traceback.format_exc()}")
            return None
    
    def update_numcat_values(self, event=None):
        """Update Num_cat values display based on selected separator"""
        if self.sample_csv_df is None:
            self.load_sample_csv()
        
        if self.sample_csv_df is None:
            return
        
        separator = self.trial_sep_combo.get()
        if separator:
            # Find rows where state == separator
            sep_rows = self.sample_csv_df[self.sample_csv_df['state'] == separator]
            
            # Get unique Cat values (Entry, Exit, etc.) - these are in column D (Cat column)
            unique_cat = sep_rows['Cat'].dropna().unique().tolist()
            unique_cat = sorted(set(str(v) for v in unique_cat))
            
            # Get Num_cat values for display
            unique_numcat = sorted(sep_rows['Num_cat'].dropna().unique().tolist(), key=str)
            
            display_text = f"Cat (Column D): {', '.join(unique_cat)}  |  Num_cat: {', '.join(map(str, unique_numcat))}"
            self.numcat_label.config(text=display_text)
            
            # Populate the Cat selection dropdown with Entry, Exit, etc.
            # Add "Both" as an option if there are multiple values
            cat_options = unique_cat.copy()
            if len(unique_cat) > 1:
                cat_options.append("Both")
            
            self.cat_combo['values'] = cat_options
            if cat_options:
                # Auto-select first value
                self.cat_combo.current(0)
    
    def find_trials(self):
        """Find and display trial information"""
        if self.sample_csv_df is None:
            self.load_sample_csv()
        
        if self.sample_csv_df is None:
            return
        
        separator = self.trial_sep_combo.get()
        if not separator:
            messagebox.showerror("Error", "Please select a trial separator")
            return
        
        cat_value = self.cat_combo.get()
        if not cat_value:
            messagebox.showerror("Error", "Please select a Cat value (Entry/Exit)")
            return
        
        try:
            df = self.sample_csv_df
            
            # Find trial start positions
            # state == separator (column F) AND Cat == selected value (column D)
            if cat_value == "Both":
                # If "Both" selected, find all rows with the separator regardless of Cat
                trial_starts = df[df['state'] == separator].index.tolist()
            else:
                # Find rows where state == separator AND Cat == selected value (Entry or Exit)
                trial_starts = df[(df['state'] == separator) & (df['Cat'] == cat_value)].index.tolist()
            
            # Exclude trials that contain 'Finish' in Cat column
            valid_trials = []
            for i, start_idx in enumerate(trial_starts):
                # Determine trial end
                if i < len(trial_starts) - 1:
                    end_idx = trial_starts[i + 1]
                else:
                    end_idx = len(df)
                
                # Check if this trial contains 'Finish'
                trial_segment = df.iloc[start_idx:end_idx]
                if 'Finish' not in trial_segment['Cat'].values:
                    valid_trials.append((start_idx, end_idx))
            
            # Display results
            result = f"Trial Separator: {separator}\n"
            result += f"Cat Value: {cat_value}\n"
            result += f"Total trial markers found: {len(trial_starts)}\n"
            result += f"Valid trials (excluding incomplete): {len(valid_trials)}\n\n"
            result += "First 5 valid trials:\n"
            
            for i, (start, end) in enumerate(valid_trials[:5]):
                result += f"Trial {i+1}: Line {start} to Line {end-1} ({end-start} lines)\n"
            
            self.trial_text.delete('1.0', tk.END)
            self.trial_text.insert('1.0', result)
            
            # Enable execute button
            self.execute_btn['state'] = 'normal'
            
            messagebox.showinfo("Trial Detection", f"Found {len(valid_trials)} valid trials!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to find trials:\n{str(e)}\n\n{traceback.format_exc()}")
    
    def execute_extraction(self):
        """Execute the data extraction process"""
        try:
            # Validate configuration
            if not self.validate_config():
                return
            
            # Get file list from catalog
            filename_col = self.filename_col_combo.get()
            exptype_col = self.exptype_col_combo.get()
            exp_filter = self.exptype_filter_combo.get()
            
            # Filter files by experiment type
            filtered_df = self.catalog_df[self.catalog_df[exptype_col] == exp_filter]
            file_list = []
            for idx, row in filtered_df.iterrows():
                fname = str(row[filename_col]).strip()
                if fname and fname.lower() != 'nan':
                    if not fname.endswith('.csv'):
                        fname = fname + '.csv'
                    file_list.append(fname)
            
            if not file_list:
                messagebox.showerror("Error", "No files found matching the selected experiment type")
                return
            
            # Create output directory
            output_dir = os.path.join(self.catalog_dir, 'processed_data')
            os.makedirs(output_dir, exist_ok=True)
            
            # Prepare aggregation data
            agg_data = []
            
            # Process each file
            total_files = len(file_list)
            for idx, filename in enumerate(file_list, 1):
                self.progress_label.config(text=f"Crunching {idx} out of {total_files}")
                self.root.update()
                
                result = self.process_file(filename, output_dir)
                agg_data.append(result)
            
            # Create aggregated file
            self.create_aggregated_file(agg_data, output_dir, exp_filter)
            
            self.progress_label.config(text=f"Complete! Processed {total_files} files.")
            messagebox.showinfo("Success", f"Data extraction complete!\n\nProcessed {total_files} files.\nOutput location: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Extraction failed:\n{str(e)}\n\n{traceback.format_exc()}")
            self.progress_label.config(text="Error occurred")
    
    def validate_config(self):
        """Validate that all necessary configuration is complete"""
        if self.catalog_df is None:
            messagebox.showerror("Error", "Please load catalog sheet first")
            return False
        
        if not self.filename_col_combo.get():
            messagebox.showerror("Error", "Please select file name column")
            return False
        
        if not self.exptype_col_combo.get():
            messagebox.showerror("Error", "Please select experiment type column")
            return False
        
        if not self.exptype_filter_combo.get():
            messagebox.showerror("Error", "Please select experiment type filter")
            return False
        
        if not self.trial_sep_combo.get():
            messagebox.showerror("Error", "Please select trial separator")
            return False
        
        if not self.cat_combo.get():
            messagebox.showerror("Error", "Please select Cat value (Entry/Exit)")
            return False
        
        # Check at least one marker is configured
        configured_markers = [m for m in self.markers if m['state_combo'].get()]
        if not configured_markers:
            messagebox.showerror("Error", "Please configure at least one marker")
            return False
        
        return True
    
    def process_file(self, filename, output_dir):
        """Process a single CSV file"""
        try:
            # Try to find the CSV file
            csv_path = os.path.join(self.catalog_dir, 'data', filename)
            if not os.path.exists(csv_path):
                csv_path = os.path.join(self.catalog_dir, filename)
            
            if not os.path.exists(csv_path):
                # File not found - return empty result with "not present"
                result = {'filename': filename.replace('.csv', '.xlsx'), 'status': 'not present'}
                for i, marker in enumerate(self.markers):
                    if marker['state_combo'].get():
                        marker_name = marker['state_combo'].get()
                        result[f'{marker_name}_sum'] = 'not present'
                        result[f'{marker_name}_avg_time'] = 'not present'
                        if marker['reward_var'].get() and marker['reward_combo'].get():
                            reward_name = marker['reward_combo'].get()
                            result[f'{marker_name}_reward_{reward_name}_sum'] = 'not present'
                            result[f'{marker_name}_reward_{reward_name}_avg_time'] = 'not present'
                return result
            
            # Read header (first 11 rows)
            with open(csv_path, 'r', encoding='latin-1') as f:
                header_lines = [f.readline() for _ in range(11)]
            
            # Read CSV data
            df = pd.read_csv(
                csv_path,
                skiprows=11,
                usecols=range(8),
                encoding='latin-1',
                header=None,
                names=['Num_line', 'S', 'MS', 'Cat', 'Num_cat', 'state', 'Display', 'null']
            )
            
            # Extract trials
            trials_df = self.extract_trials(df)
            
            # Create Excel output
            output_filename = filename.replace('.csv', '.xlsx')
            output_path = os.path.join(output_dir, output_filename)
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # Sheet 1: raw
                df.to_excel(writer, sheet_name='raw', index=False)
                
                # Sheet 2: trial
                trials_df.to_excel(writer, sheet_name='trial', index=False)
                
                # Sheet 3: header
                header_df = pd.DataFrame({'Header': [line.strip() for line in header_lines]})
                header_df.to_excel(writer, sheet_name='header', index=False)
            
            # Calculate aggregation data
            result = {'filename': output_filename, 'status': 'processed'}
            for i, marker in enumerate(self.markers):
                if marker['state_combo'].get():
                    marker_name = marker['state_combo'].get()
                    
                    # Use actual marker name in column headers
                    presence_col = f'{marker_name}_present'
                    time_col = f'{marker_name}_time_ms'
                    if presence_col in trials_df.columns:
                        result[f'{marker_name}_sum'] = trials_df[presence_col].sum()
                        present_times = trials_df[trials_df[presence_col] == 1][time_col]
                        result[f'{marker_name}_avg_time'] = present_times.mean() if len(present_times) > 0 else 0
                    else:
                        result[f'{marker_name}_sum'] = 0
                        result[f'{marker_name}_avg_time'] = 0
                    
                    # Reward marker if applicable
                    if marker['reward_var'].get() and marker['reward_combo'].get():
                        reward_name = marker['reward_combo'].get()
                        reward_presence_col = f'{marker_name}_reward_{reward_name}_present'
                        reward_time_col = f'{marker_name}_reward_{reward_name}_time_ms'
                        if reward_presence_col in trials_df.columns:
                            result[f'{marker_name}_reward_{reward_name}_sum'] = trials_df[reward_presence_col].sum()
                            present_reward_times = trials_df[trials_df[reward_presence_col] == 1][reward_time_col]
                            result[f'{marker_name}_reward_{reward_name}_avg_time'] = present_reward_times.mean() if len(present_reward_times) > 0 else 0
                        else:
                            result[f'{marker_name}_reward_{reward_name}_sum'] = 0
                            result[f'{marker_name}_reward_{reward_name}_avg_time'] = 0
            
            return result
            
        except Exception as e:
            # Error processing file - return error status
            result = {'filename': filename.replace('.csv', '.xlsx'), 'status': f'error: {str(e)}'}
            return result
    
    def extract_trials(self, df):
        """Extract trial data from dataframe"""
        separator = self.trial_sep_combo.get()
        cat_value = self.cat_combo.get()
        
        # Find trial starts
        # state == separator (column F) AND Cat == selected value (column D)
        if cat_value == "Both":
            # If "Both" selected, find all rows with the separator regardless of Cat
            trial_starts = df[df['state'] == separator].index.tolist()
        else:
            # Find rows where state == separator AND Cat == selected value (Entry or Exit)
            trial_starts = df[(df['state'] == separator) & (df['Cat'] == cat_value)].index.tolist()
        
        trial_data = []
        
        for trial_num, start_idx in enumerate(trial_starts, 1):
            # Determine trial end
            if trial_num < len(trial_starts):
                end_idx = trial_starts[trial_num]
            else:
                end_idx = len(df)
            
            # Check if trial contains 'Finish' - if so, skip
            trial_segment = df.iloc[start_idx:end_idx]
            if 'Finish' in trial_segment['Cat'].values:
                continue
            
            # Calculate trial start time in MS
            start_row = df.iloc[start_idx]
            trial_start_time = start_row['S'] * 1000 + start_row['MS']
            
            # Initialize trial record
            trial_record = {
                'trial_num': trial_num,
                'start_line': start_idx,
                'stop_line': end_idx - 1,
                'trial_start_time_ms': trial_start_time
            }
            
            # Check each configured marker
            for i, marker in enumerate(self.markers):
                marker_state = marker['state_combo'].get()
                if not marker_state:
                    continue
                
                # Use actual marker name in column headers
                # Find marker in trial segment
                marker_rows = trial_segment[trial_segment['state'] == marker_state]
                if len(marker_rows) > 0:
                    # Marker present
                    marker_row = marker_rows.iloc[0]
                    marker_time = marker_row['S'] * 1000 + marker_row['MS']
                    relative_time = marker_time - trial_start_time
                    
                    trial_record[f'{marker_state}_present'] = 1
                    trial_record[f'{marker_state}_time_ms'] = relative_time
                else:
                    # Marker not present
                    trial_record[f'{marker_state}_present'] = 0
                    trial_record[f'{marker_state}_time_ms'] = 0
                
                # Check reward marker if applicable
                if marker['reward_var'].get() and marker['reward_combo'].get():
                    reward_state = marker['reward_combo'].get()
                    reward_rows = trial_segment[trial_segment['state'] == reward_state]
                    if len(reward_rows) > 0:
                        reward_row = reward_rows.iloc[0]
                        reward_time = reward_row['S'] * 1000 + reward_row['MS']
                        reward_relative_time = reward_time - trial_start_time
                        
                        trial_record[f'{marker_state}_reward_{reward_state}_present'] = 1
                        trial_record[f'{marker_state}_reward_{reward_state}_time_ms'] = reward_relative_time
                    else:
                        trial_record[f'{marker_state}_reward_{reward_state}_present'] = 0
                        trial_record[f'{marker_state}_reward_{reward_state}_time_ms'] = 0
            
            trial_data.append(trial_record)
        
        return pd.DataFrame(trial_data)
    
    def create_aggregated_file(self, agg_data, output_dir, exp_type):
        """Create aggregated Excel file with summary statistics"""
        catalog_name = Path(self.catalog_path).stem
        agg_filename = f"{catalog_name}_{self.sheet_name}_{exp_type}.xlsx"
        agg_path = os.path.join(output_dir, agg_filename)
        
        agg_df = pd.DataFrame(agg_data)
        
        # Reorder columns for better readability
        cols = ['filename', 'status']
        for col in agg_df.columns:
            if col not in cols:
                cols.append(col)
        
        agg_df = agg_df[cols]
        
        # TODO #2: Create parameters sheet with GUI configuration
        params_data = {
            'Parameter': [
                'Catalog File',
                'Sheet Name',
                'Filename Column',
                'Experiment Type Column',
                'Experiment Type Filter',
                'Trial Separator (state)',
                'Cat Value',
                '---Markers Configuration---',
            ],
            'Value': [
                Path(self.catalog_path).name,
                self.sheet_name,
                self.filename_col_combo.get(),
                self.exptype_col_combo.get(),
                self.exptype_filter_combo.get(),
                self.trial_sep_combo.get(),
                self.cat_combo.get(),
                '',
            ]
        }
        
        # Add marker configurations
        for i, marker in enumerate(self.markers, 1):
            if marker['state_combo'].get():
                params_data['Parameter'].append(f'Marker {i}')
                params_data['Value'].append(marker['state_combo'].get())
                
                if marker['reward_var'].get() and marker['reward_combo'].get():
                    params_data['Parameter'].append(f'  - Reward for Marker {i}')
                    params_data['Value'].append(marker['reward_combo'].get())
        
        params_df = pd.DataFrame(params_data)
        
        # Write both sheets to Excel
        with pd.ExcelWriter(agg_path, engine='openpyxl') as writer:
            agg_df.to_excel(writer, sheet_name='aggregated_data', index=False)
            params_df.to_excel(writer, sheet_name='parameters', index=False)


def main():
    root = tk.Tk()
    app = CSVTrialExtractor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
