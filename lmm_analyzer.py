"""
Linear Mixed Model (LMM) Analysis Tool
Performs statistical analysis on behavioral experiment data with repeated measures
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages
import scipy.stats as stats
import os


class LMMAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("Linear Mixed Model Analyzer")
        self.root.geometry("1200x900")
        
        self.data_df = None
        self.data_path = None
        self.analysis_results = []
        self.figures_for_export = []  # Store figures for PDF export
        
        self.create_gui()
    
    def create_gui(self):
        """Create the main GUI layout"""
        # Create main container with scrollbar
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True)
        
        # Top section: File loading and variable selection
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill='x', padx=10, pady=10)
        
        self.create_file_section(top_frame)
        self.create_variable_section(top_frame)
        
        # Bottom section: Results display
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.create_results_section(bottom_frame)
    
    def create_file_section(self, parent):
        """Create file selection section"""
        file_frame = ttk.LabelFrame(parent, text="Data File", padding=10)
        file_frame.pack(fill='x', pady=5)
        
        self.file_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.file_label.pack(side='left', fill='x', expand=True)
        
        ttk.Button(file_frame, text="Browse Excel File...", 
                  command=self.browse_file).pack(side='right')
    
    def create_variable_section(self, parent):
        """Create variable selection section"""
        var_frame = ttk.LabelFrame(parent, text="Variable Selection", padding=10)
        var_frame.pack(fill='x', pady=5)
        
        # Animal ID selection
        id_frame = ttk.Frame(var_frame)
        id_frame.pack(fill='x', pady=5)
        ttk.Label(id_frame, text="Animal ID Column:", width=20).pack(side='left')
        self.id_combo = ttk.Combobox(id_frame, state='readonly', width=30)
        self.id_combo.pack(side='left', padx=5)
        
        # Independent Variables (VI) - 3 dropdowns
        vi_frame = ttk.LabelFrame(var_frame, text="Independent Variables (up to 3)", padding=5)
        vi_frame.pack(fill='x', pady=5)
        
        self.vi_combos = []
        for i in range(3):
            frame = ttk.Frame(vi_frame)
            frame.pack(fill='x', pady=2)
            ttk.Label(frame, text=f"VI {i+1}:", width=10).pack(side='left')
            combo = ttk.Combobox(frame, state='readonly', width=30)
            combo.pack(side='left', padx=5)
            self.vi_combos.append(combo)
        
        # Dependent Variables (VD) - 6 dropdowns
        vd_frame = ttk.LabelFrame(var_frame, text="Dependent Variables (up to 6)", padding=5)
        vd_frame.pack(fill='x', pady=5)
        
        self.vd_combos = []
        for i in range(6):
            frame = ttk.Frame(vd_frame)
            frame.pack(fill='x', pady=2)
            ttk.Label(frame, text=f"VD {i+1}:", width=10).pack(side='left')
            combo = ttk.Combobox(frame, state='readonly', width=30)
            combo.pack(side='left', padx=5)
            self.vd_combos.append(combo)
        
        # Analyze button
        button_frame = ttk.Frame(var_frame)
        button_frame.pack(fill='x', pady=10)
        
        self.export_btn = ttk.Button(button_frame, text="Export to PDF", 
                                     command=self.export_to_pdf, state='disabled')
        self.export_btn.pack(side='right', padx=5)
        
        self.analyze_btn = ttk.Button(button_frame, text="Analyze This", 
                                      command=self.run_analysis, state='disabled')
        self.analyze_btn.pack(side='right')
    
    def create_results_section(self, parent):
        """Create results display section"""
        results_frame = ttk.LabelFrame(parent, text="Analysis Results", padding=10)
        results_frame.pack(fill='both', expand=True)
        
        # Create canvas with scrollbar for results
        canvas = tk.Canvas(results_frame)
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=canvas.yview)
        self.results_frame = ttk.Frame(canvas)
        
        self.results_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.results_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Initial message
        self.status_label = ttk.Label(self.results_frame, 
                                      text="Load data and select variables to begin analysis",
                                      font=('Arial', 10, 'italic'))
        self.status_label.pack(pady=20)
    
    def browse_file(self):
        """Browse for Excel file"""
        filepath = filedialog.askopenfilename(
            title="Select Excel Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if filepath:
            try:
                self.data_path = filepath
                self.data_df = pd.read_excel(filepath, engine='openpyxl')
                
                self.file_label.config(text=filepath, foreground='black')
                
                # Populate column dropdowns
                columns = [''] + list(self.data_df.columns)
                
                self.id_combo['values'] = columns
                for combo in self.vi_combos:
                    combo['values'] = columns
                for combo in self.vd_combos:
                    combo['values'] = columns
                
                # Enable analyze button
                self.analyze_btn['state'] = 'normal'
                
                messagebox.showinfo("Success", 
                                  f"Data loaded successfully!\n\n{len(self.data_df)} rows\n{len(self.data_df.columns)} columns")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
    
    def run_analysis(self):
        """Run LMM analysis"""
        if self.data_df is None:
            messagebox.showerror("Error", "Please load data file first")
            return
        
        # Get selected variables
        animal_id = self.id_combo.get()
        if not animal_id:
            messagebox.showerror("Error", "Please select Animal ID column")
            return
        
        # Get independent variables (only non-empty selections)
        vi_vars = [combo.get() for combo in self.vi_combos if combo.get()]
        if not vi_vars:
            messagebox.showerror("Error", "Please select at least one Independent Variable")
            return
        
        # Get dependent variables (only non-empty selections)
        vd_vars = [combo.get() for combo in self.vd_combos if combo.get()]
        if not vd_vars:
            messagebox.showerror("Error", "Please select at least one Dependent Variable")
            return
        
        # Clear previous results
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        
        # Clear stored figures
        self.figures_for_export = []
        self.analysis_results = []
        
        self.status_label = ttk.Label(self.results_frame, 
                                      text="Running analysis...", 
                                      font=('Arial', 12, 'bold'))
        self.status_label.pack(pady=10)
        self.root.update()
        
        # Add relative day column if not present
        self.add_relative_day_column(animal_id)
        
        # Run analysis for each dependent variable
        for vd in vd_vars:
            self.analyze_variable(animal_id, vi_vars, vd)
        
        self.status_label.config(text="Analysis Complete!")
        
        # Enable export button
        self.export_btn['state'] = 'normal'
    
    def add_relative_day_column(self, animal_id):
        """Add relative day column for temporal ordering"""
        if 'Relative_Day' not in self.data_df.columns:
            # Check if we have date columns
            date_cols = ['Year', 'Month', 'Day']
            if all(col in self.data_df.columns for col in date_cols):
                try:
                    # Create date column
                    self.data_df['Date'] = pd.to_datetime(
                        self.data_df[['Year', 'Month', 'Day']].astype(str).agg('-'.join, axis=1)
                    )
                    
                    # Calculate relative day for each animal
                    relative_days = []
                    for animal in self.data_df[animal_id].unique():
                        animal_data = self.data_df[self.data_df[animal_id] == animal].sort_values('Date')
                        if len(animal_data) > 0:
                            first_date = animal_data['Date'].min()
                            for idx, row in animal_data.iterrows():
                                days = (row['Date'] - first_date).days
                                relative_days.append((idx, days))
                    
                    # Add relative day column
                    for idx, days in relative_days:
                        self.data_df.at[idx, 'Relative_Day'] = days
                        
                except Exception as e:
                    print(f"Could not create Relative_Day column: {e}")
    
    def analyze_variable(self, animal_id, vi_vars, vd_var):
        """Analyze one dependent variable"""
        # Store analysis info for export
        analysis_info = {
            'animal_id': animal_id,
            'vi_vars': vi_vars,
            'vd_var': vd_var,
            'qq_figure': None,
            'viz_figure': None,
            'stats_text': '',
            'qq_assessment': ''
        }
        
        # Create frame for this analysis
        analysis_frame = ttk.LabelFrame(self.results_frame, 
                                       text=f"Analysis: {vd_var}", 
                                       padding=10)
        analysis_frame.pack(fill='x', padx=10, pady=10)
        
        # Display variable information
        info_text = f"Animal ID: {animal_id}\n"
        info_text += f"Independent Variables: {', '.join(vi_vars)}\n"
        info_text += f"Dependent Variable: {vd_var}\n"
        
        ttk.Label(analysis_frame, text=info_text, font=('Arial', 9)).pack(anchor='w', pady=5)
        
        # Prepare data - remove NaN values
        analysis_data = self.data_df[[animal_id] + vi_vars + [vd_var]].dropna()
        
        if len(analysis_data) < 10:
            ttk.Label(analysis_frame, text="Insufficient data for analysis (< 10 observations)", 
                     foreground='red').pack(pady=5)
            return
        
        # Create QQ plot
        qq_fig, qq_assessment = self.create_qq_plot(analysis_frame, analysis_data[vd_var], vd_var)
        analysis_info['qq_figure'] = qq_fig
        analysis_info['qq_assessment'] = qq_assessment
        
        # Perform LMM analysis (simplified version using statsmodels if available, otherwise ANOVA)
        stats_text = self.perform_statistical_test(analysis_frame, analysis_data, animal_id, vi_vars, vd_var)
        analysis_info['stats_text'] = stats_text
        
        # Create visualization
        viz_fig = self.create_visualization(analysis_frame, analysis_data, vi_vars, vd_var)
        analysis_info['viz_figure'] = viz_fig
        
        # Store for export
        self.analysis_results.append(analysis_info)
    
    def create_qq_plot(self, parent, data, var_name):
        """Create QQ plot for normality assessment"""
        fig = Figure(figsize=(6, 4))
        ax = fig.add_subplot(111)
        
        # QQ plot
        stats.probplot(data, dist="norm", plot=ax)
        ax.set_title(f'Q-Q Plot: {var_name}')
        ax.grid(True, alpha=0.3)
        
        # Calculate normality test
        _, p_value = stats.shapiro(data)
        
        # Determine assessment
        if p_value > 0.05:
            assessment = "PASS"
            color = 'green'
            message = f"Data appears normally distributed (p={p_value:.4f})"
        elif p_value > 0.01:
            assessment = "WEAK"
            color = 'orange'
            message = f"Borderline normality (p={p_value:.4f})"
        else:
            assessment = "FAIL"
            color = 'red'
            message = f"Data not normally distributed (p={p_value:.4f})"
        
        # Display plot
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=5)
        
        # Display assessment
        assessment_label = ttk.Label(parent, text=f"{assessment}: {message}", 
                                    foreground=color, font=('Arial', 10, 'bold'))
        assessment_label.pack(pady=5)
        
        return fig, f"{assessment}: {message}"
    
    def perform_statistical_test(self, parent, data, animal_id, vi_vars, vd_var):
        """Perform statistical analysis"""
        results_frame = ttk.Frame(parent)
        results_frame.pack(fill='x', pady=10)
        
        ttk.Label(results_frame, text="Statistical Results:", 
                 font=('Arial', 11, 'bold')).pack(anchor='w', pady=5)
        
        stats_text = ""
        
        try:
            # Try to import statsmodels for proper LMM
            try:
                import statsmodels.api as sm
                import statsmodels.formula.api as smf
                has_statsmodels = True
            except ImportError:
                has_statsmodels = False
            
            if has_statsmodels:
                # Build formula for LMM
                formula = f"{vd_var} ~ {' + '.join(vi_vars)}"
                
                try:
                    # Fit mixed linear model
                    model = smf.mixedlm(formula, data, groups=data[animal_id])
                    result = model.fit()
                    
                    # Display results
                    result_text = scrolledtext.ScrolledText(results_frame, height=10, width=80)
                    result_text.pack(pady=5)
                    
                    summary = str(result.summary())
                    result_text.insert('1.0', summary)
                    
                    stats_text = summary
                    
                    # Highlight significant results
                    p_values = result.pvalues
                    sig_vars = [var for var, p in p_values.items() if p < 0.05 and var != 'Intercept']
                    
                    if sig_vars:
                        sig_text = "Significant effects (p < 0.05): " + ", ".join(sig_vars)
                        sig_label = ttk.Label(results_frame, text=sig_text, 
                                             foreground='red', font=('Arial', 10, 'bold'))
                        sig_label.pack(pady=5)
                        stats_text += f"\n\n{sig_text}"
                    
                except Exception as e:
                    ttk.Label(results_frame, text=f"LMM failed: {str(e)}", 
                             foreground='orange').pack(pady=5)
                    stats_text = self.perform_anova(results_frame, data, vi_vars, vd_var)
            else:
                # Fallback to ANOVA if statsmodels not available
                ttk.Label(results_frame, 
                         text="Note: statsmodels not installed. Using simplified ANOVA.\nFor full LMM: pip install statsmodels",
                         foreground='blue', font=('Arial', 9, 'italic')).pack(pady=5)
                stats_text = self.perform_anova(results_frame, data, vi_vars, vd_var)
                
        except Exception as e:
            error_msg = f"Analysis error: {str(e)}"
            ttk.Label(results_frame, text=error_msg, 
                     foreground='red').pack(pady=5)
            stats_text = error_msg
        
        return stats_text
    
    def perform_anova(self, parent, data, vi_vars, vd_var):
        """Perform ANOVA as fallback"""
        result_text = scrolledtext.ScrolledText(parent, height=8, width=80)
        result_text.pack(pady=5)
        
        result_text.insert('1.0', "ANOVA Results (Simplified Analysis):\n\n")
        
        anova_text = "ANOVA Results (Simplified Analysis):\n\n"
        
        # Perform one-way ANOVA for each independent variable
        for vi in vi_vars:
            groups = []
            for level in data[vi].unique():
                group_data = data[data[vi] == level][vd_var]
                if len(group_data) > 0:
                    groups.append(group_data)
            
            if len(groups) >= 2:
                f_stat, p_value = stats.f_oneway(*groups)
                
                sig_marker = " ***SIGNIFICANT***" if p_value < 0.05 else ""
                line = f"{vi}: F={f_stat:.4f}, p={p_value:.4f}{sig_marker}\n"
                result_text.insert(tk.END, line)
                anova_text += line
        
        return anova_text
    
    def create_visualization(self, parent, data, vi_vars, vd_var):
        """Create visualization of the data"""
        fig = Figure(figsize=(10, 4))
        
        # Create subplots for each independent variable
        n_plots = len(vi_vars)
        
        for i, vi in enumerate(vi_vars, 1):
            ax = fig.add_subplot(1, n_plots, i)
            
            # Box plot
            unique_levels = sorted(data[vi].unique())
            plot_data = [data[data[vi] == level][vd_var].values for level in unique_levels]
            
            ax.boxplot(plot_data, labels=[str(level) for level in unique_levels])
            ax.set_xlabel(vi)
            ax.set_ylabel(vd_var)
            ax.set_title(f'{vd_var} by {vi}')
            ax.grid(True, alpha=0.3)
        
        fig.tight_layout()
        
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)
        
        return fig
    
    def export_to_pdf(self):
        """Export analysis results to PDF"""
        if not self.analysis_results:
            messagebox.showerror("Error", "No analysis results to export")
            return
        
        try:
            # Generate PDF filename from data filename
            if self.data_path:
                base_name = Path(self.data_path).stem
                default_filename = f"{base_name}_LMM.pdf"
                output_dir = Path(self.data_path).parent
            else:
                default_filename = "LMM_Analysis.pdf"
                output_dir = Path.home()
            
            # Ask user where to save
            pdf_path = filedialog.asksaveasfilename(
                title="Save Analysis Report",
                defaultextension=".pdf",
                initialfile=default_filename,
                initialdir=output_dir,
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            
            if not pdf_path:
                return
            
            # Create PDF
            with PdfPages(pdf_path) as pdf:
                # Title page
                fig = plt.figure(figsize=(8.5, 11))
                fig.text(0.5, 0.7, "Linear Mixed Model Analysis Report", 
                        ha='center', va='center', fontsize=20, weight='bold')
                fig.text(0.5, 0.6, f"Data File: {Path(self.data_path).name if self.data_path else 'Unknown'}", 
                        ha='center', va='center', fontsize=12)
                fig.text(0.5, 0.5, f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}", 
                        ha='center', va='center', fontsize=10)
                fig.text(0.5, 0.3, f"Total Analyses: {len(self.analysis_results)}", 
                        ha='center', va='center', fontsize=12)
                pdf.savefig(fig, bbox_inches='tight')
                plt.close(fig)
                
                # Each analysis
                for i, result in enumerate(self.analysis_results, 1):
                    # Analysis info page
                    fig = plt.figure(figsize=(8.5, 11))
                    fig.text(0.5, 0.95, f"Analysis {i}: {result['vd_var']}", 
                            ha='center', va='top', fontsize=16, weight='bold')
                    
                    info_text = f"Animal ID: {result['animal_id']}\n"
                    info_text += f"Independent Variables: {', '.join(result['vi_vars'])}\n"
                    info_text += f"Dependent Variable: {result['vd_var']}\n"
                    
                    fig.text(0.1, 0.85, info_text, ha='left', va='top', fontsize=10, family='monospace')
                    
                    # QQ plot
                    if result['qq_figure']:
                        ax = fig.add_subplot(3, 1, 2)
                        # Re-create QQ plot on this figure
                        ax.text(0.5, 0.5, "Q-Q Plot", ha='center', va='center', fontsize=12, weight='bold')
                        
                    fig.text(0.1, 0.45, f"Normality Assessment: {result['qq_assessment']}", 
                            ha='left', va='top', fontsize=9)
                    
                    # Statistical results
                    stats_display = result['stats_text'][:1000] if len(result['stats_text']) > 1000 else result['stats_text']
                    fig.text(0.1, 0.35, "Statistical Results:", ha='left', va='top', fontsize=11, weight='bold')
                    fig.text(0.1, 0.32, stats_display, ha='left', va='top', fontsize=7, family='monospace', wrap=True)
                    
                    pdf.savefig(fig, bbox_inches='tight')
                    plt.close(fig)
                    
                    # QQ Plot page
                    if result['qq_figure']:
                        pdf.savefig(result['qq_figure'], bbox_inches='tight')
                    
                    # Visualization page
                    if result['viz_figure']:
                        pdf.savefig(result['viz_figure'], bbox_inches='tight')
            
            messagebox.showinfo("Success", f"Analysis report exported to:\n{pdf_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF:\n{str(e)}")


def main():
    root = tk.Tk()
    app = LMMAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
