"""
Generalized Linear Mixed Model (GLMM) Analysis Tool

Performs statistical analysis on behavioral experiment data with repeated measures.
Automatically selects the appropriate model based on the dependent variable suffix:
  - GLMM Negative Binomial (log-link)  for variables ending in '_sum'
  - GLMM Gamma (log-link)              for variables ending in '_avg_time'
  - LMM Gaussian (identity-link)       for all other variables

Model fitting priority:
  1. pymer4 (true GLMM via R/lme4) — if installed
  2. statsmodels GEE              — quasi-GLMM fallback
  3. One-way ANOVA                — last resort
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import numpy as np
from pathlib import Path
import re

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.backends.backend_pdf import PdfPages
import scipy.stats as stats


class GLMMAnalyzer:
    def __init__(self, root):
        self.root = root
        self.root.title("GLMM Analyzer")
        self.root.geometry("1200x900")

        self.data_df = None
        self.data_path = None
        self.analysis_results = []
        self.figures_for_export = []

        self.create_gui()

    # ─────────────────────────────── GUI ────────────────────────────────────── #

    def create_gui(self):
        """Create the main GUI layout."""
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill='both', expand=True)

        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill='x', padx=10, pady=10)

        self.create_file_section(top_frame)
        self.create_variable_section(top_frame)

        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.create_results_section(bottom_frame)

    def create_file_section(self, parent):
        """Create file selection section."""
        file_frame = ttk.LabelFrame(parent, text="Data File", padding=10)
        file_frame.pack(fill='x', pady=5)

        self.file_label = ttk.Label(file_frame, text="No file selected", foreground='gray')
        self.file_label.pack(side='left', fill='x', expand=True)

        ttk.Button(file_frame, text="Browse Excel File...",
                   command=self.browse_file).pack(side='right')

    def create_variable_section(self, parent):
        """Create variable selection section."""
        var_frame = ttk.LabelFrame(parent, text="Variable Selection", padding=10)
        var_frame.pack(fill='x', pady=5)

        # Animal ID
        id_frame = ttk.Frame(var_frame)
        id_frame.pack(fill='x', pady=5)
        ttk.Label(id_frame, text="Animal ID Column:", width=20).pack(side='left')
        self.id_combo = ttk.Combobox(id_frame, state='readonly', width=30)
        self.id_combo.pack(side='left', padx=5)

        # Independent Variables
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

        # Dependent Variables
        vd_frame = ttk.LabelFrame(var_frame, text="Dependent Variables (up to 6)", padding=5)
        vd_frame.pack(fill='x', pady=5)
        self.vd_combos = []
        self.vd_model_combos = []
        _model_options = [
            "Auto (from suffix)",
            "GLMM Gamma (Log-link)",
            "GLMM Negative Binomial (Log-link)",
            "LMM Gaussian",
        ]
        for i in range(6):
            frame = ttk.Frame(vd_frame)
            frame.pack(fill='x', pady=2)
            ttk.Label(frame, text=f"VD {i+1}:", width=10).pack(side='left')
            combo = ttk.Combobox(frame, state='readonly', width=28)
            combo.pack(side='left', padx=5)
            self.vd_combos.append(combo)
            ttk.Label(frame, text="Model:", font=('Arial', 8)).pack(side='left', padx=(10, 2))
            model_combo = ttk.Combobox(frame, state='readonly', width=32,
                                       values=_model_options)
            model_combo.set("Auto (from suffix)")
            model_combo.pack(side='left', padx=5)
            self.vd_model_combos.append(model_combo)
            # Auto-suggest model type when the user picks a VD column
            combo.bind("<<ComboboxSelected>>",
                       lambda e, mc=model_combo, vc=combo: self._on_vd_select(vc, mc))

        # Buttons
        button_frame = ttk.Frame(var_frame)
        button_frame.pack(fill='x', pady=10)

        self.export_btn = ttk.Button(button_frame, text="Export to PDF",
                                     command=self.export_to_pdf, state='disabled')
        self.export_btn.pack(side='right', padx=5)

        self.analyze_btn = ttk.Button(button_frame, text="Analyze This",
                                      command=self.run_analysis, state='disabled')
        self.analyze_btn.pack(side='right')

    def create_results_section(self, parent):
        """Create results display section."""
        results_frame = ttk.LabelFrame(parent, text="Analysis Results", padding=10)
        results_frame.pack(fill='both', expand=True)

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

        self.status_label = ttk.Label(
            self.results_frame,
            text="Load data and select variables to begin analysis",
            font=('Arial', 10, 'italic')
        )
        self.status_label.pack(pady=20)

    # ──────────────────────────── File loading ──────────────────────────────── #

    def browse_file(self):
        """Browse for Excel file."""
        filepath = filedialog.askopenfilename(
            title="Select Excel Data File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not filepath:
            return
        try:
            self.data_path = filepath
            self.data_df = pd.read_excel(filepath, engine='openpyxl')
            self.file_label.config(text=filepath, foreground='black')

            columns = [''] + list(self.data_df.columns)
            self.id_combo['values'] = columns
            for combo in self.vi_combos:
                combo['values'] = columns
            for combo in self.vd_combos:
                combo['values'] = columns

            self.analyze_btn['state'] = 'normal'
            messagebox.showinfo(
                "Success",
                f"Data loaded successfully!\n\n{len(self.data_df)} rows\n"
                f"{len(self.data_df.columns)} columns"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")

    def _on_vd_select(self, vd_combo, model_combo):
        """Auto-suggest the model type when the user selects a VD column."""
        col = vd_combo.get()
        if not col:
            model_combo.set("Auto (from suffix)")
            return
        _, label = self.detect_model_type(col)
        model_combo.set(label)

    # ────────────────────────── Analysis orchestration ──────────────────────── #

    def run_analysis(self):
        """Run GLMM analysis on all selected dependent variables."""
        if self.data_df is None:
            messagebox.showerror("Error", "Please load data file first")
            return

        animal_id = self.id_combo.get()
        if not animal_id:
            messagebox.showerror("Error", "Please select Animal ID column")
            return

        vi_vars = [combo.get() for combo in self.vi_combos if combo.get()]
        if not vi_vars:
            messagebox.showerror("Error", "Please select at least one Independent Variable")
            return

        vd_pairs = [
            (vc.get(), mc.get())
            for vc, mc in zip(self.vd_combos, self.vd_model_combos)
            if vc.get()
        ]
        vd_vars = [vd for vd, _ in vd_pairs]
        if not vd_vars:
            messagebox.showerror("Error", "Please select at least one Dependent Variable")
            return

        for widget in self.results_frame.winfo_children():
            widget.destroy()

        self.figures_for_export = []
        self.analysis_results = []

        self.status_label = ttk.Label(
            self.results_frame, text="Running analysis...", font=('Arial', 12, 'bold')
        )
        self.status_label.pack(pady=10)
        self.root.update()

        self.add_relative_day_column(animal_id)

        for vd, model_override in vd_pairs:
            self.analyze_variable(animal_id, vi_vars, vd, model_override)

        self.status_label.config(text="Analysis Complete!")
        self.export_btn['state'] = 'normal'

    def add_relative_day_column(self, animal_id):
        """Add relative day column for temporal ordering if Year/Month/Day exist."""
        if 'Relative_Day' in self.data_df.columns:
            return
        date_cols = ['Year', 'Month', 'Day']
        if not all(col in self.data_df.columns for col in date_cols):
            return
        try:
            self.data_df['Date'] = pd.to_datetime(
                self.data_df[['Year', 'Month', 'Day']].astype(str).agg('-'.join, axis=1)
            )
            relative_days = []
            for animal in self.data_df[animal_id].unique():
                animal_data = self.data_df[self.data_df[animal_id] == animal].sort_values('Date')
                if len(animal_data) > 0:
                    first_date = animal_data['Date'].min()
                    for idx, row in animal_data.iterrows():
                        days = (row['Date'] - first_date).days
                        relative_days.append((idx, days))
            for idx, days in relative_days:
                self.data_df.at[idx, 'Relative_Day'] = days
        except Exception as e:
            print(f"Could not create Relative_Day column: {e}")

    # ───────────────────────── Model type detection ─────────────────────────── #

    def detect_model_type(self, vd_var):
        """
        Detect appropriate model type from the VD column suffix.

        Returns (model_type, model_label):
          'gamma'    → GLMM Gamma (Log-link)         — suffix '_avg_time'
          'negbin'   → GLMM Negative Binomial         — suffix '_sum'
          'gaussian' → LMM Gaussian                   — everything else
        """
        if vd_var.endswith('_sum'):
            return 'negbin', 'GLMM Negative Binomial (Log-link)'
        elif vd_var.endswith('_avg_time'):
            return 'gamma', 'GLMM Gamma (Log-link)'
        else:
            return 'gaussian', 'LMM Gaussian'

    # ────────────────────────── Per-variable analysis ───────────────────────── #

    # Map from the override combo string → (model_type, model_label)
    _OVERRIDE_MAP = {
        "GLMM Gamma (Log-link)":             ('gamma',    'GLMM Gamma (Log-link)'),
        "GLMM Negative Binomial (Log-link)": ('negbin',   'GLMM Negative Binomial (Log-link)'),
        "LMM Gaussian":                       ('gaussian', 'LMM Gaussian'),
    }

    def analyze_variable(self, animal_id, vi_vars, vd_var,
                         model_override="Auto (from suffix)"):
        """
        Analyze one dependent variable.
        Statistical model is fit FIRST; its residuals are then used for the Q-Q plot.
        model_override: value from the per-VD model Combobox;
                        'Auto (from suffix)' uses suffix-based detection.
        """
        if model_override in self._OVERRIDE_MAP:
            model_type, model_label = self._OVERRIDE_MAP[model_override]
        else:
            model_type, model_label = self.detect_model_type(vd_var)

        analysis_info = {
            'animal_id': animal_id,
            'vi_vars': vi_vars,
            'vd_var': vd_var,
            'model_label': model_label,
            'qq_figure': None,
            'viz_figure': None,
            'stats_text': '',
            'qq_assessment': '',
        }

        # Build frame
        analysis_frame = ttk.LabelFrame(
            self.results_frame, text=f"Analysis: {vd_var}", padding=10
        )
        analysis_frame.pack(fill='x', padx=10, pady=10)

        info_text = (
            f"Animal ID: {animal_id}\n"
            f"Independent Variables: {', '.join(vi_vars)}\n"
            f"Dependent Variable: {vd_var}\n"
            f"Model: {model_label}\n"
        )
        ttk.Label(analysis_frame, text=info_text, font=('Arial', 9)).pack(anchor='w', pady=5)

        # Prepare data
        analysis_data = self.data_df[[animal_id] + vi_vars + [vd_var]].dropna()
        if len(analysis_data) < 10:
            ttk.Label(
                analysis_frame,
                text="Insufficient data for analysis (< 10 observations)",
                foreground='red'
            ).pack(pady=5)
            return

        # 1 ── Fit model (returns residuals for diagnostic plot)
        stats_text, residuals = self.perform_statistical_test(
            analysis_frame, analysis_data, animal_id, vi_vars, vd_var,
            model_type, model_label
        )
        analysis_info['stats_text'] = stats_text

        # 2 ── Q-Q plot on model residuals (NOT raw data)
        qq_fig, qq_assessment = self.create_qq_plot(
            analysis_frame, residuals, vd_var, model_type
        )
        analysis_info['qq_figure'] = qq_fig
        analysis_info['qq_assessment'] = qq_assessment

        # 3 ── Box-plot visualization
        viz_fig = self.create_visualization(analysis_frame, analysis_data, vi_vars, vd_var)
        analysis_info['viz_figure'] = viz_fig

        self.analysis_results.append(analysis_info)

    # ─────────────────────────── Statistical tests ──────────────────────────── #

    @staticmethod
    def _make_safe_name(name):
        """Return a formula-safe Python identifier."""
        return re.sub(r'[^a-zA-Z0-9_]', '_', str(name))

    def _sanitize_data(self, data, animal_id, vi_vars, vd_var):
        """
        Rename columns to formula-safe names.
        Returns (safe_data, safe_animal_id, safe_vi_vars, safe_vd, formula).
        """
        safe_vd = self._make_safe_name(vd_var)
        safe_vi = [self._make_safe_name(v) for v in vi_vars]
        safe_aid = self._make_safe_name(animal_id)

        col_map = {}
        if vd_var != safe_vd:
            col_map[vd_var] = safe_vd
        if animal_id != safe_aid:
            col_map[animal_id] = safe_aid
        for orig, safe in zip(vi_vars, safe_vi):
            if orig != safe:
                col_map[orig] = safe

        safe_data = data.copy()
        if col_map:
            safe_data.rename(columns=col_map, inplace=True)

        formula = f"{safe_vd} ~ {' + '.join(safe_vi)}"
        return safe_data, safe_aid, safe_vi, safe_vd, formula

    def perform_statistical_test(self, parent, data, animal_id, vi_vars, vd_var,
                                  model_type='gaussian', model_label='LMM Gaussian'):
        """
        Dispatch to the appropriate model and return (stats_text, residuals).
        residuals are Pearson residuals for GEE/GLMs, raw residuals for LMMs.
        """
        results_frame = ttk.Frame(parent)
        results_frame.pack(fill='x', pady=10)

        ttk.Label(
            results_frame,
            text=f"Statistical Results — {model_label}:",
            font=('Arial', 11, 'bold')
        ).pack(anchor='w', pady=5)

        # Sanitize column names for the formula API once
        safe_data, safe_aid, safe_vi, safe_vd, formula = self._sanitize_data(
            data, animal_id, vi_vars, vd_var
        )

        # Probe statsmodels availability
        try:
            import statsmodels.api as sm  # noqa: F401
            has_statsmodels = True
        except ImportError:
            has_statsmodels = False

        stats_text = ""
        residuals = np.array([])

        try:
            if model_type == 'gamma':
                stats_text, residuals = self._fit_gamma_model(
                    results_frame, safe_data, safe_aid, safe_vi, safe_vd,
                    formula, has_statsmodels
                )
            elif model_type == 'negbin':
                stats_text, residuals = self._fit_negbin_model(
                    results_frame, safe_data, safe_aid, safe_vi, safe_vd,
                    formula, has_statsmodels
                )
            else:
                stats_text, residuals = self._fit_gaussian_lmm(
                    results_frame, safe_data, safe_aid, safe_vi, safe_vd,
                    formula, has_statsmodels
                )
        except Exception as exc:
            err = f"Unexpected analysis error: {exc}"
            ttk.Label(results_frame, text=err, foreground='red').pack(pady=5)
            stats_text = err

        # Fallback residuals if model failed completely
        if len(residuals) == 0:
            y = safe_data[safe_vd]
            residuals = (y - y.mean()).values

        return stats_text, residuals

    # ── Gamma ────────────────────────────────────────────────────────────────── #

    def _fit_gamma_model(self, parent, data, animal_id, vi_vars, vd_var,
                         formula, has_statsmodels):
        """Try pymer4 → GEE Gamma → ANOVA fallback."""
        stats_text = ""
        residuals = np.array([])

        # Handle zero / negative values (Gamma requires strictly positive data)
        y = data[vd_var]
        if (y <= 0).any():
            n_bad = int((y <= 0).sum())
            ttk.Label(
                parent,
                text=f"Warning: {n_bad} zero/negative value(s) found — adding +0.001 offset for Gamma model.",
                foreground='orange', font=('Arial', 9, 'italic')
            ).pack(anchor='w')
            data = data.copy()
            data[vd_var] = y.clip(lower=0.001)

        # 1) pymer4 — true GLMM via R / lme4
        try:
            from pymer4.models import Lmer  # noqa: F401 (optional dependency)
            lme4_formula = f"{vd_var} ~ {' + '.join(vi_vars)} + (1|{animal_id})"
            m = Lmer(lme4_formula, data=data, family='gamma')
            m.fit()
            summary = str(m.coefs)
            self._show_results_text(parent, summary, "pymer4 GLMM Gamma — R/lme4 backend")
            stats_text = summary
            try:
                residuals = np.array(m.residuals)
            except Exception:
                residuals = np.array([])
            return stats_text, residuals

        except ImportError:
            pass
        except Exception as exc:
            ttk.Label(
                parent,
                text=f"pymer4 failed ({exc}). Falling back to GEE.",
                foreground='orange', font=('Arial', 9, 'italic')
            ).pack(anchor='w')

        # 2) GEE Gamma — quasi-GLMM with exchangeable correlation by Animal ID
        if has_statsmodels:
            try:
                from statsmodels.genmod.generalized_estimating_equations import GEE
                from statsmodels.genmod.families import Gamma
                from statsmodels.genmod.families.links import Log

                gee = GEE.from_formula(
                    formula, groups=data[animal_id], data=data,
                    family=Gamma(link=Log())
                )
                result = gee.fit()
                summary = str(result.summary())
                self._show_results_text(
                    parent, summary,
                    "GEE Gamma (quasi-GLMM · exchangeable correlation by Animal ID)"
                )
                self._highlight_significant(parent, result)
                stats_text = summary
                residuals = result.resid_pearson
                return stats_text, residuals

            except Exception as exc:
                ttk.Label(
                    parent,
                    text=f"GEE Gamma failed ({exc}). Falling back to ANOVA.",
                    foreground='red'
                ).pack(pady=5)

        # 3) ANOVA fallback
        stats_text = self.perform_anova(parent, data, vi_vars, vd_var)
        residuals = (data[vd_var] - data[vd_var].mean()).values
        return stats_text, residuals

    # ── Negative Binomial ────────────────────────────────────────────────────── #

    def _fit_negbin_model(self, parent, data, animal_id, vi_vars, vd_var,
                          formula, has_statsmodels):
        """Try pymer4 → GEE Negative Binomial → ANOVA fallback."""
        stats_text = ""
        residuals = np.array([])

        # 1) pymer4
        try:
            from pymer4.models import Lmer  # noqa: F401
            lme4_formula = f"{vd_var} ~ {' + '.join(vi_vars)} + (1|{animal_id})"
            m = Lmer(lme4_formula, data=data, family='nbinom2')
            m.fit()
            summary = str(m.coefs)
            self._show_results_text(
                parent, summary,
                "pymer4 GLMM Negative Binomial — R/lme4 backend"
            )
            stats_text = summary
            try:
                residuals = np.array(m.residuals)
            except Exception:
                residuals = np.array([])
            return stats_text, residuals

        except ImportError:
            pass
        except Exception as exc:
            ttk.Label(
                parent,
                text=f"pymer4 failed ({exc}). Falling back to GEE.",
                foreground='orange', font=('Arial', 9, 'italic')
            ).pack(anchor='w')

        # 2) GEE Negative Binomial
        if has_statsmodels:
            try:
                from statsmodels.genmod.generalized_estimating_equations import GEE
                from statsmodels.genmod.families import NegativeBinomial
                from statsmodels.genmod.families.links import Log

                gee = GEE.from_formula(
                    formula, groups=data[animal_id], data=data,
                    family=NegativeBinomial(link=Log())
                )
                result = gee.fit()
                summary = str(result.summary())
                self._show_results_text(
                    parent, summary,
                    "GEE Negative Binomial (quasi-GLMM · exchangeable correlation by Animal ID)"
                )
                self._highlight_significant(parent, result)
                stats_text = summary
                residuals = result.resid_pearson
                return stats_text, residuals

            except Exception as exc:
                ttk.Label(
                    parent,
                    text=f"GEE Negative Binomial failed ({exc}). Falling back to ANOVA.",
                    foreground='red'
                ).pack(pady=5)

        # 3) ANOVA fallback
        stats_text = self.perform_anova(parent, data, vi_vars, vd_var)
        residuals = (data[vd_var] - data[vd_var].mean()).values
        return stats_text, residuals

    # ── Gaussian LMM ──────────────────────────────────────────────────────────── #

    def _fit_gaussian_lmm(self, parent, data, animal_id, vi_vars, vd_var,
                           formula, has_statsmodels):
        """Gaussian LMM using statsmodels mixedlm (original behaviour)."""
        stats_text = ""
        residuals = np.array([])

        if has_statsmodels:
            try:
                import statsmodels.formula.api as smf
                model = smf.mixedlm(formula, data, groups=data[animal_id])
                result = model.fit()
                summary = str(result.summary())
                self._show_results_text(parent, summary,
                                        "LMM Gaussian — statsmodels mixedlm")
                self._highlight_significant(parent, result)
                stats_text = summary
                residuals = result.resid.values
                return stats_text, residuals

            except Exception as exc:
                ttk.Label(
                    parent,
                    text=f"LMM failed ({exc}). Falling back to ANOVA.",
                    foreground='orange'
                ).pack(pady=5)
        else:
            ttk.Label(
                parent,
                text=(
                    "statsmodels not installed — using simplified ANOVA.\n"
                    "Install with: pip install statsmodels"
                ),
                foreground='blue', font=('Arial', 9, 'italic')
            ).pack(pady=5)

        stats_text = self.perform_anova(parent, data, vi_vars, vd_var)
        residuals = (data[vd_var] - data[vd_var].mean()).values
        return stats_text, residuals

    # ── Shared helpers ────────────────────────────────────────────────────────── #

    def _show_results_text(self, parent, text, method_note=""):
        """Render model summary in a scrolled text widget."""
        widget = scrolledtext.ScrolledText(parent, height=10, width=80)
        widget.pack(pady=5)
        widget.insert('1.0', text)
        if method_note:
            ttk.Label(
                parent, text=f"Method: {method_note}",
                foreground='blue', font=('Arial', 9, 'italic')
            ).pack(anchor='w')

    def _highlight_significant(self, parent, result):
        """Show a bold red label listing any significant predictors (p < 0.05)."""
        try:
            sig = [
                v for v, p in result.pvalues.items()
                if p < 0.05 and v != 'Intercept'
            ]
            if sig:
                ttk.Label(
                    parent,
                    text="Significant effects (p < 0.05): " + ", ".join(sig),
                    foreground='red', font=('Arial', 10, 'bold')
                ).pack(pady=5)
        except Exception:
            pass

    def perform_anova(self, parent, data, vi_vars, vd_var):
        """One-way ANOVA for each independent variable (last-resort fallback)."""
        widget = scrolledtext.ScrolledText(parent, height=8, width=80)
        widget.pack(pady=5)
        widget.insert('1.0', "ANOVA Results (Simplified Fallback):\n\n")
        anova_text = "ANOVA Results (Simplified Fallback):\n\n"

        for vi in vi_vars:
            groups = [
                data[data[vi] == lvl][vd_var]
                for lvl in data[vi].unique()
                if len(data[data[vi] == lvl][vd_var]) > 0
            ]
            if len(groups) >= 2:
                f_stat, p_value = stats.f_oneway(*groups)
                sig = " ***SIGNIFICANT***" if p_value < 0.05 else ""
                line = f"{vi}: F={f_stat:.4f}, p={p_value:.4f}{sig}\n"
                widget.insert(tk.END, line)
                anova_text += line

        return anova_text

    # ────────────────────── Q-Q plot (model residuals) ──────────────────────── #

    def create_qq_plot(self, parent, residuals, var_name, model_type='gaussian'):
        """
        Create a Q-Q plot of model residuals (Pearson / deviance / raw) to assess
        model fit.

        This replaces the old 'raw-data normality check'.  The Shapiro-Wilk test
        is here applied to the RESIDUALS (a valid diagnostic), not to the raw VD.
        """
        ttk.Label(
            parent,
            text="Model Diagnostics (Residuals Analysis)",
            font=('Arial', 11, 'bold')
        ).pack(anchor='w', pady=(10, 2))

        fig = Figure(figsize=(6, 4))
        ax = fig.add_subplot(111)

        residuals = np.asarray(residuals, dtype=float)
        residuals = residuals[np.isfinite(residuals)]   # drop inf / NaN

        assessment = "Insufficient residuals"
        color = 'gray'

        if len(residuals) >= 3:
            stats.probplot(residuals, dist="norm", plot=ax)

        ax.set_title(f"Q-Q Plot of Model Residuals\n{var_name}  [{model_type}]")
        ax.grid(True, alpha=0.3)

        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=5)

        # Shapiro-Wilk on residuals
        if len(residuals) >= 3:
            if len(residuals) <= 5000:
                _, p = stats.shapiro(residuals)
                if p > 0.05:
                    assessment = (
                        f"Residuals appear normal (Shapiro-Wilk p={p:.4f}) "
                        f"— model fit acceptable"
                    )
                    color = 'green'
                elif p > 0.01:
                    assessment = (
                        f"Residuals borderline normal (Shapiro-Wilk p={p:.4f}) "
                        f"— check model fit"
                    )
                    color = 'orange'
                else:
                    assessment = (
                        f"Residuals not normal (Shapiro-Wilk p={p:.4f}) "
                        f"— consider model revision"
                    )
                    color = 'red'
            else:
                assessment = (
                    f"n={len(residuals)} — Shapiro-Wilk not applied (n > 5000); "
                    f"inspect Q-Q plot visually"
                )
                color = 'gray'

        ttk.Label(
            parent, text=assessment, foreground=color,
            font=('Arial', 10, 'bold')
        ).pack(pady=5)

        return fig, assessment

    # ─────────────────────────── Visualization ──────────────────────────────── #

    def create_visualization(self, parent, data, vi_vars, vd_var):
        """Box-plots of VD per level of each VI."""
        fig = Figure(figsize=(10, 4))
        n_plots = len(vi_vars)

        for i, vi in enumerate(vi_vars, 1):
            ax = fig.add_subplot(1, n_plots, i)
            unique_levels = sorted(data[vi].unique())
            plot_data = [data[data[vi] == lvl][vd_var].values for lvl in unique_levels]
            ax.boxplot(plot_data, labels=[str(lvl) for lvl in unique_levels])
            ax.set_xlabel(vi)
            ax.set_ylabel(vd_var)
            ax.set_title(f'{vd_var} by {vi}')
            ax.grid(True, alpha=0.3)

        fig.tight_layout()
        canvas = FigureCanvasTkAgg(fig, parent)
        canvas.draw()
        canvas.get_tk_widget().pack(pady=10)
        return fig

    # ───────────────────────────── PDF export ───────────────────────────────── #

    def export_to_pdf(self):
        """Export all analysis results to a multi-page PDF report."""
        if not self.analysis_results:
            messagebox.showerror("Error", "No analysis results to export")
            return

        try:
            if self.data_path:
                base_name = Path(self.data_path).stem
                default_filename = f"{base_name}_GLMM.pdf"
                output_dir = Path(self.data_path).parent
            else:
                default_filename = "GLMM_Analysis.pdf"
                output_dir = Path.home()

            pdf_path = filedialog.asksaveasfilename(
                title="Save GLMM Analysis Report",
                defaultextension=".pdf",
                initialfile=default_filename,
                initialdir=output_dir,
                filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
            )
            if not pdf_path:
                return

            with PdfPages(pdf_path) as pdf:
                # ── Title page ──────────────────────────────────────────────── #
                fig = plt.figure(figsize=(8.5, 11))
                fig.text(0.5, 0.7, "GLMM Analysis Report",
                         ha='center', va='center', fontsize=20, weight='bold')
                fig.text(
                    0.5, 0.6,
                    f"Data File: {Path(self.data_path).name if self.data_path else 'Unknown'}",
                    ha='center', va='center', fontsize=12
                )
                fig.text(
                    0.5, 0.5,
                    f"Generated: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}",
                    ha='center', va='center', fontsize=10
                )
                fig.text(
                    0.5, 0.3,
                    f"Total Analyses: {len(self.analysis_results)}",
                    ha='center', va='center', fontsize=12
                )
                pdf.savefig(fig, bbox_inches='tight')
                plt.close(fig)

                # ── One summary page + figures per analysis ─────────────────── #
                for i, result in enumerate(self.analysis_results, 1):
                    fig = plt.figure(figsize=(8.5, 11))

                    fig.text(
                        0.5, 0.95, f"Analysis {i}: {result['vd_var']}",
                        ha='center', va='top', fontsize=16, weight='bold'
                    )

                    info_text = (
                        f"Animal ID: {result['animal_id']}\n"
                        f"Independent Variables: {', '.join(result['vi_vars'])}\n"
                        f"Dependent Variable: {result['vd_var']}\n"
                        f"Model: {result.get('model_label', 'N/A')}\n"
                    )
                    fig.text(0.1, 0.85, info_text,
                             ha='left', va='top', fontsize=10, family='monospace')

                    fig.text(0.1, 0.67,
                             "Model Diagnostics (Residuals Analysis):",
                             ha='left', va='top', fontsize=11, weight='bold')
                    fig.text(0.1, 0.64, result['qq_assessment'],
                             ha='left', va='top', fontsize=9)

                    stats_display = result['stats_text']
                    if len(stats_display) > 1200:
                        stats_display = stats_display[:1200] + "\n[... truncated, see console ...]"
                    fig.text(0.1, 0.57, "Statistical Results:",
                             ha='left', va='top', fontsize=11, weight='bold')
                    fig.text(0.1, 0.54, stats_display,
                             ha='left', va='top', fontsize=7, family='monospace', wrap=True)

                    pdf.savefig(fig, bbox_inches='tight')
                    plt.close(fig)

                    if result['qq_figure']:
                        pdf.savefig(result['qq_figure'], bbox_inches='tight')

                    if result['viz_figure']:
                        pdf.savefig(result['viz_figure'], bbox_inches='tight')

            messagebox.showinfo("Success", f"Analysis report exported to:\n{pdf_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF:\n{str(e)}")


def main():
    root = tk.Tk()
    app = GLMMAnalyzer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
