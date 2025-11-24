import os
import sys
import pandas as pd
import numpy as np
from scipy.optimize import minimize
from sklearn.linear_model import Ridge
# from sklearn.model_selection import cross_val_score # Removed
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font as tkFont
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import matplotlib.patches # For Wedge
from PIL import Image, ImageDraw, ImageFont
import math

class _SilentStream:
    def write(self, _msg=None):
        pass

    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = _SilentStream()
if sys.stderr is None:
    sys.stderr = _SilentStream()

from OACD import OACD

class CSRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python CSR Integration App")
        self.root.geometry("{0}x{1}+0+0".format(self.root.winfo_screenwidth(), self.root.winfo_screenheight()))

        # --- Style Configuration ---
        self.style = ttk.Style()
        available_themes = self.style.theme_names()
        if 'vista' in available_themes and self.root.tk.call('tk', 'windowingsystem') == 'win32':
            self.style.theme_use('vista')
        elif 'clam' in available_themes:
            self.style.theme_use('clam')
        elif 'aqua' in available_themes and self.root.tk.call('tk', 'windowingsystem') == 'aqua':
            self.style.theme_use('aqua')

        # Define custom fonts
        base_font_family = "Segoe UI" if self.root.tk.call('tk', 'windowingsystem') == 'win32' else "Helvetica"

        self.default_font = tkFont.nametofont("TkDefaultFont")
        self.default_font.configure(family=base_font_family, size=10)
        self.root.option_add("*Font", self.default_font)

        try:
            tkFont.Font(family=f"{base_font_family} Semibold", size=10).actual()
            semibold_family = f"{base_font_family} Semibold"
            bold_weight = tkFont.NORMAL
        except tk.TclError:
            semibold_family = base_font_family
            bold_weight = tkFont.BOLD

        self.title_font = tkFont.Font(family=semibold_family, size=11, weight=bold_weight if semibold_family == base_font_family else tkFont.NORMAL)
        self.label_font = tkFont.Font(family=base_font_family, size=10)
        self.button_font = tkFont.Font(family=base_font_family, size=10, weight=tkFont.BOLD)
        self.entry_font = tkFont.Font(family=base_font_family, size=10)
        self.text_widget_font = tkFont.Font(family="Consolas" if self.root.tk.call('tk', 'windowingsystem') == 'win32' else "Courier New", size=10)

        # Configure ttk widget styles
        self.style.configure("TButton", padding=(8, 6), relief="flat", font=self.button_font)
        self.style.map("TButton",
            foreground=[('pressed', 'black'), ('active', '#005A9E')],
            background=[('pressed', '!disabled', '#E0E0E0'), ('active', '#E5F1FB')]
        )
        self.style.configure("TLabel", font=self.label_font, padding=(0, 2, 0, 2))
        self.style.configure("TEntry", font=self.entry_font, padding=4)
        self.style.configure("TCombobox", font=self.entry_font, padding=4)
        self.style.configure("Treeview.Heading", font=(semibold_family, 10, bold_weight if semibold_family == base_font_family else tkFont.NORMAL), padding=5)
        self.style.configure("Treeview", rowheight=28, font=(base_font_family, 10))
        self.style.configure("TLabelframe.Label", font=self.title_font, padding=(0, 0, 0, 5))
        self.style.configure("TNotebook.Tab", font=(semibold_family, 10, bold_weight if semibold_family == base_font_family else tkFont.NORMAL), padding=[12, 6])
        self.style.configure("App.TFrame", background='#F0F0F0')

        self.root.configure(background='#F0F0F0')

        self.df = None
        self.coefficients = None
        self.bits_array = None
        self.factor_cols = []
        self.X = None
        self.X_original_scale = None
        self.y = None
        self.y_pred = None
        self.col_name_mapping = {}
        self.original_col_names = []
        self.extremum_point = None
        self.x_min_orig = None
        self.x_max_orig = None
        self.norm_x_min = None
        self.norm_x_max = None

        self.term_type_colors = {
            'Constant': '#800080',    # Matplotlib Blue
            'Linear': '#2ca02c',      # Matplotlib Green
            'Quadratic': '#d62728',   # Matplotlib Red
            'Interaction': '#1f77b4', # Matplotlib Purple
            'Default': '#7f7f7f'      # Matplotlib Grey (fallback)
        }
        self.bits_array = None # Ensure bits_array is initialized to None

        # Initialize CSR limits (similar to OACD limits)
        self.csr_limits = {}
        self.csr_limit_names = {}
        self.csr_factor_checkboxes = {}

        self.notebook = ttk.Notebook(self.root, style="TNotebook")

        self.create_widgets()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self.root, style="TNotebook")
        self.notebook.pack(expand=True, fill='both', padx=10, pady=10)

        self.tab1 = ttk.Frame(self.notebook, padding=10, style="App.TFrame")
        self.tab2 = ttk.Frame(self.notebook, padding=10, style="App.TFrame")
        self.tab3 = ttk.Frame(self.notebook, padding=10, style="App.TFrame")

        self.notebook.add(self.tab1, text='CSR Integration')
        self.notebook.add(self.tab2, text='Coefficient Analysis')
        self.notebook.add(self.tab3, text='OACD Table Builder')

        self.create_csr_integration_tab()
        self.create_coefficient_analysis_tab()
        self.create_oacd_tab()

    def create_csr_integration_tab(self):
        self.main_paned = tk.PanedWindow(self.tab1, orient=tk.HORIZONTAL, sashrelief=tk.GROOVE, sashwidth=8, background="#D0D0D0", bd=0)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        # === Left Column with Scrollbar ===
        left_container = ttk.Frame(self.main_paned, style="App.TFrame")
        self.main_paned.add(left_container, width=430, minsize=400, sticky="nsew")
        
        # Left scrollbar setup
        left_canvas = tk.Canvas(left_container, highlightthickness=0)
        left_scrollbar = ttk.Scrollbar(left_container, orient="vertical", command=left_canvas.yview)
        left_scrollable_frame = ttk.Frame(left_canvas, style="App.TFrame")
        
        left_scrollable_frame.bind(
            "<Configure>",
            lambda e: left_canvas.configure(scrollregion=left_canvas.bbox("all"))
        )
        
        left_canvas.create_window((0, 0), window=left_scrollable_frame, anchor="nw")
        left_canvas.configure(yscrollcommand=left_scrollbar.set)
        
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scrollbar.pack(side="right", fill="y")

        # === Center Column with Scrollbar ===
        center_container = ttk.Frame(self.main_paned, style="App.TFrame")
        self.main_paned.add(center_container, width=650, minsize=500, sticky="nsew")
        
        # Center scrollbar setup
        center_canvas = tk.Canvas(center_container, highlightthickness=0)
        center_scrollbar = ttk.Scrollbar(center_container, orient="vertical", command=center_canvas.yview)
        center_scrollable_frame = ttk.Frame(center_canvas, style="App.TFrame")
        
        center_scrollable_frame.bind(
            "<Configure>",
            lambda e: center_canvas.configure(scrollregion=center_canvas.bbox("all"))
        )
        
        center_canvas.create_window((0, 0), window=center_scrollable_frame, anchor="nw")
        center_canvas.configure(yscrollcommand=center_scrollbar.set)
        
        center_canvas.pack(side="left", fill="both", expand=True)
        center_scrollbar.pack(side="right", fill="y")

        # === Right Column with Scrollbar ===
        right_container = ttk.Frame(self.main_paned, style="App.TFrame")
        self.main_paned.add(right_container, width=450, minsize=420, sticky="nsew")
        
        # Right scrollbar setup
        right_canvas = tk.Canvas(right_container, highlightthickness=0)
        right_scrollbar = ttk.Scrollbar(right_container, orient="vertical", command=right_canvas.yview)
        right_scrollable_frame = ttk.Frame(right_canvas, style="App.TFrame")
        
        right_scrollable_frame.bind(
            "<Configure>",
            lambda e: right_canvas.configure(scrollregion=right_canvas.bbox("all"))
        )
        
        right_canvas.create_window((0, 0), window=right_scrollable_frame, anchor="nw")
        right_canvas.configure(yscrollcommand=right_scrollbar.set)
        
        right_canvas.pack(side="left", fill="both", expand=True)
        right_scrollbar.pack(side="right", fill="y")

        # === Left Panel Contents ===
        left_frame = ttk.Frame(left_scrollable_frame, padding=15, style="App.TFrame")
        left_frame.pack(fill='both', expand=True)
        
        self.style.configure("Accent.TButton", font=self.button_font, foreground="white", background="#0078D7")
        self.style.map("Accent.TButton", background=[('active', '#005A9E'), ('pressed', '!disabled', '#004C8A')])
        ttk.Button(left_frame, text="Select Data File", command=self.select_file, style="Accent.TButton").pack(pady=(0,15), padx=5, fill='x')

        ttk.Label(left_frame, text="Project Name / File:").pack(anchor='w', padx=5)
        self.project_name_entry = ttk.Entry(left_frame, font=self.entry_font)
        self.project_name_entry.pack(fill='x', padx=5, pady=(2,10))

        ttk.Label(left_frame, text="Extremum Objective:").pack(anchor='w', padx=5)
        self.weight_combo = ttk.Combobox(left_frame,
                                        values=["Maximum", "Minimum",
                                                "Maximum absolute value",
                                                "Minimum absolute value"],
                                        state="readonly", font=self.entry_font)
        self.weight_combo.pack(fill='x', padx=5, pady=(2,10))
        self.weight_combo.current(0)
        
        ttk.Label(left_frame, text="Normalization:").pack(anchor='w', padx=5)
        self.norm_select = ttk.Combobox(left_frame, values=["[-1, 1]", "[0, 1]", "No normalization"], state="readonly", font=self.entry_font)
        self.norm_select.pack(fill='x', padx=5, pady=(2,10))
        self.norm_select.current(1)  # Default to [0, 1]

        # === Show All Factor Combinations Option ===
        self.show_all_combinations_var = tk.BooleanVar(value=False)  # Default to showing all
        show_all_frame = ttk.Frame(left_frame, style="App.TFrame")
        show_all_frame.pack(fill='x', pady=(0,10), padx=5)
        ttk.Checkbutton(show_all_frame, text="Show results for all parameter combinations (N down to 1)", 
                        variable=self.show_all_combinations_var, 
                        command=self._update_show_all_combinations).pack(anchor='w')
        
        # === CSR Factor Limits Section ===
        csr_limits_frame = ttk.LabelFrame(left_frame, text="CSR Parameter Limits", padding=(10,5,10,10))
        csr_limits_frame.pack(fill='x', pady=10, padx=5)
        
        # Limit value input
        limit_input_frame = ttk.Frame(csr_limits_frame, style="App.TFrame")
        limit_input_frame.pack(fill='x', pady=(0,5))
        ttk.Label(limit_input_frame, text="Limit Value:").pack(side='left')
        self.csr_limit_value = tk.DoubleVar(value=100.0)
        limit_value_entry = tk.Entry(limit_input_frame, textvariable=self.csr_limit_value, width=8, font=self.entry_font)
        limit_value_entry.pack(side='left', padx=(2,8))
        ttk.Button(limit_input_frame, text="Add Limit", command=self._csr_add_limit, width=10).pack(side='left')
        
        # Current limits display
        limits_display_frame = ttk.Frame(csr_limits_frame, style="App.TFrame")
        limits_display_frame.pack(fill='x', pady=(5,0))
        ttk.Label(limits_display_frame, text="Current Limits:").pack(anchor='w')
        self.csr_limits_listbox = tk.Listbox(limits_display_frame, height=4, font=self.entry_font)
        self.csr_limits_listbox.pack(fill='x', pady=(2,5))
        limits_button_frame = ttk.Frame(limits_display_frame, style="App.TFrame")
        limits_button_frame.pack(fill='x')
        ttk.Button(limits_button_frame, text="Remove Selected", command=self._csr_remove_limit, width=12).pack(side='left')
        ttk.Button(limits_button_frame, text="Clear All", command=self._csr_clear_limits, width=8).pack(side='left', padx=(5,0))

        # === Parameter Limits Selection ===
        factor_limits_frame = ttk.LabelFrame(left_frame, text="Parameter Limits Selection", padding=(10,5,10,10))
        factor_limits_frame.pack(fill='x', pady=10, padx=5)
        
        self.factor_limits_frame = ttk.Frame(factor_limits_frame, style="App.TFrame")
        self.factor_limits_frame.pack(fill='both', expand=True)
        
        ttk.Label(self.factor_limits_frame, text="Select parameters for limits (will be applied during optimization):", 
                font=self.label_font).pack(anchor='w', pady=(0,5))
        
        # Container for factor limit checkboxes
        self.factor_limits_container = ttk.Frame(self.factor_limits_frame, style="App.TFrame")
        self.factor_limits_container.pack(fill='x', pady=5)

        ttk.Button(left_frame, text="Run Fitting Process", command=self.run_fitting).pack(pady=15, padx=5, fill='x', ipady=5)

        # === Center Panel Contents ===
        center_frame = ttk.Frame(center_scrollable_frame, padding=(5,15,15,15), style="App.TFrame")
        center_frame.pack(fill='both', expand=True)

        # Factor selection frame (moved to top of center column)
        factor_frame = ttk.LabelFrame(center_frame, text="Parameter or Outcome?", padding=10)
        factor_frame.pack(fill='x', pady=(0,10))
        
        self.factor_selection_frame = ttk.Frame(factor_frame, style="App.TFrame")
        self.factor_selection_frame.pack(fill='both', expand=True)
        
        # Initialize selection dictionary
        self.factor_checkboxes = {}
        self.factor_selection = {}
        
        # Add a label for the factor and result selection
        ttk.Label(self.factor_selection_frame, text="Select parameters and outcomes to include in analysis:", 
                font=self.label_font).pack(anchor='w', pady=(0,10))
        
        # Key Results Frame (now contains CSR Equation and Parameter Definitions)
        results_frame = ttk.LabelFrame(center_frame, text="Key Results", padding=(10,5,10,10))
        results_frame.pack(fill='x', pady=(10,0), expand=True)
        
        # Function Definition Frame (MOVED inside Key Results)
        function_def_frame = ttk.LabelFrame(results_frame, text="CSR Equation", padding=(10, 5, 10, 10))
        function_def_frame.pack(fill='x', pady=(0, 10), padx=5)

        eqn_text_frame = ttk.Frame(function_def_frame, style="App.TFrame")
        eqn_text_frame.pack(fill='both', expand=True, padx=0, pady=(0, 5))

        # Create frame for text widget and scrollbars
        eqn_container = ttk.Frame(eqn_text_frame, style="App.TFrame")
        eqn_container.pack(fill='both', expand=True)

        self.equation_text = tk.Text(eqn_container, height=6, wrap=tk.WORD, state='disabled',  # ← CHANGED: wrap=tk.WORD and increased height
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=8, pady=5, spacing1=2, spacing2=2, spacing3=2)

        # Vertical scrollbar only (remove horizontal scrollbar for equation)
        v_scrollbar_eqn = ttk.Scrollbar(eqn_container, command=self.equation_text.yview, orient=tk.VERTICAL)
        self.equation_text['yscrollcommand'] = v_scrollbar_eqn.set

        # Simple layout - no horizontal scrollbar needed for word-wrapped text
        self.equation_text.pack(side=tk.LEFT, fill='both', expand=True)
        v_scrollbar_eqn.pack(side=tk.RIGHT, fill=tk.Y)

        # Factor Definition Frame (MOVED inside Key Results)
        factor_def_frame = ttk.LabelFrame(results_frame, text="Parameter Definitions", padding=(10, 5, 10, 10))
        factor_def_frame.pack(fill='x', pady=(0, 10), padx=5)

        factor_def_text_frame = ttk.Frame(factor_def_frame, style="App.TFrame")
        factor_def_text_frame.pack(fill='both', expand=True, padx=0, pady=(0, 5))

        self.factor_definitions_text = tk.Text(factor_def_text_frame, height=3, wrap=tk.WORD, state='disabled',
                                            font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=8, pady=5,spacing1=2, spacing2=2, spacing3=2)
        scrollbar_factor_def = ttk.Scrollbar(factor_def_text_frame, command=self.factor_definitions_text.yview, orient=tk.VERTICAL)
        self.factor_definitions_text['yscrollcommand'] = scrollbar_factor_def.set
        scrollbar_factor_def.pack(side=tk.RIGHT, fill=tk.Y)
        self.factor_definitions_text.pack(side=tk.LEFT, fill='both', expand=True)

        # Extremum section (now below CSR Equation and Parameter Definitions)
        extremum_frame = ttk.LabelFrame(results_frame, text="Extremum (Original Scale)", padding=(10,5,10,10))
        extremum_frame.pack(fill='x', pady=(0,10))

        # Factors display
        factors_display_frame = ttk.Frame(extremum_frame, style="App.TFrame")
        factors_display_frame.pack(fill='x', pady=5)
        ttk.Label(factors_display_frame, text="Results:").pack(anchor='w', pady=(0,5))  # ← Moved to top

        # Create container for text widget and scrollbars
        factors_container = ttk.Frame(factors_display_frame, style="App.TFrame")
        factors_container.pack(fill='both', expand=True)

        self.factors_text = tk.Text(factors_container, height=8, wrap=tk.NONE, state='disabled',  # ← Changed to NONE for horizontal scrolling
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)

        # Vertical scrollbar
        v_scrollbar_factors = ttk.Scrollbar(factors_container, command=self.factors_text.yview, orient=tk.VERTICAL)
        self.factors_text['yscrollcommand'] = v_scrollbar_factors.set

        # Horizontal scrollbar  
        h_scrollbar_factors = ttk.Scrollbar(factors_container, command=self.factors_text.xview, orient=tk.HORIZONTAL)
        self.factors_text['xscrollcommand'] = h_scrollbar_factors.set

        # Grid layout for text and scrollbars
        self.factors_text.grid(row=0, column=0, sticky='nsew')
        v_scrollbar_factors.grid(row=0, column=1, sticky='ns')
        h_scrollbar_factors.grid(row=1, column=0, sticky='ew')

        factors_container.grid_rowconfigure(0, weight=1)
        factors_container.grid_columnconfigure(0, weight=1)
        
        # R² frame
        self.r2_frame = ttk.LabelFrame(results_frame, text="Model Fit (R²)", padding=(10,5,10,10))
        self.r2_frame.pack(fill='x', pady=(0,10))
        
        # Initialize R² text widget
        self.r2_text = tk.Text(self.r2_frame, height=1, wrap=tk.NONE, state='disabled',
                            font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
        self.r2_text.pack(fill='x', padx=5, pady=5)

        # === Right Panel Contents ===
        right_frame = ttk.Frame(right_scrollable_frame, style="App.TFrame", padding=0)
        right_frame.pack(fill='both', expand=True)

        # Create vertical paned window for plots
        right_paned = tk.PanedWindow(right_frame, orient=tk.VERTICAL, sashrelief=tk.GROOVE, sashwidth=8, background="#D0D0D0", bd=0)
        right_paned.pack(fill='both', expand=True, padx=5, pady=0)

        plot1_frame = ttk.LabelFrame(right_paned, text="Actual vs. Predicted Values", padding=10)
        right_paned.add(plot1_frame)
        right_paned.paneconfig(plot1_frame)

        self.figure1 = Figure(figsize=(5, 4), dpi=100, facecolor='#F0F0F0')
        self.figure1.subplots_adjust(bottom=0.18, left=0.18, top=0.9, right=0.95)
        self.canvas1 = FigureCanvasTkAgg(self.figure1, master=plot1_frame)
        self.canvas1.get_tk_widget().pack(fill='both', expand=True, padx=5, pady=5)
        toolbar1 = NavigationToolbar2Tk(self.canvas1, plot1_frame)
        toolbar1.update()
        toolbar1.configure(background='#F0F0F0')
        for child_widget in toolbar1.winfo_children(): child_widget.configure(background='#F0F0F0')

        plot2_frame = ttk.LabelFrame(right_paned, text="CSR Response Surface Plot", padding=10)
        right_paned.add(plot2_frame)
        right_paned.paneconfig(plot2_frame)

        factor_control_frame = ttk.Frame(plot2_frame, style="App.TFrame")
        factor_control_frame.pack(fill='x', pady=(5,8))
        ttk.Label(factor_control_frame, text="X-axis:").pack(side='left', padx=(0,3))
        self.x_factor_combo = ttk.Combobox(factor_control_frame, state='readonly', width=12, font=self.entry_font)
        self.x_factor_combo.pack(side='left', padx=(0,10))
        ttk.Label(factor_control_frame, text="Y-axis:").pack(side='left', padx=(0,3))
        self.y_factor_combo = ttk.Combobox(factor_control_frame, state='readonly', width=12, font=self.entry_font)
        self.y_factor_combo.pack(side='left', padx=(0,10))
        ttk.Button(factor_control_frame, text="Update", command=self.update_3d_plot, width=8).pack(side='left', padx=5)

        self.figure2 = Figure(figsize=(5, 4), dpi=100, facecolor='#F0F0F0')
        self.figure2.subplots_adjust(bottom=0.18, left=0.1, right=0.85, top=0.9)
        self.canvas2 = FigureCanvasTkAgg(self.figure2, master=plot2_frame)
        self.canvas2.get_tk_widget().pack(fill='both', expand=True, padx=5, pady=5)
        toolbar2 = NavigationToolbar2Tk(self.canvas2, plot2_frame)
        toolbar2.update()
        toolbar2.configure(background='#F0F0F0')
        for child_widget in toolbar2.winfo_children(): child_widget.configure(background='#F0F0F0')

        # Update the UI after initial setup
        self._update_csr_factor_limits_ui()

    # CSR Limits Management Methods
    def _csr_add_limit(self):
        """Add a new limit with selected factors"""
        if not hasattr(self, 'factor_cols') or not self.factor_cols:
            messagebox.showwarning("No Factors", "Please load data with factors first.")
            return
            
        limit_value = self.csr_limit_value.get()
        if limit_value <= 0:
            messagebox.showwarning("Invalid Value", "Limit value must be positive.")
            return
            
        # Get selected factors from checkboxes
        selected_factors = []
        for i, var in self.csr_factor_checkboxes.items():
            if var.get():  # Checkbox is checked
                selected_factors.append(i)
        
        if not selected_factors:
            messagebox.showwarning("No Selection", "Please select at least one factor for the limit.")
            return
            
        # Add limit to CSR limits - CHANGE: Add constraint type option
        limit_name = f"limit_{len(self.csr_limits) + 1}"
        self.csr_limits[limit_name] = {
            'factors': selected_factors,
            'value': limit_value,
            'type': 'sum_equality'  # CHANGE: Use equality constraint
        }
        self.csr_limit_names[limit_name] = f"Limit {len(self.csr_limits)}: Sum of F{','.join(map(str, [f+1 for f in selected_factors]))} = {limit_value}"
        
        # Update limits display
        self._csr_update_limits_display()

    def _csr_remove_limit(self):
        """Remove the selected limit"""
        selection = self.csr_limits_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a limit to remove.")
            return
            
        # Get the limit key from the listbox
        limit_name = list(self.csr_limits.keys())[selection[0]]
        
        # Remove the limit
        del self.csr_limits[limit_name]
        del self.csr_limit_names[limit_name]
            
        # Update limits display
        self._csr_update_limits_display()

    def _csr_clear_limits(self):
        """Clear all limits"""
        if messagebox.askyesno("Clear All Limits", "Are you sure you want to clear all limits?"):
            self.csr_limits.clear()
            self.csr_limit_names.clear()
            self._csr_update_limits_display()

    def _csr_update_limits_display(self):
        """Update the limits listbox display"""
        self.csr_limits_listbox.delete(0, tk.END)
        
        for name in self.csr_limit_names.values():
            self.csr_limits_listbox.insert(tk.END, name)

    def _update_csr_factor_limits_ui(self):
        """Update the factor limits selection UI when factors are loaded"""
        # Clear existing widgets
        for widget in self.factor_limits_container.winfo_children():
            widget.destroy()
        
        self.csr_factor_checkboxes = {}  # Change from combos to checkboxes
        
        if not hasattr(self, 'factor_cols') or not self.factor_cols:
            return
            
        # Create factor selection checkboxes using pack
        for i, factor in enumerate(self.factor_cols):
            display_name = self.col_name_mapping.get(factor, factor)
            
            # Create frame for this factor
            factor_frame = ttk.Frame(self.factor_limits_container, style="App.TFrame")
            factor_frame.pack(fill='x', pady=2)
            
            # Create checkbox instead of combobox
            var = tk.BooleanVar(value=False)
            checkbox = ttk.Checkbutton(factor_frame, text=display_name, variable=var)
            checkbox.pack(side='left', padx=(0,10))
            
            self.csr_factor_checkboxes[i] = var  # Store by index
    
    # Add this helper method
    def _update_show_all_combinations(self):
        """Update the display when the show all combinations option changes"""
        if hasattr(self, 'coefficients') and self.coefficients is not None:
            # Refresh the display with current results
            if hasattr(self, 'result_functions'):
                # For comprehensive optimization
                result_min_max = {}
                for result_col, func_data in self.result_functions.items():
                    result_min_max[result_col] = {
                        'min': func_data['min_val'],
                        'max': func_data['max_val']
                    }
                self.update_comprehensive_results_display(result_min_max)
            else:
                # For single result optimization
                n_factors = len(self.factor_cols)
                equation_str, function_defs_str, factor_defs_str = self.generate_equation_and_definitions(
                    self.coefficients, self.bits_array, n_factors)
                # Get R² value
                if hasattr(self, 'X') and self.X is not None:
                    X_design = self.create_design_matrix(self.X, self.bits_array)
                    train_r2 = np.corrcoef(self.y, self.y_pred)[0,1]**2 if hasattr(self, 'y_pred') and self.y_pred is not None else 0
                else:
                    train_r2 = 0
                self.update_results_display(equation_str, function_defs_str, factor_defs_str, train_r2)

    def create_coefficient_analysis_tab(self):
            # Main container frame
            coeff_tab_frame = ttk.Frame(self.tab2, style="App.TFrame")
            coeff_tab_frame.pack(expand=True, fill='both', padx=10, pady=10)

            # Top part with controls and charts
            top_frame = ttk.Frame(coeff_tab_frame, style="App.TFrame")
            top_frame.pack(expand=True, fill='both')

            # Controls frame
            controls_frame = ttk.LabelFrame(top_frame, text="Analysis Controls", padding=10)
            controls_frame.pack(fill='x', pady=(0,10))

            ttk.Label(controls_frame, text="Evaluate Term Contributions at Parameter Levels:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
            self.eval_point_combo = ttk.Combobox(controls_frame,
                                                 values=["Factors at Minimum",
                                                         "Factors at Extremum",
                                                         "Factors at Maximum"],
                                                 state="readonly", width=28, font=self.entry_font)
            self.eval_point_combo.grid(row=0, column=1, padx=5, pady=5, sticky='ew')
            self.eval_point_combo.current(1)

            self.refresh_coeffs_button = ttk.Button(controls_frame, text="Refresh Charts",
                                                 command=self.update_coefficient_pie_charts)
            self.refresh_coeffs_button.grid(row=0, column=2, padx=10, pady=5, ipady=2)
            
            # --- Add Single Download All Charts Button ---
            self.download_all_button = ttk.Button(controls_frame, text="Download All Charts",
                                                  command=self._download_all_charts)
            self.download_all_button.grid(row=0, column=3, padx=10, pady=5, ipady=2) # Added to the right
            # --- End Single Download Button ---

            controls_frame.columnconfigure(1, weight=1) # Ensure combobox expands if needed


            # Definitions frame - always visible (remains the same)
            self.definitions_frame = ttk.LabelFrame(top_frame, text="Parameter Definitions", padding=10)
            # ... (rest of definitions_frame as before) ...
            self.definitions_frame.pack(fill='x', pady=(0,10))

            self.definitions_text = tk.Text(
                self.definitions_frame,
                height=4,
                wrap=tk.WORD,
                font=self.text_widget_font,
                background='#F0F0F0',
                relief=tk.FLAT
            )
            self.definitions_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            self.definitions_text.insert(tk.END, "Parameter definitions will appear here after fitting")
            self.definitions_text.config(state=tk.DISABLED)


            # Charts frame (layout remains the same, individual download buttons removed from here)
            charts_frame = ttk.Frame(top_frame, style="App.TFrame")
            charts_frame.pack(expand=True, fill='both', pady=10)

            if not hasattr(self, 'pie_figures'): self.pie_figures = {}
            if not hasattr(self, 'pie_canvas'): self.pie_canvas = {}
            if not hasattr(self, 'pie_axes'): self.pie_axes = {}

            self.pie_chart_titles = [ # Ensure this is defined if not in __init__
                "1. Overall Term Type Distribution",
                "2. Linear Term Factor Contributions",
                "3. Quadratic Term Factor Contributions",
                "4. Interaction Term Factor Contributions"
            ]

            for i, title_text in enumerate(self.pie_chart_titles):
                row, col = divmod(i, 2)
                chart_container = ttk.LabelFrame(charts_frame, text=title_text, padding=10)
                chart_container.grid(row=row, column=col, sticky="nsew", padx=5, pady=5)
                charts_frame.grid_columnconfigure(col, weight=1)
                charts_frame.grid_rowconfigure(row, weight=1)

                # Frame to hold canvas (no button here anymore)
                content_frame = ttk.Frame(chart_container, style="App.TFrame")
                content_frame.pack(fill=tk.BOTH, expand=True)

                fig = Figure(figsize=(6, 4.5), dpi=100, facecolor='#F0F0F0')
                fig.subplots_adjust(left=0.05, right=0.48, top=0.92, bottom=0.08)

                ax = fig.add_subplot(111)
                canvas = FigureCanvasTkAgg(fig, master=content_frame)
                canvas_widget = canvas.get_tk_widget()
                canvas_widget.pack(fill=tk.BOTH, expand=True, side=tk.TOP)

                self.pie_figures[i] = fig
                self.pie_canvas[i] = canvas
                self.pie_axes[i] = ax
                
    def clear_state(self):
        self.clear_results_and_plots()
        self.coefficients = None
        self.bits_array = None
        self.X = None
        self.X_original_scale = None
        self.y = None
        self.y_pred = None
        self.extremum_point = None
        if hasattr(self, 'result_functions'):
            del self.result_functions
        if hasattr(self, 'comprehensive_function'):
            del self.comprehensive_function

    def select_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if not file_path:
                return

            # Clear previous state completely
            self.clear_state()

            # Rest of the method remains the same...
            self.df = pd.read_excel(file_path, header=0)
            if self.df.empty:
                raise ValueError("The file is empty")

            self.original_col_names = list(self.df.columns)
            if len(self.original_col_names) < 2:
                raise ValueError("Data must have at least 2 columns (factors + result)")

            # Create mapping
            self.col_name_mapping = {
                f"factor{i+1}": name 
                for i, name in enumerate(self.original_col_names[:-1])
            }
            self.col_name_mapping["result"] = self.original_col_names[-1]

            # Rename columns
            new_cols = list(self.col_name_mapping.keys())
            self.df.columns = new_cols

            # Process data
            self.df = self.df.apply(pd.to_numeric, errors='coerce').dropna()
            if self.df.empty:
                raise ValueError("No valid numeric data found")

            self.factor_cols = [col for col in new_cols if col.startswith('factor')]
            self.X_original_scale = self.df[self.factor_cols].values
            self.x_min_orig = self.X_original_scale.min(axis=0)
            self.x_max_orig = self.X_original_scale.max(axis=0)

            # Update UI
            self.project_name_entry.delete(0, tk.END)
            self.project_name_entry.insert(0, os.path.basename(file_path))
            self.update_table_view()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
            self.clear_results_and_plots()
                  
    def update_table_view(self):
        # Clear existing checkboxes
        for widget in self.factor_selection_frame.winfo_children():
            widget.destroy()
        
        self.factor_checkboxes = {}
        self.factor_selection = {}
        
        if self.df is None or self.df.empty:
            return
        
        # Get all column names in order (EXCLUDE 'residual')
        all_columns = [col for col in list(self.df.columns) if col != 'residual']
        
        # DEBUG: Show what columns we have
        print(f"DEBUG: Available columns: {all_columns}")
        
        # Add combobox for all features
        cb_options = ["parameter", "outcome(+)", "outcome(-)", "ignore"]
        
        # Create selection for each column
        for i, col in enumerate(all_columns):
            display_name = self.col_name_mapping.get(col, col)
            
            # Create frame for this factor's controls
            cb_frame = ttk.Frame(self.factor_selection_frame, style="App.TFrame")
            cb_frame.pack(fill='x', pady=2)
            
            # Determine default value - LAST column should be outcome, others parameters
            if i == len(all_columns) - 1:  # Last column default to outcome
                var = tk.StringVar(value="outcome(+)")
            else:
                var = tk.StringVar(value="parameter")
            
            lbl = ttk.Label(cb_frame, text=f"{display_name} ({col})", width=30)  # Show both names for debugging
            lbl.pack(side='left', padx=(0,10))
            
            cb = ttk.Combobox(cb_frame, values=cb_options, textvariable=var, width=12)
            cb.set(var.get())
            cb.pack(side='left', padx=(0,10))
            self.factor_checkboxes[col] = var

        # Update factor limits UI
        self._update_csr_factor_limits_ui()

    def clear_results_and_plots(self):
        # Clear text widgets safely
        text_widgets = [
            'equation_text', 
            'factor_definitions_text',
            'factors_text', 
            'r2_text'
        ]
        
        for widget_name in text_widgets:
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    widget.config(state='normal')
                    widget.delete(1.0, tk.END)
                    widget.config(state='disabled')
                except tk.TclError:
                    pass  # Widget might not exist anymore

        # Clear R² frame if it exists
        if hasattr(self, 'r2_frame'):
            try:
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
            except tk.TclError:
                pass

        # Clear plots
        if hasattr(self, 'figure1'):
            try:
                self.figure1.clf()
                ax = self.figure1.add_subplot(111, facecolor='white')
                ax.text(0.5, 0.5, "Load data and run fitting", 
                       ha="center", va="center", color='gray')
                self.canvas1.draw()
            except:
                pass
                
        if hasattr(self, 'figure2'):
            try:
                self.figure2.clf()
                ax = self.figure2.add_subplot(111, facecolor='white')
                ax.text(0.5, 0.5, "Load data and run fitting", 
                       ha="center", va="center", color='gray')
                self.canvas2.draw()
            except:
                pass
        
        # Clear combo boxes
        for combo_name in ['x_factor_combo', 'y_factor_combo']:
            if hasattr(self, combo_name):
                try:
                    combo = getattr(self, combo_name)
                    combo.set('')
                    combo['values'] = []
                except tk.TclError:
                    pass

    def update_max_iter_value(self, value_str):
        try:
            value = float(value_str)
            max_iter = int(10 ** value)
            self.max_iter_value_label.config(text=f"{max_iter}")
        except ValueError:
            pass

    def run_fitting(self):
        try:
            self.clear_state()
            
            if self.df is None:
                raise ValueError("Please load data first.")
                
            # Store current checkbox states before updating
            current_states = {}
            if hasattr(self, 'factor_checkboxes'):
                current_states = {col: var.get() for col, var in self.factor_checkboxes.items()}
                
            # Get selected columns
            if current_states:
                self.result_cols = [col for col, state in current_states.items() 
                                if state.startswith("outcome") and col != "residual"]
                self.factor_cols = [col for col, state in current_states.items() if state == "parameter" and col != "residual"]
            else:
                # Default to just the outcome column if no checkboxes exist yet
                self.result_cols = ["result"]
            
            if not self.result_cols:
                raise ValueError("Please select at least one result factor.")

            # DEBUG: Print what we're detecting
            print(f"DEBUG: Number of outcome columns: {len(self.result_cols)}")
            print(f"DEBUG: Outcome columns: {self.result_cols}")
            print(f"DEBUG: Factor columns: {self.factor_cols}")

            # MODIFIED: Use single result fitting for single outcome, comprehensive for multiple
            if len(self.result_cols) == 1:
                print("DEBUG: Using SINGLE result fitting")
                self._run_single_result_fitting(self.result_cols[0])
            else:
                print("DEBUG: Using COMPREHENSIVE fitting")
                self._run_comprehensive_fitting()
                
            # Update table view while preserving checkbox states
            self.update_table_view()
            
            # Restore checkbox states
            if hasattr(self, 'factor_checkboxes'):
                for col, var in self.factor_checkboxes.items():
                    if col in current_states:
                        var.set(current_states[col])
            
            self.update_3d_plot()

        except Exception as e:
            messagebox.showerror("Error", f"Fitting failed: {str(e)}")
            self.coefficients = None
            self.extremum_point = None
            self.clear_results_and_plots()

    def _run_single_result_fitting(self, result_col):
        """Handle fitting for a single result column (original behavior)"""
        try:
            # Clear any previous model state
            self.coefficients = None
            self.bits_array = None
            self.y_pred = None

            # Ensure we're working with the correct columns
            working_cols = [col for col in self.df.columns 
                        if (col.startswith('factor') and col in self.factor_cols) or col == result_col]
            
            # Create a working dataframe with just the needed columns
            working_df = self.df[working_cols].copy()

            if result_col != 'result':  
                working_df = working_df.rename(columns={result_col: 'result'})
            
            self.factor_cols = [col for col in working_df.columns if col.startswith('factor')]
            X_fit = working_df[self.factor_cols].values.copy()
            self.X_original_scale = X_fit.copy()
            self.y = working_df['result'].values
            n_factors = X_fit.shape[1]

            # Generate fresh bits array based on current number of factors
            self.bits_array = self.generate_bits_array(n_factors)
            if self.bits_array.size == 0:
                messagebox.showerror("Error", "Could not generate terms matrix (bits_array).")
                return

            # Rest of the method remains the same...
            self.norm_x_min = None
            self.norm_x_max = None

            norm_type = self.norm_select.get()
            if norm_type == "[-1, 1]":
                self.norm_x_min = X_fit.min(axis=0)
                self.norm_x_max = X_fit.max(axis=0)
                range_val = self.norm_x_max - self.norm_x_min
                range_val[range_val == 0] = 1
                X_fit = 2 * (X_fit - self.norm_x_min) / range_val - 1
            elif norm_type == "[0, 1]":
                self.norm_x_min = X_fit.min(axis=0)
                self.norm_x_max = X_fit.max(axis=0)
                range_val = self.norm_x_max - self.norm_x_min
                range_val[range_val == 0] = 1
                X_fit = (X_fit - self.norm_x_min) / range_val
            else:  # "No normalization"
                self.norm_x_min = None
                self.norm_x_max = None
                # Keep X_fit as is (original scale)

            self.X = X_fit

            X_design = self.create_design_matrix(self.X, self.bits_array)
            if X_design.shape[1] == 0:
                messagebox.showerror("Error", "Design matrix has no terms.")
                return

            alpha_val = 1e-5

            model = Ridge(alpha=alpha_val, max_iter=None, tol=1e-4, fit_intercept=False, random_state=42)
            model.fit(X_design, self.y)
            self.coefficients = model.coef_

            if model.n_iter_ is not None and model.n_iter_ >= model.max_iter:
                messagebox.showwarning("Convergence Warning",
                                    "Ridge regression may not have fully converged. Results may be suboptimal.\n"
                                    "Consider adjusting alpha or checking your data.")

            self.y_pred = model.predict(X_design)
            residuals = self.y - self.y_pred
            self.df['residual'] = 0.0  # Initialize/reset residual column
            self.df.loc[working_df.index, 'residual'] = residuals  # Only update residuals for the rows we used
            
            train_r2 = model.score(X_design, self.y)
            train_rmse = np.sqrt(np.mean((self.y - self.y_pred) ** 2))

            # SET UP OPTIMIZATION PARAMETERS FIRST
            if norm_type == "[-1, 1]":
                bounds_opt = [(-1, 1)] * n_factors
                x0_opt = np.zeros(n_factors)
            elif norm_type == "[0, 1]":
                bounds_opt = [(0, 1)] * n_factors
                x0_opt = np.full(n_factors, 0.5)
            else:  # "No normalization"
                bounds_opt = [(self.X_original_scale[:,i].min(), self.X_original_scale[:,i].max()) for i in range(n_factors)]
                x0_opt = np.mean(self.X_original_scale, axis=0)

            extremum_type_str = self.weight_combo.get().lower().replace(" ", "_")
            
            # NOW USE THE OPTIMIZATION PARAMETERS
            n_factors = len(self.factor_cols)
            all_extremum_results = []
            
            # SIMPLIFIED: Just run one optimization without artificial cardinality constraints
            print("Running optimization...")
            extremum_result = self.find_extremum(
                self.coefficients, self.bits_array, bounds_opt, x0_opt, 
                extremum_type_str, self.X)

            # Create results for all factor counts based on the single optimal solution
            if extremum_result and 'x' in extremum_result and extremum_result['x'] is not None:
                self.all_extremum_results = []
                optimal_x = extremum_result['x']
                
                # Count how many factors are significantly non-zero in the optimal solution
                significant_factors = np.sum(np.abs(optimal_x) > 1e-3)
                
                # Create results for all numbers of factors from 1 to the actual number used
                for k in range(1, significant_factors + 1):
                    # For k < significant_factors, we could create sub-optimal solutions
                    # but for now, just use the same optimal solution for all
                    self.all_extremum_results.append({
                        'active_factors': k,
                        'result': extremum_result
                    })
                
                self.extremum_point = extremum_result
                
                # Only add if we got a valid result that matches the constraint
                if extremum_result and 'x' in extremum_result and extremum_result['x'] is not None:
                    active_count = extremum_result.get('active_factors_count', k)
                    if active_count == k:  # Only add if constraint is satisfied
                        all_extremum_results.append({
                            'active_factors': k,
                            'result': extremum_result
                        })
                        print(f"Successfully added result for k={k} with {active_count} active factors")
                    else:
                        print(f"Skipping result for k={k}: expected {k} factors but got {active_count}")
                else:
                    print(f"No valid result for k={k}")

            # Store ALL results, not just the best one
            # If no cardinality-constrained results were found, create them from the regular optimization result
            if not all_extremum_results and self.extremum_point and 'x' in self.extremum_point:
                print("Creating cardinality results from regular optimization")
                extremum_x = self.extremum_point['x']
                active_count = sum(1 for val in extremum_x if abs(val) > 1e-6)
                
                # Create results for all factor counts up to the active count
                for k in range(active_count, 0, -1):
                    all_extremum_results.append({
                        'active_factors': k,
                        'result': {
                            'x': extremum_x,
                            'value': self.extremum_point['value'],
                            'active_factors_count': active_count
                        }
                    })
                
                self.all_extremum_results = all_extremum_results
                print(f"Created {len(self.all_extremum_results)} results from regular optimization")
            else:
                # Store ALL results, not just the best one
                self.all_extremum_results = all_extremum_results
                print(f"Total results stored: {len(self.all_extremum_results)}")

            # Store ALL results, not just the best one
            self.all_extremum_results = all_extremum_results
            print(f"Total cardinality-constrained results stored: {len(self.all_extremum_results)}")

            # CREATE RESULTS BASED ON USER PREFERENCE
            if self.show_all_combinations_var.get():
                # Show all combinations from N down to 1 (original behavior)
                if self.all_extremum_results:
                    # Create a set of factor counts we already have
                    existing_counts = {r['active_factors'] for r in self.all_extremum_results}
                    print(f"Existing factor counts: {sorted(existing_counts)}")
                    
                    # Create results for ALL factor counts from n_factors down to 1
                    all_factor_counts = list(range(n_factors, 0, -1))
                    print(f"Target factor counts: {all_factor_counts}")
                    
                    for k in all_factor_counts:
                        if k not in existing_counts:
                            print(f"Creating result for missing factor count: {k}")
                            
                            # Find the best available result to use as base
                            # Prefer results with higher factor counts
                            base_result = None
                            for test_k in range(k, n_factors + 1):
                                if test_k in existing_counts:
                                    base_result = next(r for r in self.all_extremum_results if r['active_factors'] == test_k)
                                    break
                            
                            # If no higher count found, use the highest available
                            if base_result is None:
                                base_result = max(self.all_extremum_results, key=lambda x: x['active_factors'])
                            
                            original_result = base_result['result']
                            original_x = original_result['x']
                            
                            # Create new solution with exactly k active factors
                            # Find the k factors with largest absolute values
                            factor_values = [(i, abs(val)) for i, val in enumerate(original_x)]
                            factor_values.sort(key=lambda x: x[1], reverse=True)
                            
                            # Keep only the top k factors, set others to zero
                            new_x = np.zeros_like(original_x)
                            for i in range(min(k, len(factor_values))):
                                idx = factor_values[i][0]
                                new_x[idx] = original_x[idx]
                            
                            # Recalculate the objective value with the new solution
                            if self.norm_select.get() == "[-1, 1]":
                                new_x_norm = 2 * (new_x - self.norm_x_min) / (self.norm_x_max - self.norm_x_min) - 1
                            elif self.norm_select.get() == "[0, 1]":
                                new_x_norm = (new_x - self.norm_x_min) / (self.norm_x_max - self.norm_x_min)
                            else:
                                new_x_norm = new_x
                            
                            design_row = self.create_design_matrix(new_x_norm.reshape(1, -1), self.bits_array)
                            new_value = np.dot(design_row[0], self.coefficients)
                            
                            new_result = {
                                'x': new_x,
                                'value': new_value,
                                'x_normalized': new_x_norm,
                                'active_factors_count': k
                            }
                            
                            self.all_extremum_results.append({
                                'active_factors': k,
                                'result': new_result
                            })
                            print(f"Created result for k={k} with value: {new_value:.4f}")
                    
                    # Re-sort all results by factor count (descending)
                    self.all_extremum_results.sort(key=lambda x: x['active_factors'], reverse=True)
                    print(f"Final results count: {len(self.all_extremum_results)}")
            else:
                # Only show the case with all factors
                if self.all_extremum_results:
                    # Filter to keep only the result with all factors
                    full_result = next((r for r in self.all_extremum_results if r['active_factors'] == n_factors), None)
                    if full_result:
                        self.all_extremum_results = [full_result]
                        print("Showing only full factor combination")
                    else:
                        # If no full result, use the one with highest factor count
                        highest_result = max(self.all_extremum_results, key=lambda x: x['active_factors'])
                        self.all_extremum_results = [highest_result]
                        print(f"Showing highest available factor combination: {highest_result['active_factors']} factors")

            # FALLBACK: If no cardinality results at all, create from regular optimization
            if not self.all_extremum_results and self.extremum_point and 'x' in self.extremum_point:
                print("No cardinality-constrained results found, creating from regular optimization")
                extremum_x = self.extremum_point['x']
                active_count = sum(1 for val in extremum_x if abs(val) > 1e-6)
                
                # Create results for all factor counts from active_count down to 1
                self.all_extremum_results = []
                for k in range(active_count, 0, -1):
                    self.all_extremum_results.append({
                        'active_factors': k,
                        'result': {
                            'x': extremum_x,
                            'value': self.extremum_point['value'],
                            'x_normalized': self.extremum_point.get('x_normalized', extremum_x),
                            'active_factors_count': active_count
                        }
                    })
                
                print(f"Created {len(self.all_extremum_results)} results from regular optimization")

            # Use the result with all factors active as the main result for plotting
            if self.all_extremum_results:
                # Find the result with all factors active for main display/plotting
                full_result = next((r for r in self.all_extremum_results if r['active_factors'] == n_factors), None)
                if full_result:
                    self.extremum_point = full_result['result']
                    print(f"Using full result with {n_factors} factors for plotting")
                else:
                    # Fallback to result with highest factor count
                    self.extremum_point = self.all_extremum_results[0]['result']
                    print(f"Using highest available result with {self.all_extremum_results[0]['active_factors']} factors for plotting")
            else:
                # Final fallback to regular optimization
                print("No results available, using regular optimization")
                self.extremum_point = self.find_extremum(self.coefficients, self.bits_array,
                                                bounds_opt, x0_opt, extremum_type_str, self.X)
        
            # Generate equation and definitions
            equation_str, function_defs_str, factor_defs_str = self.generate_equation_and_definitions(
                self.coefficients, self.bits_array, n_factors)
                        
            # Clear R² frame for single optimization
            if hasattr(self, 'r2_frame'):
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
                    
                # Create a frame for R² and RMSE display
                metrics_frame = ttk.Frame(self.r2_frame)
                metrics_frame.pack(fill='x', pady=(0,5))
                
                # R² display
                ttk.Label(metrics_frame, text="R²:", font=self.label_font).grid(row=0, column=0, sticky='w')
                r2_text = tk.Text(metrics_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                r2_text.grid(row=0, column=1, padx=(5,15), sticky='w')
                r2_text.config(state='normal')
                r2_text.insert(tk.END, f"{train_r2:.4f}")
                r2_text.config(state='disabled')
                
                # RMSE display
                ttk.Label(metrics_frame, text="RMSE:", font=self.label_font).grid(row=0, column=2, sticky='w', padx=(10,0))
                rmse_text = tk.Text(metrics_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                rmse_text.grid(row=0, column=3, padx=(5,0), sticky='w')
                rmse_text.config(state='normal')
                rmse_text.insert(tk.END, f"{train_rmse:.4f}")
                rmse_text.config(state='disabled')
                
                # Add interpretation for R²
                interpretation = ""
                if train_r2 >= 0.9:
                    interpretation = "Excellent fit"
                elif train_r2 >= 0.7:
                    interpretation = "Good fit"
                elif train_r2 >= 0.5:
                    interpretation = "Moderate fit"
                else:
                    interpretation = "Poor fit"
                
                interp_text = tk.Text(metrics_frame, height=1, width=20, wrap=tk.NONE, state='disabled',
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                interp_text.grid(row=0, column=4, padx=(15,0), sticky='w')
                interp_text.config(state='normal')
                interp_text.insert(tk.END, interpretation)
                interp_text.config(state='disabled')

            self.update_results_display(equation_str, function_defs_str, factor_defs_str, train_r2)
            display_factor_names = [self.col_name_mapping.get(f, f) for f in self.factor_cols]
            self.x_factor_combo['values'] = display_factor_names
            self.y_factor_combo['values'] = display_factor_names
            if display_factor_names:
                self.x_factor_combo.current(0)
                self.y_factor_combo.current(min(1, len(display_factor_names)-1))

            self.plot_results(n_factors, 0, min(1, n_factors -1) if n_factors > 1 else 0)
            self.update_coefficient_pie_charts()

            self.update_3d_plot()

        except Exception as e:
            raise Exception(f"Single result fitting failed: {str(e)}")
        
    def _run_comprehensive_fitting(self):
        """Handle fitting for multiple result columns (comprehensive optimization)"""
        try:
            print(f"DEBUG: Data ranges for each outcome:")
            for result_col in self.result_cols:
                data_range = self.df[result_col].max() - self.df[result_col].min()
                print(f"DEBUG: {result_col}: min={self.df[result_col].min():.2f}, max={self.df[result_col].max():.2f}, range={data_range:.2f}")
            print(f"DEBUG: Running comprehensive fitting with {len(self.result_cols)} outcomes")
            print(f"DEBUG: Outcome columns: {self.result_cols}")
            
            # FIX: Add validation for factor columns
            if not hasattr(self, 'factor_cols') or not self.factor_cols:
                raise ValueError("No factor columns available for comprehensive fitting")
                
            self.X_original_scale = self.df[self.factor_cols].values
            
            # First fit individual CSR functions for each result
            self.result_functions = {}
            result_min_max = {}
            
            # Initialize residuals column
            self.df['residual'] = 0.0
            
            self.result_functions = {}
            result_min_max = {}
            self.outcome_polarities = {}  # Store polarity for each outcome

            for result_col in self.result_cols:
                # Determine polarity from checkbox selection
                polarity = 1  # Default to positive
                if self.factor_checkboxes[result_col].get() == "outcome(-)":
                    polarity = -1
                
                # Create temporary df with this result column
                temp_df = self.df[self.factor_cols + [result_col]].copy()
                temp_df.columns = [f"factor{i+1}" for i in range(len(self.factor_cols))] + ["result"]
                
                # Store min/max for reference
                result_min_max[result_col] = {
                    'min': temp_df['result'].min(),
                    'max': temp_df['result'].max()
                }
                
                # Fit CSR for this result
                X_fit = temp_df[[f for f in temp_df.columns if f.startswith('factor')]].values
                y_fit = temp_df['result'].values
                        
                # Normalize if selected
                norm_type = self.norm_select.get()
                x_min = X_fit.min(axis=0)
                x_max = X_fit.max(axis=0)
                
                if norm_type == "[-1, 1]":
                    range_val = x_max - x_min
                    range_val[range_val == 0] = 1
                    X_fit = 2 * (X_fit - x_min) / range_val - 1
                elif norm_type == "[0, 1]":
                    range_val = x_max - x_min
                    range_val[range_val == 0] = 1
                    X_fit = (X_fit - x_min) / range_val
                
                # Generate bits array and design matrix
                n_factors = X_fit.shape[1]
                bits_array = self.generate_bits_array(n_factors)
                X_design = self.create_design_matrix(X_fit, bits_array)
                
                # Fit model
                alpha_val = 1e-5
                model = Ridge(alpha=alpha_val, fit_intercept=False)
                model.fit(X_design, y_fit)

                # Calculate predictions and residuals
                y_pred = model.predict(X_design)
                residuals = y_fit - y_pred

                # Calculate RMSE
                rmse = np.sqrt(np.mean(residuals ** 2))

                # Store the function with polarity
                self.result_functions[result_col] = {
                    'coefficients': model.coef_,
                    'bits_array': bits_array,
                    'x_min': x_min if norm_type != "No normalization" else None,
                    'x_max': x_max if norm_type != "No normalization" else None,
                    'norm_type': norm_type,
                    'model': model,
                    'X_design': X_design,
                    'y': y_fit,
                    'y_pred': y_pred,
                    'residuals': residuals,
                    'min_val': result_min_max[result_col]['min'],
                    'max_val': result_min_max[result_col]['max'],
                    'rmse': rmse,
                    'polarity': polarity
                }
                
                # Store polarity separately for display
                self.outcome_polarities[result_col] = polarity
                
                # Add residuals to main dataframe (average if multiple results)
                self.df['residual'] += residuals / len(self.result_cols)
            
            # CALCULATE INDIVIDUAL EXTREMUM VALUES
            print("DEBUG: Calculating individual extremum values...")
            for result_col, func_data in self.result_functions.items():
                # Set up optimization parameters for this outcome
                bounds_individual = [(self.df[col].min(), self.df[col].max()) for col in self.factor_cols]
                x0_individual = np.array([(min_val + max_val)/2 for min_val, max_val in bounds_individual])
                
                # Find individual extremum using improved method
                individual_extremum = self._find_individual_extremum(
                    func_data['coefficients'], 
                    func_data['bits_array'],
                    bounds_individual,
                    x0_individual,
                    self.weight_combo.get().lower().replace(" ", "_"),
                    func_data['X_design']
                )
                
                # Store the extremum value with better fallback
                if individual_extremum and 'value' in individual_extremum and not np.isnan(individual_extremum['value']):
                    func_data['extremum_val'] = individual_extremum['value']
                    print(f"DEBUG: {result_col} individual extremum: {individual_extremum['value']:.4f}")
                else:
                    # Better fallback: use data range
                    data_range = func_data['max_val'] - func_data['min_val']
                    if data_range > 1e-9:
                        if self.weight_combo.get() == "Maximum":
                            func_data['extremum_val'] = func_data['max_val']
                        elif self.weight_combo.get() == "Minimum":
                            func_data['extremum_val'] = func_data['min_val']
                        else:
                            # For absolute value objectives, use the larger absolute value
                            max_abs = max(abs(func_data['min_val']), abs(func_data['max_val']))
                            func_data['extremum_val'] = max_abs
                    else:
                        func_data['extremum_val'] = 1.0  # Default to 1 to avoid division by zero
                    print(f"DEBUG: {result_col} using fallback extremum: {func_data['extremum_val']:.4f}")
            
            # Create comprehensive function - IMPROVED NORMALIZATION
            def comprehensive_func(x):
                scores = []
                
                for result_col, func_data in self.result_functions.items():
                    # Evaluate individual CSR function
                    if func_data['norm_type'] == "[-1, 1]":
                        x_norm = 2 * (x - func_data['x_min']) / (func_data['x_max'] - func_data['x_min']) - 1
                    elif func_data['norm_type'] == "[0, 1]":
                        x_norm = (x - func_data['x_min']) / (func_data['x_max'] - func_data['x_min'])
                    else:
                        x_norm = x
                        
                    design_row = self.create_design_matrix(x_norm.reshape(1, -1), func_data['bits_array'])
                    raw_val = np.dot(design_row[0], func_data['coefficients'])
                    
                    # BOUND the predictions to reasonable ranges
                    data_min, data_max = func_data['min_val'], func_data['max_val']
                    bounded_val = np.clip(raw_val, data_min, data_max)
                    
                    # Normalize to [0,1] using ACTUAL data range
                    data_range = data_max - data_min
                    if data_range > 1e-9:
                        normalized_val = (bounded_val - data_min) / data_range
                    else:
                        normalized_val = 0.5  # Middle value if no range
                        
                    # Apply polarity and objective
                    polarity = func_data.get('polarity', 1)
                    
                    if self.weight_combo.get() == "Maximum":
                        contribution = polarity * normalized_val
                    elif self.weight_combo.get() == "Minimum":
                        contribution = polarity * (1 - normalized_val)
                    else:
                        # Handle absolute value objectives
                        contribution = polarity * normalized_val
                        
                    scores.append(contribution)
                
                # Return average score
                return np.mean(scores) if scores else 0.0
                    
            self.comprehensive_function = comprehensive_func
            
            # FIXED: Set bounds and run optimization
            bounds_opt = [(self.df[col].min(), self.df[col].max()) for col in self.factor_cols]
            x0_opt = np.array([(min_val + max_val)/2 for min_val, max_val in bounds_opt])

            extremum_type_str = 'maximum'  # Always maximize the comprehensive score
            
            # Find extremum - FIXED: Store result properly
            print("DEBUG: Finding comprehensive extremum...")
            extremum_result = self.find_extremum_comprehensive(bounds_opt, x0_opt, extremum_type_str)
            
            if extremum_result and 'x' in extremum_result and extremum_result['x'] is not None:
                self.extremum_point = extremum_result
                print(f"DEBUG: Comprehensive extremum found at: {self.extremum_point['x']}")
                print(f"DEBUG: Comprehensive extremum value: {self.extremum_point['value']}")
                
                # FIXED: Create all_extremum_results for consistent display
                n_factors = len(self.factor_cols)
                self.all_extremum_results = [{
                    'active_factors': n_factors,
                    'result': self.extremum_point
                }]
                
                # Verify the extremum
                test_value = self.comprehensive_function(self.extremum_point['x'])
                print(f"DEBUG: Verified extremum value: {test_value:.4f}")
            else:
                print("DEBUG: No valid extremum found for comprehensive optimization")
                self.extremum_point = None
                self.all_extremum_results = []

            # Update displays - FIXED: Call the proper display methods
            self.update_comprehensive_results_display(result_min_max)
            
            # Generate and display equations
            equation_str, function_defs_str, factor_defs_str = self._generate_comprehensive_equation_and_definitions()
            self.update_comprehensive_equation_display(equation_str, function_defs_str, factor_defs_str)
            
            # Set up factor combos for plotting
            display_factor_names = [self.col_name_mapping.get(f, f) for f in self.factor_cols]
            self.x_factor_combo['values'] = display_factor_names
            self.y_factor_combo['values'] = display_factor_names
            if display_factor_names:
                self.x_factor_combo.current(0)
                self.y_factor_combo.current(min(1, len(display_factor_names)-1))
            
            self.update_3d_plot()

        except Exception as e:
            import traceback
            print(f"DEBUG: Comprehensive fitting error: {str(e)}")
            print(f"DEBUG: Traceback: {traceback.format_exc()}")
            raise Exception(f"Comprehensive fitting failed: {str(e)}")
        
    def _generate_comprehensive_equation_and_definitions(self):
        """Generate equation display for comprehensive optimization"""
        if not hasattr(self, 'result_functions'):
            return "Comprehensive model (multiple outcomes)", "", ""
        
        # Generate equation strings for each outcome
        equation_parts = ["Comprehensive Optimization Model", "=" * 30, ""]
        
        for result_col, func_data in self.result_functions.items():
            n_factors = len(self.factor_cols)
            eq_str = self._generate_single_result_equation(
                func_data['coefficients'], 
                func_data['bits_array'], 
                n_factors
            )
            
            display_name = self.col_name_mapping.get(result_col, result_col)
            polarity_symbol = "(+)" if func_data.get('polarity', 1) == 1 else "(-)"
            equation_parts.append(f"{display_name} {polarity_symbol}:")
            equation_parts.append(f"  {eq_str}")
            equation_parts.append("")  # Empty line for spacing
        
        # Generate factor definitions
        factor_definitions = []
        if hasattr(self, 'factor_cols') and self.factor_cols:
            for i, factor in enumerate(self.factor_cols):
                short_name = f"c{i+1}"
                original_name = self.col_name_mapping.get(factor, f"Factor {i+1}")
                factor_definitions.append(f"{short_name} = {original_name}")
        
        equation_str = "\n".join(equation_parts)
        function_defs_str = "Multiple outcome functions combined into comprehensive score"
        factor_defs_str = "\n".join(factor_definitions) if factor_definitions else "No factor definitions available"
        
        return equation_str, function_defs_str, factor_defs_str

    def update_comprehensive_equation_display(self, equation_str, function_defs_str, factor_defs_str):
        """Update the equation and definitions display for comprehensive optimization"""
        # Update Equation Text
        if hasattr(self, 'equation_text'):
            try:
                self.equation_text.config(state='normal')
                self.equation_text.delete(1.0, tk.END)
                self.equation_text.insert(tk.END, equation_str)
                self.equation_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating equation text: {e}")

        # Update Factor Definitions Text
        if hasattr(self, 'factor_definitions_text'):
            try:
                self.factor_definitions_text.config(state='normal')
                self.factor_definitions_text.delete(1.0, tk.END)
                self.factor_definitions_text.insert(tk.END, factor_defs_str)
                self.factor_definitions_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating parameter definitions: {e}")
            
    def _find_individual_extremum(self, coefficients, bits_array, bounds, x0, extremum_type, X_context=None):
        """Find extremum for an individual outcome function with better error handling"""
        if coefficients is None or bits_array is None:
            print("DEBUG: Individual extremum - coefficients or bits_array is None")
            return None
        
        def individual_func(x_point):
            try:
                x_point_reshaped = np.array(x_point).reshape(1, -1)
                design_row = self.create_design_matrix(x_point_reshaped, bits_array)
                if design_row.shape[1] == 0: 
                    return 0
                return np.dot(design_row[0], coefficients)
            except Exception as e:
                print(f"DEBUG: Error in individual_func: {e}")
                return 0
        
        # Set up objective function
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -individual_func(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = individual_func
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(individual_func(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(individual_func(x_vals))
        else:
            objective_to_minimize = individual_func
        
        try:
            res = minimize(objective_to_minimize, x0, bounds=bounds, method='L-BFGS-B')
            
            if res.success:
                calculated_value = individual_func(res.x)
                print(f"DEBUG: Individual extremum found: {calculated_value:.4f}")
                return {
                    'x': res.x,
                    'value': calculated_value
                }
            else:
                print(f"DEBUG: Individual extremum optimization failed: {res.message}")
                
                # Try a simpler approach - evaluate at bounds and center
                test_points = [
                    x0,  # Center point
                    [b[0] for b in bounds],  # Lower bounds
                    [b[1] for b in bounds]   # Upper bounds
                ]
                
                best_value = None
                best_point = None
                
                for point in test_points:
                    try:
                        value = individual_func(point)
                        if best_value is None or (
                            (extremum_type == 'maximum' and value > best_value) or
                            (extremum_type == 'minimum' and value < best_value) or
                            (extremum_type == 'maximum_absolute_value' and abs(value) > abs(best_value)) or
                            (extremum_type == 'minimum_absolute_value' and abs(value) < abs(best_value))
                        ):
                            best_value = value
                            best_point = point
                    except:
                        continue
                
                if best_value is not None:
                    print(f"DEBUG: Using fallback extremum: {best_value:.4f}")
                    return {
                        'x': best_point,
                        'value': best_value
                    }
                    
        except Exception as e:
            print(f"Individual extremum finding failed: {str(e)}")
        
        # Final fallback: return None to use data-based fallback
        print("DEBUG: Individual extremum calculation completely failed")
        return None

    def find_extremum_comprehensive_with_active_factors(self, bounds_for_opt, x0_for_opt, extremum_type, max_active_factors):
        """Find extremum for comprehensive function with constraint on number of active factors"""
        if not hasattr(self, 'comprehensive_function'):
            return {'x': np.array([]), 'value': np.nan}

        def comprehensive_func_for_optimizer(x_point_in_opt_scale):
            return self.comprehensive_function(x_point_in_opt_scale)

        # Set up the objective function based on extremum type
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -comprehensive_func_for_optimizer(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = comprehensive_func_for_optimizer
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(comprehensive_func_for_optimizer(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(comprehensive_func_for_optimizer(x_vals))
        else:
            objective_to_minimize = comprehensive_func_for_optimizer

        # FIX: Get the actual number of factors from the data, not from x0
        if hasattr(self, 'factor_cols'):
            num_factors = len(self.factor_cols)
        else:
            num_factors = len(x0_for_opt)  # Fallback
        
        max_active = int(max_active_factors)
        
        print(f"Comprehensive optimization with max {max_active} active factors out of {num_factors} total factors")
        
        # FIX: Ensure bounds and initial guess match the expected factor count
        if len(bounds_for_opt) != num_factors:
            print(f"DEBUG: Adjusting bounds from {len(bounds_for_opt)} to {num_factors} factors")
            # Use default bounds if mismatch
            bounds_for_opt = [(self.df[col].min(), self.df[col].max()) for col in self.factor_cols]
        
        if len(x0_for_opt) != num_factors:
            print(f"DEBUG: Adjusting initial guess from {len(x0_for_opt)} to {num_factors} factors")
            x0_for_opt = np.array([(min_val + max_val)/2 for min_val, max_val in bounds_for_opt])
        
        # Add constraints from CSR limits
        constraints = []
        for limit_name, limit_data in self.csr_limits.items():
            factors = limit_data.get('factors', [])
            limit_value = limit_data.get('value', 0)
            limit_type = limit_data.get('type', 'sum')
            
            if limit_type == 'sum':
                def make_sum_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        total = sum(x_orig[i] for i in factors_list)
                        return limit_val - total  # total <= limit_val
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_sum_constraint(factors, limit_value)}
                constraints.append(constraint)
                
            elif limit_type == 'sum_equality':  # NEW: Equality constraint
                def make_sum_equality_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        total = sum(x_orig[i] for i in factors_list)
                        return total - limit_val  # total - limit_val = 0 => total = limit_val
                    return constraint_func
                
                constraint = {'type': 'eq',  # CHANGE: Use equality constraint type
                            'fun': make_sum_equality_constraint(factors, limit_value)}
                constraints.append(constraint)
                
            elif limit_type == 'product':
                def make_product_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        product = np.prod([x_orig[i] for i in factors_list])
                        return limit_val - product
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_product_constraint(factors, limit_value)}
                constraints.append(constraint)

        # Use similar iterative optimization approach as single result case
        try:
            # Initial optimization without cardinality constraint but WITH CSR limits
            if constraints:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='SLSQP', constraints=constraints, options={'disp': False, 'maxiter': 1000})
            else:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='SLSQP', options={'disp': False, 'maxiter': 1000})
            
            if res.success:
                # Get the initial solution
                x_initial = res.x
                
                # Identify which factors are most important
                factor_importance = []
                for i in range(num_factors):
                    midpoint = (bounds_for_opt[i][0] + bounds_for_opt[i][1]) / 2
                    importance = abs(x_initial[i] - midpoint) / (bounds_for_opt[i][1] - bounds_for_opt[i][0])
                    factor_importance.append(importance)
                
                factor_importance = np.array(factor_importance)
                
                # Get indices of the top 'max_active' most important factors
                top_indices = np.argsort(factor_importance)[-max_active:]
                
                # Now optimize only the selected factors, keeping others at zero
                def constrained_objective(x_active):
                    x_full = np.zeros(num_factors)
                    x_full[top_indices] = x_active
                    return objective_to_minimize(x_full)
                
                # Bounds for active factors only
                active_bounds = [bounds_for_opt[i] for i in top_indices]
                active_x0 = x_initial[top_indices]
                
                # Add constraints for the active factors optimization
                active_constraints = []
                for limit_name, limit_data in self.csr_limits.items():
                    factors = limit_data.get('factors', [])
                    limit_value = limit_data.get('value', 0)
                    limit_type = limit_data.get('type', 'sum')
                    
                    # Only include constraints that involve the active factors
                    active_factors_in_limit = [f for f in factors if f in top_indices]
                    if not active_factors_in_limit:
                        continue
                    
                    if limit_type == 'sum':
                        def make_active_sum_constraint(factors_list, limit_val, top_indices_map):
                            def constraint_func(x_active):
                                x_full = np.zeros(num_factors)
                                x_full[top_indices] = x_active
                                total = sum(x_full[i] for i in factors_list)
                                return limit_val - total
                            return constraint_func
                        
                        constraint = {'type': 'ineq', 
                                    'fun': make_active_sum_constraint(factors, limit_value, top_indices)}
                        active_constraints.append(constraint)
                
                # Optimize only the active factors with constraints
                if active_constraints:
                    res_constrained = minimize(constrained_objective, active_x0, bounds=active_bounds,
                                            method='SLSQP', constraints=active_constraints, options={'disp': False})
                else:
                    res_constrained = minimize(constrained_objective, active_x0, bounds=active_bounds,
                                            method='L-BFGS-B', options={'disp': False})
                
                if res_constrained.success:
                    # Reconstruct full solution
                    x_final = np.zeros(num_factors)
                    x_final[top_indices] = res_constrained.x
                    
                    # Calculate final value
                    extremum_value = comprehensive_func_for_optimizer(x_final)
                    
                    # Count actually active factors (above threshold)
                    active_count = np.sum(np.abs(x_final) > 1e-6)
                    
                    return {
                        'x': x_final, 
                        'value': extremum_value,
                        'active_factors_count': active_count,
                        'active_factor_indices': top_indices.tolist()
                    }
        
        except Exception as e:
            print(f"Comprehensive cardinality-constrained optimization failed: {str(e)}")
        
        # Fall back to regular comprehensive optimization
        return self.find_extremum_comprehensive(bounds_for_opt, x0_for_opt, extremum_type)

    def update_results_display(self, equation_str, function_defs_str, factor_defs_str, train_r2):
        """Update all the display widgets with the fitting results"""
        
        # Update Equation Text
        if hasattr(self, 'equation_text'):
            try:
                self.equation_text.config(state='normal')
                self.equation_text.delete(1.0, tk.END)
                self.equation_text.insert(tk.END, equation_str)
                self.equation_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating equation text: {e}")

        # Update Factor Definitions Text
        if hasattr(self, 'factor_definitions_text'):
            try:
                self.factor_definitions_text.config(state='normal')
                self.factor_definitions_text.delete(1.0, tk.END)
                self.factor_definitions_text.insert(tk.END, factor_defs_str)
                self.factor_definitions_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating parameter definitions: {e}")

        # Update Extremum Factors Display
        if hasattr(self, 'factors_text'):
            try:
                self.factors_text.config(state='normal')
                self.factors_text.delete(1.0, tk.END)
                                
                if hasattr(self, 'all_extremum_results') and self.all_extremum_results:
                    # Check if we're showing all combinations or just one
                    if len(self.all_extremum_results) > 1:
                        # Show all results with separators - NO ACTIVE FACTOR FILTERING
                        sorted_results = sorted(self.all_extremum_results, 
                                            key=lambda x: x['active_factors'], reverse=True) 

                        for result_info in sorted_results:
                            k = result_info['active_factors']
                            result = result_info['result']
                            
                            if result and 'x' in result and result['x'] is not None:
                                # First line: Number of Factors
                                self.factors_text.insert(tk.END, f"Number of Parameters: {k}\n")
                                
                                # Second line: Extremum point with ALL factors (no filtering)
                                all_factors = []
                                for i, val in enumerate(result['x']):
                                    # SHOW ALL FACTORS - no active factor filtering
                                    factor_name = self.col_name_mapping.get(
                                        self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                                        f"c{i+1}"
                                    )
                                    all_factors.append(f"{factor_name}: {val:.4f}")
                                
                                if all_factors:
                                    self.factors_text.insert(tk.END, f"Extremum point: {', '.join(all_factors)}\n")
                                else:
                                    self.factors_text.insert(tk.END, "Extremum point: No factors\n")
                                
                                # Third line: Extremum value
                                self.factors_text.insert(tk.END, f"Extremum value: {result['value']:.4f}\n")
                                
                                # Blank line between results
                                self.factors_text.insert(tk.END, "\n")
                    else:
                        # Only show the case with all factors
                        result_info = self.all_extremum_results[0]
                        k = result_info['active_factors']
                        result = result_info['result']
                        
                        # First line: Number of Factors
                        self.factors_text.insert(tk.END, f"Number of Parameters: {k}\n")
                        
                        # Second line: Extremum point with ALL factors
                        all_factors = []
                        for i, val in enumerate(result['x']):
                            # SHOW ALL FACTORS - no active factor filtering
                            factor_name = self.col_name_mapping.get(
                                self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                                f"c{i+1}"
                            )
                            all_factors.append(f"{factor_name}: {val:.4f}")
                        
                        self.factors_text.insert(tk.END, f"Extremum point: {', '.join(all_factors)}\n")
                        
                        # Third line: Extremum value
                        self.factors_text.insert(tk.END, f"Extremum value: {result['value']:.4f}\n")
                
                elif self.extremum_point and 'x' in self.extremum_point:
                    # Fallback to single result display
                    if hasattr(self, 'comprehensive_function'):
                        extremum_x = self.extremum_point['x']
                    else:
                        extremum_x = self.extremum_point['x']
                    
                    if extremum_x is not None:
                        # First line: Number of Factors - just show total count
                        total_factors = len(extremum_x)
                        self.factors_text.insert(tk.END, f"Number of Parameters: {total_factors}\n")
                        
                        # Second line: Extremum point with ALL factors
                        all_factors = []
                        for i, val in enumerate(extremum_x):
                            factor_name = self.col_name_mapping.get(
                                self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                                f"c{i+1}"
                            )
                            all_factors.append(f"{factor_name}: {val:.4f}")
                        
                        self.factors_text.insert(tk.END, f"Extremum point: {', '.join(all_factors)}\n")
                        
                        # Third line: Extremum value
                        self.factors_text.insert(tk.END, f"Extremum value: {self.extremum_point['value']:.4f}\n")
                
                self.factors_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating factors text: {e}")
                
        # Add verification info (optional - for debugging)
        verification_info = self.verify_extremum_calculation()
        print(verification_info)  # This will show in console for debugging
            
    def clear_results_and_plots(self):
        # Clear text widgets safely
        text_widgets = [
            'equation_text', 
            'factor_definitions_text',
            'factors_text', 
            'r2_text'
        ]
        
        for widget_name in text_widgets:
            if hasattr(self, widget_name):
                widget = getattr(self, widget_name)
                try:
                    widget.config(state='normal')
                    widget.delete(1.0, tk.END)
                    widget.config(state='disabled')
                except tk.TclError:
                    pass  # Widget might not exist anymore

        # Clear individual results frame if it exists
        if hasattr(self, 'individual_results_frame'):
            try:
                for widget in self.individual_results_frame.winfo_children():
                    widget.destroy()
            except tk.TclError:
                pass

        # Clear R² frame if it exists
        if hasattr(self, 'r2_frame'):
            try:
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
            except tk.TclError:
                pass

        # Clear plots
        if hasattr(self, 'figure1'):
            try:
                self.figure1.clf()
                ax = self.figure1.add_subplot(111, facecolor='white')
                ax.text(0.5, 0.5, "Load data and run fitting", 
                       ha="center", va="center", color='gray')
                self.canvas1.draw()
            except:
                pass
                
        if hasattr(self, 'figure2'):
            try:
                self.figure2.clf()
                ax = self.figure2.add_subplot(111, facecolor='white')
                ax.text(0.5, 0.5, "Load data and run fitting", 
                       ha="center", va="center", color='gray')
                self.canvas2.draw()
            except:
                pass
        
        # Clear combo boxes
        for combo_name in ['x_factor_combo', 'y_factor_combo']:
            if hasattr(self, combo_name):
                try:
                    combo = getattr(self, combo_name)
                    combo.set('')
                    combo['values'] = []
                except tk.TclError:
                    pass

    def update_comprehensive_results_display(self, result_min_max):
        """Update the display for comprehensive fitting results"""
        # Update extremum display
        if hasattr(self, 'factors_text'):
            self.factors_text.config(state='normal')
            self.factors_text.delete(1.0, tk.END)

            # FIXED: Properly display extremum for comprehensive optimization
            if hasattr(self, 'extremum_point') and self.extremum_point and 'x' in self.extremum_point:
                extremum_x = self.extremum_point['x']
                extremum_value = self.extremum_point['value']
                
                # Display basic extremum info
                self.factors_text.insert(tk.END, f"Number of Parameters: {len(self.factor_cols)}\n")
                
                # Display extremum point coordinates
                active_factors = []
                for i, val in enumerate(extremum_x):
                    factor_name = self.col_name_mapping.get(
                        self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                        f"c{i+1}"
                    )
                    active_factors.append(f"{factor_name}: {val:.4f}")
                
                self.factors_text.insert(tk.END, f"Optimal point: {', '.join(active_factors)}\n")
                self.factors_text.insert(tk.END, f"Comprehensive score: {extremum_value:.4f}\n\n")
                
                # Display individual outcome values at extremum
                if hasattr(self, 'result_functions'):
                    self.factors_text.insert(tk.END, "Individual outcomes at optimal point:\n")
                    outcome_values = []
                    for result_col, func_data in self.result_functions.items():
                        # Calculate value for this outcome at extremum
                        x_point = extremum_x
                        
                        # Normalize for each function
                        if func_data['norm_type'] == "[-1, 1]":
                            x_norm = 2 * (x_point - func_data['x_min']) / (func_data['x_max'] - func_data['x_min']) - 1
                        elif func_data['norm_type'] == "[0, 1]":
                            x_norm = (x_point - func_data['x_min']) / (func_data['x_max'] - func_data['x_min'])
                        else:
                            x_norm = x_point
                        
                        design_row = self.create_design_matrix(x_norm.reshape(1, -1), func_data['bits_array'])
                        val = np.dot(design_row[0], func_data['coefficients'])
                        
                        display_name = self.col_name_mapping.get(result_col, result_col)
                        polarity_symbol = "(+)" if func_data.get('polarity', 1) == 1 else "(-)"
                        outcome_values.append(f"{display_name}{polarity_symbol}: {val:.4f}")
                    
                    for outcome in outcome_values:
                        self.factors_text.insert(tk.END, f"  {outcome}\n")
            else:
                self.factors_text.insert(tk.END, "No extremum found for comprehensive optimization\n")

            self.factors_text.config(state='disabled')

        # FIXED: Update R² display for comprehensive optimization
        if hasattr(self, 'result_functions'):
            # Clear R² frame first
            if hasattr(self, 'r2_frame'):
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
                
                # Add R² values in a grid layout
                r2_row_num = 0
                for result_col, func_data in self.result_functions.items():
                    display_name = self.col_name_mapping.get(result_col, result_col)
                    r2 = func_data['model'].score(func_data['X_design'], func_data['y'])
                    rmse = func_data['rmse']
                    
                    r2_frame = ttk.Frame(self.r2_frame)
                    r2_frame.grid(row=r2_row_num, column=0, sticky='ew', pady=(0,5))
                    
                    ttk.Label(r2_frame, text=f"{display_name}:", font=self.label_font).grid(row=0, column=0, sticky='w')
                    
                    # R² display
                    ttk.Label(r2_frame, text="R²:", font=self.label_font).grid(row=0, column=1, sticky='w', padx=(5,0))
                    r2_text = tk.Text(r2_frame, height=1, width=12, wrap=tk.NONE, state='disabled',
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                    r2_text.grid(row=0, column=2, padx=(2,10), sticky='w')
                    r2_text.config(state='normal')
                    r2_text.insert(tk.END, f"{r2:.4f}")
                    r2_text.config(state='disabled')
                    
                    # RMSE display
                    ttk.Label(r2_frame, text="RMSE:", font=self.label_font).grid(row=0, column=3, sticky='w')
                    rmse_text = tk.Text(r2_frame, height=1, width=12, wrap=tk.NONE, state='disabled',
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                    rmse_text.grid(row=0, column=4, padx=(2,10), sticky='w')
                    rmse_text.config(state='normal')
                    rmse_text.insert(tk.END, f"{rmse:.4f}")
                    rmse_text.config(state='disabled')
                    
                    # Add interpretation for R²
                    interpretation = ""
                    if r2 >= 0.9:
                        interpretation = "Excellent fit"
                    elif r2 >= 0.7:
                        interpretation = "Good fit"
                    elif r2 >= 0.5:
                        interpretation = "Moderate fit"
                    else:
                        interpretation = "Poor fit"
                    
                    interp_text = tk.Text(r2_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                        font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                    interp_text.grid(row=0, column=5, padx=(5,0), sticky='w')
                    interp_text.config(state='normal')
                    interp_text.insert(tk.END, interpretation)
                    interp_text.config(state='disabled')
                    
                    r2_row_num += 1

                # Calculate and display average R² and RMSE
                if self.result_functions:
                    avg_r2 = np.mean([func['model'].score(func['X_design'], func['y']) for func in self.result_functions.values()])
                    avg_rmse = np.mean([func['rmse'] for func in self.result_functions.values()])

                    avg_frame = ttk.Frame(self.r2_frame)
                    avg_frame.grid(row=r2_row_num, column=0, sticky='ew', pady=(10,0))

                    ttk.Label(avg_frame, text="Average R²:", font=(self.label_font.cget("family"), 
                            self.label_font.cget("size"), "bold")).grid(row=0, column=0, sticky='w')

                    avg_r2_text = tk.Text(avg_frame, height=1, width=12, wrap=tk.NONE, state='disabled',
                                        font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                    avg_r2_text.grid(row=0, column=1, padx=(5,15), sticky='w')
                    avg_r2_text.config(state='normal')
                    avg_r2_text.insert(tk.END, f"{avg_r2:.4f}")
                    avg_r2_text.config(state='disabled')

                    # Average RMSE display
                    ttk.Label(avg_frame, text="Average RMSE:", font=(self.label_font.cget("family"), 
                            self.label_font.cget("size"), "bold")).grid(row=0, column=2, sticky='w', padx=(10,0))

                    avg_rmse_text = tk.Text(avg_frame, height=1, width=12, wrap=tk.NONE, state='disabled',
                                        font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                    avg_rmse_text.grid(row=0, column=3, padx=(5,0), sticky='w')
                    avg_rmse_text.config(state='normal')
                    avg_rmse_text.insert(tk.END, f"{avg_rmse:.4f}")
                    avg_rmse_text.config(state='disabled')

    def _generate_single_result_equation(self, beta, bits_array, n_factors):
        """Helper method to generate equation string for a single result (same as single optimization)"""
        if beta is None or bits_array is None or bits_array.size == 0 or len(beta) != bits_array.shape[0]:
            return "0"

        # Initialize lists to store terms of each type
        constant_terms = []
        linear_terms = []
        quadratic_terms = []
        interaction_terms = []

        # Process each term and classify it
        for i, (coef, bits) in enumerate(zip(beta, bits_array)):
            if abs(coef) < 1e-7:
                continue  # Skip negligible terms

            term_type = self._classify_term(bits)
            term_str = self._format_term(coef, bits, n_factors)

            if term_type == 'constant':
                constant_terms.append(term_str)
            elif term_type == 'linear':
                # Store with factor index for sorting (f1, f2, f3,...)
                factor_idx = np.where(bits == 1)[0][0]
                linear_terms.append((factor_idx, term_str))
            elif term_type == 'quadratic':
                # Store with factor index for sorting (f1², f2², f3²,...)
                factor_idx = np.where(bits == 2)[0][0]
                quadratic_terms.append((factor_idx, term_str))
            elif term_type == 'interaction':
                # Store with factor indices for sorting (f1×f2, f1×f3,..., f2×f3, f2×f4,...)
                idxs = np.where(bits == 1)[0]
                interaction_terms.append((idxs[0], idxs[1], term_str))

        # Sort terms according to specified ordering
        # Linear terms: f1, f2, f3,...
        linear_terms.sort(key=lambda x: x[0])
        
        # Quadratic terms: f1², f2², f3²,...
        quadratic_terms.sort(key=lambda x: x[0])
        
        # Interaction terms: f1×f2, f1×f3,..., f2×f3, f2×f4,...
        interaction_terms.sort(key=lambda x: (x[0], x[1]))

        # Combine all terms in the specified order
        equation_terms = []
        
        # Add constant term if present
        if constant_terms:
            equation_terms.append(constant_terms[0])  # There should be only one constant term

        # Add linear terms (f1, f2, f3,...)
        equation_terms.extend([t[1] for t in linear_terms])
        
        # Add quadratic terms (f1², f2², f3²,...)
        equation_terms.extend([t[1] for t in quadratic_terms])
        
        # Add interaction terms (f1×f2, f1×f3,..., f2×f3, f2×f4,...)
        equation_terms.extend([t[2] for t in interaction_terms])

        # Construct the final equation string
        if not equation_terms:
            return "0"
        else:
            # Join terms with appropriate signs
            equation_str = " ".join(equation_terms).replace(" + -", " - ")
            # Handle case where first term is negative
            if equation_str.startswith("- "):
                return equation_str[2:]  # Remove the "- " since we'll add it in the display
            return equation_str
        
    def update_3d_plot(self):
        if self.df is None or not self.factor_cols:
            return

        x_factor_display_name = self.x_factor_combo.get()
        y_factor_display_name = self.y_factor_combo.get()

        if not x_factor_display_name or not y_factor_display_name:
            return

        try:
            # Find internal factor names from display names
            x_internal_name = None
            y_internal_name = None
            
            for internal_name, display_name in self.col_name_mapping.items():
                if display_name == x_factor_display_name and internal_name.startswith("factor"):
                    x_internal_name = internal_name
                if display_name == y_factor_display_name and internal_name.startswith("factor"):
                    y_internal_name = internal_name

            if not x_internal_name or not y_internal_name:
                messagebox.showerror("Error", "Could not map selected display factor names to internal factor names.")
                return

            x_idx = self.factor_cols.index(x_internal_name)
            y_idx = self.factor_cols.index(y_internal_name)
            n_factors = len(self.factor_cols)

            # Clear the figure and prepare for new plot
            self.figure2.clf()
            self.figure2.subplots_adjust(bottom=0.18, left=0.1, right=0.85, top=0.9)
            self.figure2.patch.set_facecolor('#F0F0F0')
            ax2 = self.figure2.add_subplot(111, projection='3d', facecolor='#ffffff')
            
            # Configure the 3D plot appearance
            ax2.grid(True, linestyle=':', alpha=0.5)
            for pane_ax in [ax2.xaxis, ax2.yaxis, ax2.zaxis]: 
                pane_ax.set_pane_color((1.0, 1.0, 1.0, 0.0))
                pane_ax.pane.set_edgecolor('#D0D0D0')

            # Get the ranges for the selected factors
            x1_orig_plot_axis = np.linspace(self.X_original_scale[:, x_idx].min(), 
                                        self.X_original_scale[:, x_idx].max(), 20)
            x2_orig_plot_axis = np.linspace(self.X_original_scale[:, y_idx].min(), 
                                        self.X_original_scale[:, y_idx].max(), 20)
            x1_grid_orig, x2_grid_orig = np.meshgrid(x1_orig_plot_axis, x2_orig_plot_axis)
            z_csr_values_grid = np.zeros_like(x1_grid_orig)

            # Get fixed values for other factors (use extremum point if available, otherwise mean)
            fixed_factor_values_original_scale = np.mean(self.X_original_scale, axis=0)
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                if hasattr(self, 'comprehensive_function'):
                    fixed_factor_values_original_scale = self.extremum_point['x']
                else:
                    unnorm_temp = self._unnormalize_point(self.extremum_point.get('x_normalized', self.extremum_point['x']))
                    if unnorm_temp is not None:
                        fixed_factor_values_original_scale = unnorm_temp

            # Evaluate the CSR function across the grid
            for i_grid in range(x1_grid_orig.shape[0]):
                for j_grid in range(x1_grid_orig.shape[1]):
                    current_full_point_orig_scale = fixed_factor_values_original_scale.copy()
                    current_full_point_orig_scale[x_idx] = x1_grid_orig[i_grid, j_grid]
                    current_full_point_orig_scale[y_idx] = x2_grid_orig[i_grid, j_grid]
                    
                    if hasattr(self, 'comprehensive_function'):
                        # For comprehensive optimization
                        z_csr_values_grid[i_grid, j_grid] = self.comprehensive_function(current_full_point_orig_scale)
                    elif hasattr(self, 'coefficients') and hasattr(self, 'bits_array'):
                        # For single result optimization
                        current_full_point_norm_scale = self._normalize_point(current_full_point_orig_scale)
                        if current_full_point_norm_scale is not None:
                            z_csr_values_grid[i_grid, j_grid] = self.evaluate_csr_at_point(current_full_point_norm_scale)
                        else:
                            z_csr_values_grid[i_grid, j_grid] = np.nan
                    else:
                        z_csr_values_grid[i_grid, j_grid] = np.nan

            # Plot the surface
            surf = ax2.plot_surface(x1_grid_orig, x2_grid_orig, z_csr_values_grid, 
                                cmap='viridis', alpha=0.9, edgecolor='#555555', 
                                linewidth=0.15, antialiased=True)

            # FIXED: Plot extremum point if available - REPLACE THIS SECTION
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                
                # FIXED: Handle both comprehensive and single result optimization
                if hasattr(self, 'comprehensive_function'):
                    # For comprehensive optimization, we already have original scale coordinates
                    extremum_x_orig = self.extremum_point['x']
                    extremum_value = self.extremum_point['value']
                else:
                    # For single result, unnormalize if needed
                    extremum_x_orig = self.extremum_point.get('x', None)
                    extremum_value = self.extremum_point.get('value', 0)
                    # If we have normalized coordinates, unnormalize them
                    if 'x_normalized' in self.extremum_point:
                        extremum_x_orig = self._unnormalize_point(self.extremum_point['x_normalized'])
                
                if extremum_x_orig is not None and len(extremum_x_orig) > 0:
                    print(f"DEBUG: Plotting extremum at ({extremum_x_orig[x_idx]:.4f}, {extremum_x_orig[y_idx]:.4f}, {extremum_value:.4f})")
                    ax2.scatter([extremum_x_orig[x_idx]], 
                            [extremum_x_orig[y_idx]], 
                            [extremum_value], 
                            c='gold', s=200, marker='*', edgecolor='black', 
                            linewidth=1, label='Optimal Point', depthshade=True, zorder=10)
                    
                    # Add legend
                    ax2.legend(fontsize=8, facecolor='#F0F0F0', framealpha=0.8)

            # Set axis labels
            x_axis_name = self.col_name_mapping.get(self.factor_cols[x_idx], self.factor_cols[x_idx])
            y_axis_name = self.col_name_mapping.get(self.factor_cols[y_idx], self.factor_cols[y_idx])
            result_axis_name = "Combined Result" if hasattr(self, 'comprehensive_function') else self.col_name_mapping.get("result", "Result")
            
            ax2.set_xlabel(f"\n{x_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_ylabel(f"\n{y_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_zlabel(f"\n{result_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_title('CSR Response Surface', fontsize=11, fontweight='bold', y=1.02)

            # Add colorbar
            cbar = self.figure2.colorbar(surf, ax=ax2, shrink=0.6, aspect=12, pad=0.15, format="%.2f")
            cbar.ax.tick_params(labelsize=8)
            cbar.outline.set_edgecolor('gray')
            
            # Set view angle
            ax2.view_init(elev=28, azim=130)
            
            # Adjust tick parameters
            for axis_obj in [ax2.xaxis, ax2.yaxis, ax2.zaxis]:
                axis_obj.set_tick_params(pad=3, labelsize=8)
                axis_obj.label.set_size(9)

            self.canvas2.draw()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update 3D plot: {str(e)}")
            # Print the full traceback to help with debugging
            import traceback
            traceback.print_exc()

    def generate_bits_array(self, n_factors):
        if n_factors <= 0:
            return np.array([])
        
        bits = []
        # Constant term (all zeros)
        bits.append(np.zeros(n_factors, dtype=int))
        
        # Linear terms (f1, f2, f3,...)
        for i in range(n_factors):
            term = np.zeros(n_factors, dtype=int)
            term[i] = 1
            bits.append(term)
        
        # Quadratic terms (f1², f2², f3²,...)
        for i in range(n_factors):
            term = np.zeros(n_factors, dtype=int)
            term[i] = 2
            bits.append(term)
        
        # Interaction terms (f1*f2, f1*f3,..., f2*f3, f2*f4,...)
        if n_factors > 1:
            for i in range(n_factors):
                for j in range(i + 1, n_factors):
                    term = np.zeros(n_factors, dtype=int)
                    term[i] = 1
                    term[j] = 1
                    bits.append(term)
        
        # Convert to numpy array (no need to sort, order is enforced above)
        return np.array(bits, dtype=int)

    def create_design_matrix(self, X_input_scaled, bits_array):
        if X_input_scaled.size == 0 or bits_array.size == 0:
            return np.array([[]])
        n_samples = X_input_scaled.shape[0]
        n_terms = bits_array.shape[0]
        if n_terms == 0:
            return np.array([[]])
        X_design = np.ones((n_samples, n_terms))
        for term_idx, bits_row in enumerate(bits_array):
            current_term_values_for_all_samples = np.ones(n_samples)
            for factor_idx, power in enumerate(bits_row):
                if power == 1:
                    current_term_values_for_all_samples *= X_input_scaled[:, factor_idx]
                elif power == 2:
                    current_term_values_for_all_samples *= X_input_scaled[:, factor_idx]**2
                elif power > 2:
                    current_term_values_for_all_samples *= X_input_scaled[:, factor_idx]**power
            X_design[:, term_idx] = current_term_values_for_all_samples
        return X_design
            
    def find_extremum_with_active_factors(self, beta, bits_array, bounds_for_opt, x0_for_opt, extremum_type, X_context_for_opt, max_active_factors):
        """Simplified version - just run regular optimization without artificial constraints"""
        if beta is None or bits_array is None or X_context_for_opt is None:
            return {'x': np.array([]), 'value': np.nan}
        
        if len(beta) != bits_array.shape[0]:
            messagebox.showerror("Error in find_extremum", f"Coefficient count ({len(beta)}) doesn't match terms count ({bits_array.shape[0]}).")
            return {'x': np.array([]), 'value': np.nan}

        def csr_func_for_optimizer(x_point_in_opt_scale):
            x_point_reshaped = np.array(x_point_in_opt_scale).reshape(1, -1)
            design_row = self.create_design_matrix(x_point_reshaped, bits_array)
            if design_row.shape[1] == 0: return 0
            return np.dot(design_row[0], beta)

        # Set up objective function
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -csr_func_for_optimizer(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = csr_func_for_optimizer
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(csr_func_for_optimizer(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(csr_func_for_optimizer(x_vals))
        else:
            objective_to_minimize = csr_func_for_optimizer

        # SIMPLIFIED: Just run the optimization without artificial cardinality constraints
        # The CSR limits will naturally constrain the solution space
        
        # Add constraints from CSR limits
        constraints = []
        for limit_name, limit_data in self.csr_limits.items():
            factors = limit_data.get('factors', [])
            limit_value = limit_data.get('value', 0)
            limit_type = limit_data.get('type', 'sum')
            
            if limit_type == 'sum':
                def make_sum_constraint(factors_list, limit_val, norm_min, norm_max, norm_type):
                    def constraint_func(x_norm):
                        # Convert normalized values back to original scale for constraint
                        if norm_type == "[-1, 1]":
                            x_orig = (x_norm + 1) / 2 * (norm_max - norm_min) + norm_min
                        elif norm_type == "[0, 1]":
                            x_orig = x_norm * (norm_max - norm_min) + norm_min
                        else:
                            x_orig = x_norm
                        
                        total = sum(x_orig[i] for i in factors_list)
                        return limit_val - total  # total <= limit_val
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_sum_constraint(factors, limit_value, 
                                                    self.norm_x_min, self.norm_x_max, 
                                                    self.norm_select.get())}
                constraints.append(constraint)
                
            elif limit_type == 'sum_equality':  # NEW: Equality constraint
                def make_sum_equality_constraint(factors_list, limit_val, norm_min, norm_max, norm_type):
                    def constraint_func(x_norm):
                        # Convert normalized values back to original scale for constraint
                        if norm_type == "[-1, 1]":
                            x_orig = (x_norm + 1) / 2 * (norm_max - norm_min) + norm_min
                        elif norm_type == "[0, 1]":
                            x_orig = x_norm * (norm_max - norm_min) + norm_min
                        else:
                            x_orig = x_norm
                        
                        total = sum(x_orig[i] for i in factors_list)
                        return total - limit_val  # total - limit_val = 0 => total = limit_val
                    return constraint_func
                
                constraint = {'type': 'eq',  # CHANGE: Use equality constraint type
                            'fun': make_sum_equality_constraint(factors, limit_value, 
                                                            self.norm_x_min, self.norm_x_max, 
                                                            self.norm_select.get())}
                constraints.append(constraint)

        # Perform optimization with CSR constraints only
        try:
            if constraints:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='SLSQP', constraints=constraints, options={'disp': False})
            else:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='L-BFGS-B', options={'disp': False})
            
            if res.success:
                # Convert to original scale
                if self.norm_select.get() == "[-1, 1]":
                    res_orig = (res.x + 1) / 2 * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                elif self.norm_select.get() == "[0, 1]":
                    res_orig = res.x * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                else:
                    res_orig = res.x
                
                # Calculate extremum value using normalized inputs directly
                design_row_normalized = self.create_design_matrix(res.x.reshape(1, -1), bits_array)
                extremum_value = np.dot(design_row_normalized[0], beta)
                
                return {
                    'x': res_orig, 
                    'value': extremum_value, 
                    'x_normalized': res.x
                }
        
        except Exception as e:
            print(f"Optimization failed: {str(e)}")
        
        return None

    def _find_extremum_heuristic(self, beta, bits_array, bounds_for_opt, x0_for_opt, extremum_type, X_context_for_opt, max_active_factors):
        """Heuristic approach for cardinality-constrained optimization for large factor counts"""
        print(f"Using improved heuristic approach for {max_active_factors} active factors")
        
        def csr_func_for_optimizer(x_point_in_opt_scale):
            x_point_reshaped = np.array(x_point_in_opt_scale).reshape(1, -1)
            design_row = self.create_design_matrix(x_point_reshaped, bits_array)
            if design_row.shape[1] == 0: return 0
            return np.dot(design_row[0], beta)

        # Set up objective function
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -csr_func_for_optimizer(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = csr_func_for_optimizer
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(csr_func_for_optimizer(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(csr_func_for_optimizer(x_vals))
        else:
            objective_to_minimize = csr_func_for_optimizer

        num_factors = X_context_for_opt.shape[1]
        max_active = int(max_active_factors)
        
        # NEW APPROACH: Use L1 regularization to encourage sparsity, then select top factors
        try:
            # First optimization with L1-like constraint to encourage sparsity
            def sparsity_objective(x):
                main_obj = objective_to_minimize(x)
                # Add small L1 penalty to encourage zeros
                l1_penalty = 1e-6 * np.sum(np.abs(x))
                return main_obj + l1_penalty
            
            res_sparse = minimize(sparsity_objective, x0_for_opt, bounds=bounds_for_opt, 
                                method='L-BFGS-B', options={'disp': False})
            
            if res_sparse.success:
                x_sparse = res_sparse.x
                
                # Find the factors with largest absolute values (most important)
                abs_values = np.abs(x_sparse)
                
                # If we want exactly max_active non-zero factors, set others to zero
                if max_active < num_factors:
                    # Get indices of top max_active factors
                    threshold = np.sort(abs_values)[-max_active]
                    top_indices = np.where(abs_values >= threshold)[0]
                    
                    # Now optimize only these factors, allowing others to be zero
                    def constrained_objective(x_active):
                        x_full = np.zeros(num_factors)
                        x_full[top_indices] = x_active
                        return objective_to_minimize(x_full)
                    
                    active_bounds = [bounds_for_opt[i] for i in top_indices]
                    active_x0 = x_sparse[top_indices]
                    
                    # Add CSR constraints for active factors
                    active_constraints = []
                    for limit_name, limit_data in self.csr_limits.items():
                        factors = limit_data.get('factors', [])
                        limit_value = limit_data.get('value', 0)
                        limit_type = limit_data.get('type', 'sum')
                        
                        active_factors_in_limit = [f for f in factors if f in top_indices]
                        if not active_factors_in_limit:
                            continue
                        
                        if limit_type == 'sum':
                            def make_active_sum_constraint(factors_list, limit_val, norm_min, norm_max, norm_type, top_indices_list):
                                def constraint_func(x_active):
                                    x_full = np.zeros(num_factors)
                                    x_full[top_indices_list] = x_active
                                    
                                    # Convert to original scale for constraint
                                    if norm_type == "[-1, 1]":
                                        x_orig = (x_full + 1) / 2 * (norm_max - norm_min) + norm_min
                                    elif norm_type == "[0, 1]":
                                        x_orig = x_full * (norm_max - norm_min) + norm_min
                                    else:
                                        x_orig = x_full
                                    
                                    total = sum(x_orig[i] for i in factors_list)
                                    return limit_val - total
                                return constraint_func
                            
                            constraint = {'type': 'ineq', 
                                        'fun': make_active_sum_constraint(factors, limit_value, 
                                                                        self.norm_x_min, self.norm_x_max, 
                                                                        self.norm_select.get(), top_indices)}
                            active_constraints.append(constraint)
                    
                    # Optimize the active factors
                    if active_constraints:
                        res_final = minimize(constrained_objective, active_x0, bounds=active_bounds,
                                        method='SLSQP', constraints=active_constraints, options={'disp': False})
                    else:
                        res_final = minimize(constrained_objective, active_x0, bounds=active_bounds,
                                        method='L-BFGS-B', options={'disp': False})
                    
                    if res_final.success:
                        x_final = np.zeros(num_factors)
                        x_final[top_indices] = res_final.x
                    else:
                        # Fallback: use the sparse solution
                        x_final = x_sparse
                else:
                    # Use all factors
                    x_final = x_sparse
                
                # Calculate final value
                value = csr_func_for_optimizer(x_final)
                
                # Count actually active factors (above threshold)
                active_count = np.sum(np.abs(x_final) > 1e-6)
                
                # Convert to original scale
                if self.norm_select.get() == "[-1, 1]":
                    res_orig = (x_final + 1) / 2 * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                elif self.norm_select.get() == "[0, 1]":
                    res_orig = x_final * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                else:
                    res_orig = x_final
                
                return {
                    'x': res_orig, 
                    'value': value,
                    'x_normalized': x_final,
                    'active_factors_count': active_count,
                    'active_factor_indices': list(top_indices) if max_active < num_factors else list(range(num_factors))
                }
        
        except Exception as e:
            print(f"Improved heuristic optimization failed: {str(e)}")
        
        return None

    # Modified optimization methods to use CSR limits
    def find_extremum(self, beta, bits_array, bounds_for_opt, x0_for_opt, extremum_type, X_context_for_opt):
        if beta is None or bits_array is None or X_context_for_opt is None:
            return {'x': np.array([]), 'value': np.nan}
        if len(beta) != bits_array.shape[0]:
            messagebox.showerror("Error in find_extremum", f"Coefficient count ({len(beta)}) doesn't match terms count ({bits_array.shape[0]}).")
            return {'x': np.array([]), 'value': np.nan}

        def csr_func_for_optimizer(x_point_in_opt_scale):
            x_point_reshaped = np.array(x_point_in_opt_scale).reshape(1, -1)
            design_row = self.create_design_matrix(x_point_reshaped, bits_array)
            if design_row.shape[1] == 0: return 0
            return np.dot(design_row[0], beta)

        objective_func_val = csr_func_for_optimizer
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -csr_func_for_optimizer(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = csr_func_for_optimizer
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(csr_func_for_optimizer(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(csr_func_for_optimizer(x_vals))
        else:
            objective_to_minimize = csr_func_for_optimizer

        num_factors_in_context = X_context_for_opt.shape[1]
        if len(x0_for_opt) != num_factors_in_context:
            messagebox.showerror("Optimization Error", f"Initial guess x0 length ({len(x0_for_opt)}) mismatch with factor count ({num_factors_in_context}).")
            return {'x': np.array([np.nan]*num_factors_in_context), 'value': np.nan}
        if len(bounds_for_opt) != num_factors_in_context:
            messagebox.showerror("Optimization Error", f"Bounds length ({len(bounds_for_opt)}) mismatch with factor count ({num_factors_in_context}).")
            return {'x': np.array([np.nan]*num_factors_in_context), 'value': np.nan}

        # Add constraints from CSR limits - APPLY IN ORIGINAL SCALE
        constraints = []
        for limit_name, limit_data in self.csr_limits.items():
            factors = limit_data.get('factors', [])
            limit_value = limit_data.get('value', 0)
            limit_type = limit_data.get('type', 'sum')
            
            if limit_type == 'sum':
                def make_sum_constraint(factors_list, limit_val, norm_min, norm_max, norm_type):
                    def constraint_func(x_norm):
                        # Convert normalized values back to original scale for constraint
                        if norm_type == "[-1, 1]":
                            x_orig = (x_norm + 1) / 2 * (norm_max - norm_min) + norm_min
                        elif norm_type == "[0, 1]":
                            x_orig = x_norm * (norm_max - norm_min) + norm_min
                        else:
                            x_orig = x_norm
                        
                        total = sum(x_orig[i] for i in factors_list)
                        return limit_val - total  # total <= limit_val
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_sum_constraint(factors, limit_value, 
                                                    self.norm_x_min, self.norm_x_max, 
                                                    self.norm_select.get())}
                constraints.append(constraint)
                
            elif limit_type == 'sum_equality':  # NEW: Equality constraint
                def make_sum_equality_constraint(factors_list, limit_val, norm_min, norm_max, norm_type):
                    def constraint_func(x_norm):
                        # Convert normalized values back to original scale for constraint
                        if norm_type == "[-1, 1]":
                            x_orig = (x_norm + 1) / 2 * (norm_max - norm_min) + norm_min
                        elif norm_type == "[0, 1]":
                            x_orig = x_norm * (norm_max - norm_min) + norm_min
                        else:
                            x_orig = x_norm
                        
                        total = sum(x_orig[i] for i in factors_list)
                        return total - limit_val  # total - limit_val = 0 => total = limit_val
                    return constraint_func
                
                constraint = {'type': 'eq',  # CHANGE: Use equality constraint type
                            'fun': make_sum_equality_constraint(factors, limit_value, 
                                                            self.norm_x_min, self.norm_x_max, 
                                                            self.norm_select.get())}
                constraints.append(constraint)
                
            elif limit_type == 'product':
                def make_product_constraint(factors_list, limit_val, norm_min, norm_max, norm_type):
                    def constraint_func(x_norm):
                        # Convert normalized values back to original scale for constraint
                        if norm_type == "[-1, 1]":
                            x_orig = (x_norm + 1) / 2 * (norm_max - norm_min) + norm_min
                        elif norm_type == "[0, 1]":
                            x_orig = x_norm * (norm_max - norm_min) + norm_min
                        else:
                            x_orig = x_norm
                        
                        product = np.prod([x_orig[i] for i in factors_list])
                        return limit_val - product
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_product_constraint(factors, limit_value,
                                                        self.norm_x_min, self.norm_x_max,
                                                        self.norm_select.get())}
                constraints.append(constraint)

        # Debug: Print constraints before optimization
        print(f"Number of constraints: {len(constraints)}")
        print(f"Normalization type: {self.norm_select.get()}")
        print(f"Norm min: {self.norm_x_min}")
        print(f"Norm max: {self.norm_x_max}")

        # Perform optimization with constraints if any
        try:
            if constraints:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='SLSQP', constraints=constraints, options={'disp': True})
            else:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, method='L-BFGS-B')
            
            # Check if optimization was successful
            if not res.success:
                pass
            
            # Verify constraints are satisfied in ORIGINAL SCALE
            if constraints:
                print("Verifying constraints in ORIGINAL SCALE:")
                # Convert result to original scale for verification
                if self.norm_select.get() == "[-1, 1]":
                    res_orig = (res.x + 1) / 2 * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                elif self.norm_select.get() == "[0, 1]":
                    res_orig = res.x * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
                else:
                    res_orig = res.x
                    
                for i, constr in enumerate(constraints):
                    constraint_value = constr['fun'](res.x)  # This now uses original scale conversion
                    print(f"Constraint {i}: {constraint_value} (should be >= 0)")
                    if constraint_value < -1e-3:  # Allow small numerical tolerance
                        messagebox.showwarning("Constraint Violation", 
                                            f"Constraint {i} is violated: {constraint_value}")
            
            # Calculate the final value - USE NORMALIZED VALUES DIRECTLY
            if self.norm_select.get() == "[-1, 1]":
                res_orig = (res.x + 1) / 2 * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
            elif self.norm_select.get() == "[0, 1]":
                res_orig = res.x * (self.norm_x_max - self.norm_x_min) + self.norm_x_min
            else:
                res_orig = res.x
            
            # CRITICAL FIX: Calculate extremum value using normalized inputs directly
            design_row_normalized = self.create_design_matrix(res.x.reshape(1, -1), bits_array)
            extremum_value = np.dot(design_row_normalized[0], beta)
            
            # Debug: Print the calculation details
            print(f"Optimization result (normalized): {res.x}")
            print(f"Optimization result (original): {res_orig}")
            print(f"Design row (normalized): {design_row_normalized[0]}")
            print(f"Coefficients: {beta}")
            print(f"Extremum value (calculated): {extremum_value}")
            print(f"Sum of factors (original): {sum(res_orig)}")
            
            return {'x': res_orig, 'value': extremum_value, 'x_normalized': res.x}   
                 
        except Exception as e:
            messagebox.showerror("Optimization Error", f"Optimization failed: {str(e)}")
            return {'x': np.array([np.nan]*num_factors_in_context), 'value': np.nan}
    
    def find_extremum_comprehensive(self, bounds_for_opt, x0_for_opt, extremum_type):
        """Find extremum for the comprehensive function that combines multiple results"""
        if not hasattr(self, 'comprehensive_function'):
            return {'x': np.array([]), 'value': np.nan}

        def comprehensive_func_for_optimizer(x_point_in_opt_scale):
            return self.comprehensive_function(x_point_in_opt_scale)

        # Set up the objective function based on extremum type
        if extremum_type == 'maximum':
            objective_to_minimize = lambda x_vals: -comprehensive_func_for_optimizer(x_vals)
        elif extremum_type == 'minimum':
            objective_to_minimize = comprehensive_func_for_optimizer
        elif extremum_type == 'maximum_absolute_value':
            objective_to_minimize = lambda x_vals: -np.abs(comprehensive_func_for_optimizer(x_vals))
        elif extremum_type == 'minimum_absolute_value':
            objective_to_minimize = lambda x_vals: np.abs(comprehensive_func_for_optimizer(x_vals))
        else:
            objective_to_minimize = comprehensive_func_for_optimizer

        # Add constraints from CSR limits - APPLY IN ORIGINAL SCALE
        constraints = []
        for limit_name, limit_data in self.csr_limits.items():
            factors = limit_data.get('factors', [])
            limit_value = limit_data.get('value', 0)
            limit_type = limit_data.get('type', 'sum')
            
            if limit_type == 'sum':
                def make_sum_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        # Note: For comprehensive optimization, we're already in original scale
                        total = sum(x_orig[i] for i in factors_list)
                        return limit_val - total  # total <= limit_val
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_sum_constraint(factors, limit_value)}
                constraints.append(constraint)
                
            elif limit_type == 'sum_equality':  # NEW: Equality constraint
                def make_sum_equality_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        # Note: For comprehensive optimization, we're already in original scale
                        total = sum(x_orig[i] for i in factors_list)
                        return total - limit_val  # total - limit_val = 0 => total = limit_val
                    return constraint_func
                
                constraint = {'type': 'eq',  # CHANGE: Use equality constraint type
                            'fun': make_sum_equality_constraint(factors, limit_value)}
                constraints.append(constraint)
                
            elif limit_type == 'product':
                def make_product_constraint(factors_list, limit_val):
                    def constraint_func(x_orig):
                        # Note: For comprehensive optimization, we're already in original scale
                        product = np.prod([x_orig[i] for i in factors_list])
                        return limit_val - product
                    return constraint_func
                
                constraint = {'type': 'ineq', 
                            'fun': make_product_constraint(factors, limit_value)}
                constraints.append(constraint)

        # Debug: Print constraints before optimization
        print(f"Number of constraints: {len(constraints)}")
        print("Comprehensive optimization - using original scale constraints")

        # Perform the optimization - note bounds are in original scale
        try:
            if constraints:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, 
                            method='SLSQP', constraints=constraints, options={'disp': True})
            else:
                res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, method='L-BFGS-B')
            
            # Check if optimization was successful
            if not res.success:
                messagebox.showwarning("Optimization Warning", 
                                    f"Optimization may not have converged: {res.message}")
            
            # Verify constraints are satisfied
            if constraints:
                print("Verifying constraints:")
                for i, constr in enumerate(constraints):
                    constraint_value = constr['fun'](res.x)
                    print(f"Constraint {i}: {constraint_value} (should be >= 0)")
                    if constraint_value < -1e-3:  # Allow small numerical tolerance
                        messagebox.showwarning("Constraint Violation", 
                                            f"Constraint {i} is violated: {constraint_value}")
            
            # CRITICAL FIX: Always calculate using the comprehensive function
            extremum_value = comprehensive_func_for_optimizer(res.x)
            
            # Debug: Print the result
            print(f"Optimization result (original): {res.x}")
            print(f"Sum of factors: {sum(res.x)}")
            print(f"Extremum value: {extremum_value}")
            
            return {
                'x': res.x,  # Already in original scale for comprehensive
                'value': extremum_value
            }
            
        except Exception as e:
            messagebox.showerror("Optimization Error", f"Optimization failed: {str(e)}")
            return {'x': np.array([np.nan]*len(x0_for_opt)), 'value': np.nan}

    def verify_extremum_calculation(self):
        """Verify that the displayed extremum value matches manual calculation"""
        if (self.extremum_point is None or self.coefficients is None or 
            self.bits_array is None or 'x_normalized' not in self.extremum_point):
            return "Cannot verify - no extremum point or missing normalized coordinates"
        
        # Get the normalized extremum point
        x_normalized = self.extremum_point['x_normalized']
        
        # Calculate value using normalized inputs directly
        design_row = self.create_design_matrix(x_normalized.reshape(1, -1), self.bits_array)
        calculated_value = np.dot(design_row[0], self.coefficients)
        
        # Compare with displayed value
        displayed_value = self.extremum_point['value']
        
        verification_text = f"""
        Verification of Extremum Calculation:
        -----------------------------------
        Normalized extremum point: {x_normalized}
        Design row: {design_row[0]}
        Coefficients: {self.coefficients}
        Calculated value: {calculated_value:.6f}
        Displayed value: {displayed_value:.6f}
        Difference: {abs(calculated_value - displayed_value):.2e}
        """
        
        print(verification_text)
        return verification_text

    def _normalize_point(self, x_point_original_scale):
        if x_point_original_scale is None: return None
        x_original_np = np.array(x_point_original_scale, dtype=float)
        norm_type = self.norm_select.get()
        if norm_type == "No normalization" or self.norm_x_min is None or self.norm_x_max is None:
            return x_original_np
        if len(x_original_np) != len(self.norm_x_min): return None
        range_val = self.norm_x_max - self.norm_x_min
        range_val[range_val == 0] = 1
        if norm_type == "[-1, 1]":
            return 2 * (x_original_np - self.norm_x_min) / range_val - 1
        elif norm_type == "[0, 1]":
            return (x_original_np - self.norm_x_min) / range_val
        return x_original_np

    def _unnormalize_point(self, x_point_fitting_scale):
        if x_point_fitting_scale is None: return None
        x_fitting_np = np.array(x_point_fitting_scale, dtype=float)
        norm_type = self.norm_select.get()
        if norm_type == "No normalization" or self.norm_x_min is None or self.norm_x_max is None:
            return x_fitting_np
        if len(x_fitting_np) != len(self.norm_x_min): return None
        range_val = self.norm_x_max - self.norm_x_min
        if norm_type == "[-1, 1]":
            return (x_fitting_np + 1) / 2 * range_val + self.norm_x_min
        elif norm_type == "[0, 1]":
            return x_fitting_np * range_val + self.norm_x_min
        return x_fitting_np

    def generate_equation_and_definitions(self, beta, bits_array, n_factors):
        if beta is None or bits_array is None or bits_array.size == 0 or len(beta) != bits_array.shape[0]:
            return "y = (Model not fitted or error in terms)", "", ""

        # Initialize lists to store terms of each type
        constant_terms = []
        linear_terms = []
        quadratic_terms = []
        interaction_terms = []

        # Process each term and classify it
        for i, (coef, bits) in enumerate(zip(beta, bits_array)):
            if abs(coef) < 1e-7:
                continue  # Skip negligible terms

            term_type = self._classify_term(bits)
            term_str = self._format_term(coef, bits, n_factors)

            if term_type == 'constant':
                constant_terms.append(term_str)
            elif term_type == 'linear':
                # Store with factor index for sorting (f1, f2, f3,...)
                factor_idx = np.where(bits == 1)[0][0]
                linear_terms.append((factor_idx, term_str))
            elif term_type == 'quadratic':
                # Store with factor index for sorting (f1², f2², f3²,...)
                factor_idx = np.where(bits == 2)[0][0]
                quadratic_terms.append((factor_idx, term_str))
            elif term_type == 'interaction':
                # Store with factor indices for sorting (f1×f2, f1×f3,..., f2×f3, f2×f4,...)
                idxs = np.where(bits == 1)[0]
                interaction_terms.append((idxs[0], idxs[1], term_str))

        # Sort terms according to specified ordering
        # Linear terms: f1, f2, f3,...
        linear_terms.sort(key=lambda x: x[0])
        
        # Quadratic terms: f1², f2², f3²,...
        quadratic_terms.sort(key=lambda x: x[0])
        
        # Interaction terms: f1×f2, f1×f3,..., f2×f3, f2×f4,...
        interaction_terms.sort(key=lambda x: (x[0], x[1]))

        # Combine all terms in the specified order
        equation_terms = []
        
        # Add constant term if present
        if constant_terms:
            equation_terms.append(constant_terms[0])  # There should be only one constant term

        # Add linear terms (f1, f2, f3,...)
        equation_terms.extend([t[1] for t in linear_terms])
        
        # Add quadratic terms (f1², f2², f3²,...)
        equation_terms.extend([t[1] for t in quadratic_terms])
        
        # Add interaction terms (f1×f2, f1×f3,..., f2×f3, f2×f4,...)
        equation_terms.extend([t[2] for t in interaction_terms])

        # Construct the final equation string
        if not equation_terms:
            final_equation_str = "y = 0"
        else:
            # Join terms with appropriate signs
            equation_str = " ".join(equation_terms).replace(" + -", " - ")
            # Handle case where first term is negative
            if equation_str.startswith("- "):
                final_equation_str = f"y = {equation_str}"
            else:
                final_equation_str = f"y = {equation_str}"

        # Generate factor definitions
        factor_definitions = []
        if n_factors > 0 and hasattr(self, 'factor_cols') and self.factor_cols:
            for i in range(n_factors):
                short_name = f"{i+1}"
                original_name = self.col_name_mapping.get(
                    self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                    f"Factor {i+1}"
                )
                factor_definitions.append(f"{short_name} = {original_name}")
        elif n_factors > 0:
            for i in range(n_factors):
                factor_definitions.append(f"c{i+1} = Factor {i+1}")

        factor_defs_str = "\n".join(factor_definitions)
        function_defs_str = ""  # Not used in this implementation

        return final_equation_str, function_defs_str, factor_defs_str

    def _classify_term(self, bits):
        """Classify a term based on its bits pattern"""
        sum_powers = sum(bits)
        unique_powers = set(bits)
        
        if sum_powers == 0:
            return 'constant'
        elif sum_powers == 1 and unique_powers.issubset({0, 1}):
            return 'linear'
        elif sum_powers == 2 and 2 in unique_powers:
            return 'quadratic'
        elif sum_powers == 2 and unique_powers.issubset({0, 1}):
            return 'interaction'
        return 'other'

    def _format_term(self, coef, bits, n_factors):
        """Format an individual term with proper factor ordering"""
        term_parts = []
        
        # Handle constant term
        if sum(bits) == 0:
            return f"{coef:+.4f}"[1:] if coef >= 0 else f"{coef:+.4f}"
        
        # For non-constant terms, collect factor components
        for j in range(n_factors):
            if bits[j] == 1:
                term_parts.append(f"c{j+1}")
            elif bits[j] == 2:
                term_parts.append(f"c{j+1}²")
        
        # Special handling for interaction terms to ensure proper ordering
        if len(term_parts) == 2 and sum(bits) == 2 and 2 not in bits:
            # Sort interaction terms (f1×f2, not f2×f1)
            term_parts.sort(key=lambda x: int(x.replace('c', '').replace('²', '')))
        
        # Format the coefficient and term
        abs_coef = abs(coef)
        term_str = f"{abs_coef:.4f}"
        if term_parts:
            term_str += f" × {' × '.join(term_parts)}"
        
        # Add sign (handled when joining terms)
        return f"{'+' if coef >= 0 else '-'} {term_str}"

    def plot_results(self, n_factors, x_plot_idx=0, y_plot_idx=1):
        self.figure1.clf() 
        self.figure1.subplots_adjust(bottom=0.18, left=0.18, top=0.9, right=0.95) 
        self.figure1.patch.set_facecolor('#F0F0F0')
        ax1 = self.figure1.add_subplot(111, facecolor='#ffffff') 
        ax1.grid(True, linestyle=':', alpha=0.6, color='gray')
        
        for spine in ax1.spines.values(): 
            spine.set_edgecolor('gray')
        
        # Handle comprehensive optimization case
        if hasattr(self, 'result_functions'):
            # For comprehensive optimization, we'll show all individual fits
            colors = plt.cm.tab10(np.linspace(0, 1, len(self.result_functions)))
            
            for i, (result_col, func_data) in enumerate(self.result_functions.items()):
                y = func_data['y']
                y_pred = func_data['model'].predict(func_data['X_design'])
                
                ax1.scatter(y, y_pred, alpha=0.7, edgecolors='#333333', s=40, 
                            color=colors[i], label=self.col_name_mapping.get(result_col, result_col))
            
            min_val = min(min(func['y'].min(), func['model'].predict(func['X_design']).min()) 
                       for func in self.result_functions.values())
            max_val = max(max(func['y'].max(), func['model'].predict(func['X_design']).max()) 
                       for func in self.result_functions.values())
            
            ax1.plot([min_val, max_val], [min_val, max_val], 'r--', lw=1.5, label='Perfect Fit')
            ax1.legend(fontsize=9)
        elif self.y is not None and self.y_pred is not None:
            # Single result case
            ax1.scatter(self.y, self.y_pred, alpha=0.7, edgecolors='#333333', s=40, color="#007ACC")
            min_val, max_val = min(self.y.min(), self.y_pred.min()), max(self.y.max(), self.y_pred.max())
            ax1.plot([min_val, max_val], [min_val, max_val], 'r--', lw=1.5, label='Perfect Fit')
            ax1.legend(fontsize=9)
        else:
            ax1.text(0.5, 0.5, "No data to plot", ha="center", va="center", fontsize=10, color='gray')
        
        ax1.set_xlabel('Actual Values', fontsize=10, fontweight='bold')
        ax1.set_ylabel('Predicted Values', fontsize=10, fontweight='bold')
        ax1.set_title('Actual vs. Predicted Performance', fontsize=11, fontweight='bold')
        self.canvas1.draw()

        # Handle the 3D plot
        self.figure2.clf()
        self.figure2.subplots_adjust(bottom=0.18, left=0.1, right=0.85, top=0.9)
        self.figure2.patch.set_facecolor('#F0F0F0')
        ax2_facecolor = '#ffffff'
        
        if (not hasattr(self, 'comprehensive_function') and (self.coefficients is None or self.bits_array is None or 
            self.X_original_scale is None or n_factors == 0)):
            ax2 = self.figure2.add_subplot(111, facecolor=ax2_facecolor)
            ax2.text(0.5, 0.5, "Fit model to see surface plot", 
                    ha="center", va="center", fontsize=10, color='gray')
            for spine in ax2.spines.values(): 
                spine.set_edgecolor('gray')
            self.canvas2.draw()
            return

        if n_factors == 1:
            ax2 = self.figure2.add_subplot(111, facecolor=ax2_facecolor)
            ax2.grid(True, linestyle=':', alpha=0.6, color='gray')
            for spine in ax2.spines.values(): 
                spine.set_edgecolor('gray')
            
            x_orig_plot_axis = np.linspace(self.X_original_scale[:, 0].min(), 
                                          self.X_original_scale[:, 0].max(), 100)
            points_to_eval_csr_orig = x_orig_plot_axis.reshape(-1,1)
            
            # Handle comprehensive vs single result
            if hasattr(self, 'comprehensive_function'):
                z_csr_values = np.array([self.comprehensive_function(p) for p in points_to_eval_csr_orig])
            else:
                z_csr_values = np.array([self.evaluate_csr_at_point(self._normalize_point(p)) 
                                      if self._normalize_point(p) is not None else np.nan 
                                      for p in points_to_eval_csr_orig])
            
            ax2.plot(x_orig_plot_axis, z_csr_values, label='CSR Function', color="#D83B01", lw=2.5)
            
            # Plot original data points for all results in comprehensive case
            if hasattr(self, 'result_functions'):
                colors = plt.cm.tab10(np.linspace(0, 1, len(self.result_functions)))
                for i, result_col in enumerate(self.result_functions):
                    ax2.scatter(self.X_original_scale[:, 0], 
                               self.result_functions[result_col]['y'], 
                               color=colors[i], s=40, alpha=0.7, 
                               label=self.col_name_mapping.get(result_col, result_col),
                               edgecolors='#333333')
                ax2.legend(fontsize=9)
            else:
                ax2.scatter(self.X_original_scale[:, 0], self.y, color="#007ACC", 
                           s=40, alpha=0.7, label='Original Data', edgecolors='#333333')

            # Plot extremum point if available
            if self.extremum_point and self.extremum_point['x'] is not None and len(self.extremum_point['x']) > 0:
                if hasattr(self, 'comprehensive_function'):
                    extremum_x_orig = self.extremum_point['x']
                else:
                    extremum_x_orig = self._unnormalize_point(self.extremum_point['x'])
                
                if extremum_x_orig is not None and len(extremum_x_orig) > 0:
                    ax2.scatter([extremum_x_orig[0]], [self.extremum_point['value']], 
                               c='gold', s=200, marker='*', edgecolor='black', 
                               linewidth=1, label='Extremum', zorder=5)
                    ax2.legend(fontsize=9)

            factor_disp_name = self.col_name_mapping.get(self.factor_cols[0], self.factor_cols[0])
            result_disp_name = "Combined Result" if hasattr(self, 'comprehensive_function') else self.col_name_mapping.get("result", "Result")
            ax2.set_xlabel(factor_disp_name, fontsize=10, fontweight='bold')
            ax2.set_ylabel(result_disp_name, fontsize=10, fontweight='bold')
            ax2.set_title(f'CSR Output vs. {factor_disp_name}', fontsize=11, fontweight='bold')

        elif n_factors >= 2:
            ax2 = self.figure2.add_subplot(111, projection='3d', facecolor=ax2_facecolor)
            ax2.grid(True, linestyle=':', alpha=0.5)
            
            for pane_ax in [ax2.xaxis, ax2.yaxis, ax2.zaxis]: 
                pane_ax.set_pane_color((1.0, 1.0, 1.0, 0.0))
                pane_ax.pane.set_edgecolor('#D0D0D0')
            
            x1_orig_plot_axis = np.linspace(self.X_original_scale[:, x_plot_idx].min(), 
                                           self.X_original_scale[:, x_plot_idx].max(), 20)
            x2_orig_plot_axis = np.linspace(self.X_original_scale[:, y_plot_idx].min(), 
                                           self.X_original_scale[:, y_plot_idx].max(), 20)
            x1_grid_orig, x2_grid_orig = np.meshgrid(x1_orig_plot_axis, x2_orig_plot_axis)
            z_csr_values_grid = np.zeros_like(x1_grid_orig)
            
            # Get fixed values for other factors (use extremum point if available, otherwise mean)
            fixed_factor_values_original_scale = np.mean(self.X_original_scale, axis=0)
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                if hasattr(self, 'comprehensive_function'):
                    fixed_factor_values_original_scale = self.extremum_point['x']
                else:
                    unnorm_temp = self._unnormalize_point(self.extremum_point['x'])
                    if unnorm_temp is not None:
                        fixed_factor_values_original_scale = unnorm_temp
            
            # Evaluate comprehensive function across the grid
            for i_grid in range(x1_grid_orig.shape[0]):
                for j_grid in range(x1_grid_orig.shape[1]):
                    current_full_point_orig_scale = fixed_factor_values_original_scale.copy()
                    current_full_point_orig_scale[x_plot_idx] = x1_grid_orig[i_grid, j_grid]
                    current_full_point_orig_scale[y_plot_idx] = x2_grid_orig[i_grid, j_grid]
                    
                    if hasattr(self, 'comprehensive_function'):
                        z_csr_values_grid[i_grid, j_grid] = self.comprehensive_function(current_full_point_orig_scale)
                    else:
                        current_full_point_norm_scale = self._normalize_point(current_full_point_orig_scale)
                        if current_full_point_norm_scale is not None:
                            z_csr_values_grid[i_grid, j_grid] = self.evaluate_csr_at_point(current_full_point_norm_scale)
                        else:
                            z_csr_values_grid[i_grid, j_grid] = np.nan
            
            # Plot the surface
            surf = ax2.plot_surface(x1_grid_orig, x2_grid_orig, z_csr_values_grid, 
                                   cmap='viridis', alpha=0.9, edgecolor='#555555', 
                                   linewidth=0.15, antialiased=True)
            
            # Plot extremum point if available
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                if hasattr(self, 'comprehensive_function'):
                    extremum_x_orig_for_scatter = self.extremum_point['x']
                else:
                    extremum_x_orig_for_scatter = self._unnormalize_point(self.extremum_point['x'])
                
                if extremum_x_orig_for_scatter is not None:
                    ax2.scatter([extremum_x_orig_for_scatter[x_plot_idx]], 
                               [extremum_x_orig_for_scatter[y_plot_idx]], 
                               [self.extremum_point['value']], 
                               c='gold', s=200, marker='*', edgecolor='black', 
                               linewidth=1, label='Extremum', depthshade=True, zorder=10)
            
            x_axis_name = self.col_name_mapping.get(self.factor_cols[x_plot_idx], self.factor_cols[x_plot_idx])
            y_axis_name = self.col_name_mapping.get(self.factor_cols[y_plot_idx], self.factor_cols[y_plot_idx])
            result_axis_name = "Combined Result" if hasattr(self, 'comprehensive_function') else self.col_name_mapping.get("result", "Result")
            
            ax2.set_xlabel(f"\n{x_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_ylabel(f"\n{y_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_zlabel(f"\n{result_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_title('CSR Response Surface', fontsize=11, fontweight='bold', y=1.02)
            
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                ax2.legend(fontsize=8, facecolor='#F0F0F0', framealpha=0.8)
            
            cbar = self.figure2.colorbar(surf, ax=ax2, shrink=0.6, aspect=12, pad=0.15, format="%.2f")
            cbar.ax.tick_params(labelsize=8)
            cbar.outline.set_edgecolor('gray')
            ax2.view_init(elev=28, azim=130)
            
            for axis_obj in [ax2.xaxis, ax2.yaxis, ax2.zaxis]:
                axis_obj.set_tick_params(pad=3, labelsize=8)
                axis_obj.label.set_size(9)
        
        self.canvas2.draw()

    def evaluate_csr_at_point(self, x_point_in_fitting_scale):
        """Evaluate CSR function at a point (input should be in normalized scale)"""
        if hasattr(self, 'comprehensive_function'):
            # For comprehensive optimization, we need to unnormalize first
            x_orig_scale = self._unnormalize_point(x_point_in_fitting_scale)
            if x_orig_scale is None:
                return np.nan
            return self.comprehensive_function(x_orig_scale)
        elif self.coefficients is None or self.bits_array is None:
            return np.nan
        elif x_point_in_fitting_scale is None or len(x_point_in_fitting_scale) != self.bits_array.shape[1]:
            return np.nan
        
        # Direct calculation with normalized inputs
        x_point_reshaped = np.array(x_point_in_fitting_scale).reshape(1, -1)
        design_row = self.create_design_matrix(x_point_reshaped, self.bits_array)
        if design_row.shape[1] == 0:
            return np.nan
        return np.dot(design_row[0], self.coefficients)

    def get_evaluation_point_for_coeffs(self):
        if self.X_original_scale is None: messagebox.showinfo("Info", "Please load data first."); return None
        eval_choice = self.eval_point_combo.get()
        if eval_choice == "Factors at Minimum": return self.X_original_scale.min(axis=0)
        if eval_choice == "Factors at Maximum": return self.X_original_scale.max(axis=0)
        if eval_choice == "Factors at Extremum":
            if self.extremum_point and self.extremum_point['x'] is not None:
                if hasattr(self, 'X') and self.X is not None and len(self.extremum_point['x']) == self.X.shape[1]:
                    unnormalized_extremum = self._unnormalize_point(self.extremum_point['x'])
                    if unnormalized_extremum is not None: return unnormalized_extremum
                messagebox.showwarning("Warning", "Extremum point dimension mismatch or unnormalization failed. Using mean."); return np.mean(self.X_original_scale, axis=0)
            messagebox.showinfo("Info", "Extremum point not available. Using mean."); return np.mean(self.X_original_scale, axis=0)
        return np.mean(self.X_original_scale, axis=0)

    def calculate_term_contribution(self, coefficient_val, bits_row_def, x_eval_point_original_scale):
        if x_eval_point_original_scale is None or len(x_eval_point_original_scale) != len(bits_row_def): return 0
        x_eval_point_for_model_scale = self._normalize_point(x_eval_point_original_scale)
        if x_eval_point_for_model_scale is None : return 0
        term_product_val = 1.0
        for factor_idx, power_val in enumerate(bits_row_def):
            if factor_idx >= len(x_eval_point_for_model_scale): continue
            if power_val == 1: term_product_val *= x_eval_point_for_model_scale[factor_idx]
            elif power_val == 2: term_product_val *= x_eval_point_for_model_scale[factor_idx]**2
        return coefficient_val * term_product_val

    def plot_full_pie_and_get_term_details(self, ax, data_values_map, chart_title_suffix, canvas_to_draw, initial_start_angle=120.0):
        ax.cla()
        fig_parent = ax.get_figure()
        fig_parent.subplots_adjust(left=0.05, right=0.48, top=0.92, bottom=0.08)
        fig_parent.patch.set_facecolor('#F0F0F0')
        ax.set_facecolor('#FFFFFF')
        for spine in ax.spines.values(): spine.set_visible(False)
        ax.set_xticks([])
        ax.set_yticks([])

        if not data_values_map:
            ax.text(0.5, 0.5, f"No significant {chart_title_suffix.lower()} contributions",
                   ha='center', va='center', fontsize=8, wrap=True, color='gray')
            canvas_to_draw.draw()
            return {}, 0.0

        valid_slices = [{'label': l, 'size': abs(s), 'original_value': s}
                       for l, s in data_values_map.items() if abs(s) > 1e-9]

        if not valid_slices:
            ax.text(0.5, 0.5, f"Contributions for {chart_title_suffix.lower()} are negligible.",
                   ha='center', va='center', fontsize=8, wrap=True, color='gray')
            canvas_to_draw.draw()
            return {}, 0.0

        final_labels = [i['label'] for i in valid_slices]
        final_sizes_abs = np.array([i['size'] for i in valid_slices])
        final_original_values = [i['original_value'] for i in valid_slices]
        total_abs_sum_of_this_pie = np.sum(final_sizes_abs)

        if total_abs_sum_of_this_pie <= 1e-9:
            ax.text(0.5, 0.5, f"Total for {chart_title_suffix.lower()} is negligible.",
                   ha='center', va='center', fontsize=8, wrap=True, color='gray')
            canvas_to_draw.draw()
            return {}, 0.0

        # Assign colors based on our updated scheme
        slice_colors = [self.term_type_colors.get(label, self.term_type_colors['Default']) 
                       for label in final_labels]

        wedges_patches, _ = ax.pie(
            final_sizes_abs,
            colors=slice_colors,
            startangle=initial_start_angle,
            wedgeprops={'edgecolor': '#404040', 'linewidth': 0.7}
        )
        ax.axis('equal')

        legend_labels_full_pie = []
        for i, label in enumerate(final_labels):
            original_val = final_original_values[i]
            percentage_of_this_pie = (final_sizes_abs[i] / total_abs_sum_of_this_pie * 100) if total_abs_sum_of_this_pie > 0 else 0
            
            if abs(original_val) > 1e4 or (abs(original_val) < 1e-2 and abs(original_val) > 1e-9):
                formatted_value = f"{original_val:+.2e}"
            else:
                formatted_value = f"{original_val:+.3f}"
            legend_labels_full_pie.append(f"{label}: {formatted_value} ({percentage_of_this_pie:.1f}%)")

        ax.legend(
            wedges_patches, legend_labels_full_pie,
            loc="center left", bbox_to_anchor=(1.02, 0.5), fontsize=8,
            title_fontsize=9, labelspacing=0.8, borderpad=0.8,
            frameon=True, facecolor='#f8f8f8', edgecolor='gray'
        )

        term_details_for_return = {}
        current_angle_deg = initial_start_angle
        for i, label in enumerate(final_labels):
            proportion = final_sizes_abs[i] / total_abs_sum_of_this_pie
            angle_span_deg = proportion * 360.0
            end_angle_deg = current_angle_deg + angle_span_deg
            term_details_for_return[label] = {
                'start_angle': current_angle_deg % 360,
                'end_angle': end_angle_deg % 360,
                'original_value': final_original_values[i],
                'base_color': slice_colors[i]
            }
            current_angle_deg = end_angle_deg
        
        canvas_to_draw.draw()
        return term_details_for_return

    def plot_partial_pie_wedges(self, ax, factor_contributions_map,
                               sector_start_angle_deg, sector_end_angle_deg,
                               total_overall_effect_magnitude,
                               chart_title_suffix, canvas_to_draw, base_color_for_sector_hex):
        ax.cla()
        fig = ax.get_figure()
        fig.subplots_adjust(left=0.05, right=0.48, top=0.92, bottom=0.08)
        fig.patch.set_facecolor('#F0F0F0')
        ax.set_facecolor('#FFFFFF')
        for spine in ax.spines.values(): spine.set_visible(False)
        ax.set_xticks([])
        ax.set_yticks([])

        if not factor_contributions_map or total_overall_effect_magnitude < 1e-9:
            ax.text(0.5, 0.5, f"No data or negligible overall effect\nfor {chart_title_suffix.lower()}",
                   ha='center', va='center', fontsize=8, color='gray')
            canvas_to_draw.draw()
            return

        # Convert to list to maintain order
        sorted_items = list(factor_contributions_map.items())
        if not sorted_items:
            ax.text(0.5, 0.5, f"Negligible contributions\nfor {chart_title_suffix.lower()}",
                   ha='center', va='center', fontsize=8, color='gray')
            canvas_to_draw.draw()
            return

        # Prepare data for plotting
        labels = [item[0] for item in sorted_items]
        values = [item[1] for item in sorted_items]
        abs_values = np.abs(values)
        total_abs = np.sum(abs_values)
        
        # Generate color shades
        colors = self._generate_shades(base_color_for_sector_hex, len(labels))
        
        # Calculate angles
        sector_span = (sector_end_angle_deg - sector_start_angle_deg) % 360
        if sector_span <= 0: sector_span += 360
        
        # Create wedges
        wedges = []
        current_angle = sector_start_angle_deg
        for i, (label, value) in enumerate(sorted_items):
            proportion = abs(value) / total_abs if total_abs > 0 else 0
            angle_span = proportion * sector_span
            wedge = matplotlib.patches.Wedge(
                (0, 0), 1, current_angle, current_angle + angle_span,
                facecolor=colors[i], edgecolor='#404040', linewidth=0.7
            )
            ax.add_patch(wedge)
            wedges.append(wedge)
            current_angle += angle_span

        # Create legend
        legend_labels = []
        for label, value in sorted_items:
            percentage = (abs(value) / total_overall_effect_magnitude) * 100
            formatted_value = f"{value:+.2e}" if abs(value) > 1e4 or (0 < abs(value) < 1e-2) else f"{value:+.3f}"
            legend_labels.append(f"{label}: {formatted_value} ({percentage:.1f}%)")
        
        ax.legend(
            wedges, legend_labels,
            loc="center left", bbox_to_anchor=(1.02, 0.5), fontsize=8,
            title_fontsize=9, labelspacing=0.8, borderpad=0.8,
            frameon=True, facecolor='#f8f8f8', edgecolor='gray'
        )
        
        ax.set_xlim(-1.15, 1.15)
        ax.set_ylim(-1.15, 1.15)
        ax.axis('equal')
        canvas_to_draw.draw()

    def update_coefficient_pie_charts(self, initial_load=False):
        # Update factor definitions display
        factor_definitions = []
        if hasattr(self, 'factor_cols') and self.factor_cols:
            for i, factor in enumerate(self.factor_cols):
                short_name = f"c{i+1}"
                original_name = self.col_name_mapping.get(factor, f"Factor {i+1}")
                factor_definitions.append(f"{short_name} = {original_name}")
        
        definitions_text = "\n".join(factor_definitions) if factor_definitions else "No factor definitions available"
        
        if hasattr(self, 'factor_definitions_text'):
            self.factor_definitions_text.config(state=tk.NORMAL)
            self.factor_definitions_text.delete("1.0", tk.END)
            self.factor_definitions_text.insert(tk.END, definitions_text)
            self.factor_definitions_text.config(state=tk.DISABLED)
        
        if hasattr(self, 'definitions_text'):
            self.definitions_text.config(state=tk.NORMAL)
            self.definitions_text.delete("1.0", tk.END)
            self.definitions_text.insert(tk.END, definitions_text)
            self.definitions_text.config(state=tk.DISABLED)

        # Handle case when no model is fitted
        placeholder_text = 'Load data & Run Fitting' if initial_load else "No data or model fitted."
        if initial_load or self.coefficients is None or self.bits_array is None or self.df is None or not self.factor_cols:
            for i in range(4):
                if i < len(self.pie_axes):
                    self.pie_axes[i].cla()
                    self.pie_axes[i].text(0.5, 0.5, placeholder_text, ha='center', va='center', fontsize=9, color='gray')
                    for spine in self.pie_axes[i].spines.values(): spine.set_visible(False)
                    self.pie_axes[i].set_xticks([])
                    self.pie_axes[i].set_yticks([])
                    if i < len(self.pie_canvas): self.pie_canvas[i].draw()
            return

        # Get evaluation point for contributions
        x_eval_original_scale_for_contrib = self.get_evaluation_point_for_coeffs()
        if x_eval_original_scale_for_contrib is None:
            messagebox.showerror("Error", "Could not determine evaluation point for coefficient contributions.")
            return

        # Calculate term contributions
        term_contributions_map = {
            'constant': 0.0,
            'linear_total': 0.0,
            'quadratic_total': 0.0,
            'interaction_total': 0.0,
            'linear_factors': {},
            'quadratic_factors': {},
            'interaction_factors': {}
        }

        for coef_val, bits_def in zip(self.coefficients, self.bits_array):
            contribution_val = self.calculate_term_contribution(coef_val, bits_def, x_eval_original_scale_for_contrib)
            sum_of_powers = np.sum(bits_def)
            unique_powers = set(bits_def)

            if sum_of_powers == 0:
                term_contributions_map['constant'] += contribution_val
            elif sum_of_powers == 1 and unique_powers.issubset({0, 1}):
                idx = np.where(bits_def == 1)[0][0]
                name = f"c{idx+1}"
                term_contributions_map['linear_total'] += contribution_val
                term_contributions_map['linear_factors'][name] = term_contributions_map['linear_factors'].get(name, 0) + contribution_val
            elif sum_of_powers == 2 and 2 in unique_powers:
                idx = np.where(bits_def == 2)[0][0]
                name = f"c{idx+1}²"
                term_contributions_map['quadratic_total'] += contribution_val
                term_contributions_map['quadratic_factors'][name] = term_contributions_map['quadratic_factors'].get(name, 0) + contribution_val
            elif sum_of_powers == 2 and unique_powers.issubset({0, 1}):
                idxs = np.where(bits_def == 1)[0]
                if len(idxs) == 2:
                    idx1, idx2 = sorted(idxs)
                    name1 = f"c{idx1+1}"
                    name2 = f"c{idx2+1}"
                    d_name = f"{name1}×{name2}"
                    term_contributions_map['interaction_total'] += contribution_val
                    term_contributions_map['interaction_factors'][d_name] = term_contributions_map['interaction_factors'].get(d_name, 0) + contribution_val

        # Sort terms for consistent display
        def sort_linear_key(item):
            return int(item[0].replace('c', ''))
        
        def sort_quadratic_key(item):
            return int(item[0].replace('c', '').replace('²', ''))
        
        def sort_interaction_key(item):
            parts = item[0].split('×')
            return (int(parts[0].replace('c', '')), int(parts[1].replace('c', '')))

        term_contributions_map['linear_factors'] = dict(sorted(term_contributions_map['linear_factors'].items(), key=sort_linear_key))
        term_contributions_map['quadratic_factors'] = dict(sorted(term_contributions_map['quadratic_factors'].items(), key=sort_quadratic_key))
        term_contributions_map['interaction_factors'] = dict(sorted(term_contributions_map['interaction_factors'].items(), key=sort_interaction_key))

        # Calculate total effect magnitude
        overall_term_types_data = {
            'Constant': term_contributions_map['constant'],
            'Linear': term_contributions_map['linear_total'],
            'Quadratic': term_contributions_map['quadratic_total'],
            'Interaction': term_contributions_map['interaction_total']
        }
        total_overall_effect_magnitude = sum(abs(v) for v in overall_term_types_data.values())
        if total_overall_effect_magnitude < 1e-9:
            total_overall_effect_magnitude = 1.0

        # Plot Overall Pie Chart
        overall_pie_details = self.plot_full_pie_and_get_term_details(
            self.pie_axes[0],
            overall_term_types_data,
            "Overall Term Type Distribution",
            self.pie_canvas[0],
            initial_start_angle=120.0
        )

        # Plot Partial Pie Charts
        term_configs = [
            ("Linear", term_contributions_map['linear_factors'], self.pie_axes[1], self.pie_canvas[1]),
            ("Quadratic", term_contributions_map['quadratic_factors'], self.pie_axes[2], self.pie_canvas[2]),
            ("Interaction", term_contributions_map['interaction_factors'], self.pie_axes[3], self.pie_canvas[3]),
        ]

        for term_name, factor_data, ax, canvas in term_configs:
            if term_name in overall_pie_details and abs(overall_pie_details[term_name]['original_value']) > 1e-9:
                details = overall_pie_details[term_name]
                self.plot_partial_pie_wedges(
                    ax,
                    factor_data,
                    details['start_angle'],
                    details['end_angle'],
                    total_overall_effect_magnitude,
                    f"{term_name} Term Factor Contributions",
                    canvas,
                    details['base_color']
                )
            else:
                ax.cla()
                ax.text(0.5, 0.5, f"No significant {term_name.lower()}\ncontributions",
                       ha='center', va='center', fontsize=8, color='gray')
                canvas.draw()

    def _generate_shades(self, base_color_hex, n_shades):
        if n_shades == 0:
            return []
        
        base_rgb = mcolors.to_rgb(base_color_hex) # Ensure mcolors is imported
        base_h, base_s, base_v = mcolors.rgb_to_hsv(base_rgb)

        shades_rgb = []
        if n_shades == 1:
            # For a single shade, ensure it's reasonably visible
            # Adjust saturation and value to be in a 'good' range
            s_final = max(0.6, min(base_s, 0.95)) # Ensure decent saturation
            v_final = max(0.5, min(base_v if base_v > 0.1 else 0.7, 0.9)) # Avoid too dark/light
            shades_rgb.append(np.clip(mcolors.hsv_to_rgb((base_h, s_final, v_final)), 0, 1))
        else:
            # Preserve Hue. Vary Value (brightness). Adjust Saturation slightly if base is low.
            # Target saturation for shades to ensure color presence
            s_for_shades = max(base_s, 0.60) if base_s < 0.60 else base_s
            s_for_shades = min(s_for_shades, 0.95) # Cap saturation to avoid overly intense colors

            # Create a gradient of Value (brightness) from a darker to a lighter shade
            # Ensure a good range for Value to show distinction
            min_v_for_shade = 0.35
            max_v_for_shade = 0.95
            
            # If the original color is very dark, we might want to shift the range up.
            # If original color is very light, we might want to shift the range down.
            # For simplicity, we use a fixed perceptually good range for the shades' brightness.

            v_values = np.linspace(min_v_for_shade, max_v_for_shade, n_shades)
            
            for v_val in v_values:
                shades_rgb.append(np.clip(mcolors.hsv_to_rgb((base_h, s_for_shades, v_val)), 0, 1))
        return shades_rgb

    def _download_all_charts(self):
        if self.coefficients is None:
            messagebox.showwarning("No Data", "Please run the fitting process first to generate charts.")
            return

        if not self.pie_figures or not all(fig in self.pie_figures for fig in range(4)):
            messagebox.showwarning("No Charts", "No charts available to save.")
            return

        # Ask for directory to save files
        save_dir = filedialog.askdirectory(title="Select Directory to Save Charts")
        if not save_dir:
            return

        try:
            from matplotlib.backends.backend_pdf import PdfPages
            import tempfile
            import shutil

            # Create a temporary directory for intermediate files
            temp_dir = tempfile.mkdtemp()
            pdf_paths = []

            # Render all charts first
            for i in range(4):
                self.pie_canvas[i].draw()

            # Generate a single merged PDF
            merged_pdf_path = os.path.join(save_dir, "CSR_Charts_Merged.pdf")
            
            with PdfPages(merged_pdf_path) as merged_pdf:
                for i in range(4):
                    fig = self.pie_figures[i]
                    # Save each figure as a page in the merged PDF
                    merged_pdf.savefig(fig, bbox_inches='tight', facecolor=fig.get_facecolor())

            # Clean up temporary directory (not strictly needed here but good practice)
            shutil.rmtree(temp_dir, ignore_errors=True)

            messagebox.showinfo("Success", f"All charts saved as a single PDF:\n{merged_pdf_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save charts:\n{str(e)}")
            # Clean up temp dir even if error occurs
            shutil.rmtree(temp_dir, ignore_errors=True)
            
    def create_oacd_tab(self):
        # Main container for OACD tab
        oacd_frame = ttk.Frame(self.tab3, style="App.TFrame")
        oacd_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Split left (controls) and right (table display)
        left_frame = ttk.Frame(oacd_frame, style="App.TFrame")
        left_frame.pack(side='left', fill='y', padx=(0, 20), pady=5)
        right_frame = ttk.Frame(oacd_frame, style="App.TFrame")
        right_frame.pack(side='left', fill='both', expand=True, pady=5)

        # --- OACD Logic ---
        self.oacd = OACD()
        self.oacd.set_factor_num(2)
        self.oacd.set_table_size("Small")
        self.oacd.extrenum_vars = []  # List of (min_var, max_var) for each factor

        # --- Controls ---
        ttk.Label(left_frame, text="Number of Parameters:", font=self.label_font).pack(anchor='w', pady=(0,2))
        factor_num_combo = ttk.Combobox(left_frame, state="readonly", font=self.entry_font, width=8)
        factor_num_combo['values'] = list(range(2, 11))
        try:
            factor_num_combo.set(str(self.oacd.factor_num))
        except Exception:
            pass
        factor_num_combo.pack(anchor='w', pady=(0,8))
        factor_num_combo.bind('<<ComboboxSelected>>', lambda e: (self.oacd.set_factor_num(int(factor_num_combo.get())), self._oacd_update_extrenum_table()))

        ttk.Label(left_frame, text="Table Size:", font=self.label_font).pack(anchor='w', pady=(0,2))
        table_size_combo = ttk.Combobox(left_frame,state="readonly", font=self.entry_font, width=8)
        table_size_combo['values'] = ["Small", "Medium", "Large"]
        try:
            table_size_combo.set(self.oacd.table_size)
        except Exception:
            pass
        table_size_combo.pack(anchor='w', pady=(0,8))
        table_size_combo.bind('<<ComboboxSelected>>', lambda e: self.oacd.set_table_size(table_size_combo.get()))

        # --- Interactive Extrenum Table ---
        extrenum_frame = ttk.LabelFrame(left_frame, text="Factor Min/Max (Extrenum)", padding=(8,6,8,8))
        extrenum_frame.pack(fill='x', pady=(10,8))
        self.oacd_extrenum_table_frame = ttk.Frame(extrenum_frame, style="App.TFrame")
        self.oacd_extrenum_table_frame.pack(fill='x', expand=True)
        self._oacd_update_extrenum_table()

        # --- Set all min/max ---
        set_all_frame = ttk.Frame(left_frame, style="App.TFrame")
        set_all_frame.pack(fill='x', pady=(5,8))
        ttk.Label(set_all_frame, text="Set all min:").pack(side='left')
        self.oacd_set_all_min = tk.DoubleVar(value=0.0)
        min_entry = tk.Entry(set_all_frame, textvariable=self.oacd_set_all_min, width=7, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
        min_entry.pack(side='left', padx=(2,8))
        ttk.Button(set_all_frame, text="Apply", command=self._oacd_apply_all_min, width=6).pack(side='left')
        ttk.Label(set_all_frame, text="Set all max:").pack(side='left', padx=(10,0))
        self.oacd_set_all_max = tk.DoubleVar(value=0.0)
        max_entry = tk.Entry(set_all_frame, textvariable=self.oacd_set_all_max, width=7, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
        max_entry.pack(side='left', padx=(2,8))
        ttk.Button(set_all_frame, text="Apply", command=self._oacd_apply_all_max, width=6).pack(side='left')

        # --- Max Nonzero Controls ---
        max_nonzero_frame = ttk.Frame(left_frame, style="App.TFrame")
        max_nonzero_frame.pack(fill='x', pady=(5,8))
        ttk.Label(max_nonzero_frame, text="Max Nonzero Factors:").pack(side='left')
        self.oacd_max_nonzero_var = tk.IntVar(value=0)
        max_nonzero_entry = tk.Entry(max_nonzero_frame, textvariable=self.oacd_max_nonzero_var, width=7, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
        max_nonzero_entry.pack(side='left', padx=(2,8))
        ttk.Button(max_nonzero_frame, text="Apply", command=self._oacd_apply_max_nonzero, width=8).pack(side='left')

        # --- Limit Management Controls ---
        limit_frame = ttk.LabelFrame(left_frame, text="Parameter Limits", padding=(8,6,8,8))
        limit_frame.pack(fill='x', pady=(10,8))
        
        # Limit value input
        limit_input_frame = ttk.Frame(limit_frame, style="App.TFrame")
        limit_input_frame.pack(fill='x', pady=(0,5))
        ttk.Label(limit_input_frame, text="Limit Value:").pack(side='left')
        self.oacd_limit_value = tk.DoubleVar(value=100.0)
        limit_value_entry = tk.Entry(limit_input_frame, textvariable=self.oacd_limit_value, width=8, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
        limit_value_entry.pack(side='left', padx=(2,8))
        ttk.Button(limit_input_frame, text="Add Limit", command=self._oacd_add_limit, width=8).pack(side='left')
        
        # Current limits display
        limits_display_frame = ttk.Frame(limit_frame, style="App.TFrame")
        limits_display_frame.pack(fill='x', pady=(5,0))
        ttk.Label(limits_display_frame, text="Current Limits:").pack(anchor='w')
        self.oacd_limits_listbox = tk.Listbox(limits_display_frame, height=4, font=self.entry_font)
        self.oacd_limits_listbox.pack(fill='x', pady=(2,5))
        limits_button_frame = ttk.Frame(limits_display_frame, style="App.TFrame")
        limits_button_frame.pack(fill='x')
        ttk.Button(limits_button_frame, text="Remove Selected", command=self._oacd_remove_limit, width=12).pack(side='left')
        ttk.Button(limits_button_frame, text="Clear All", command=self._oacd_clear_limits, width=8).pack(side='left', padx=(5,0))

        # --- Generate and Export Buttons ---
        button_frame = ttk.Frame(left_frame, style="App.TFrame")
        button_frame.pack(fill='x', pady=(15,5))
        ttk.Button(button_frame, text="Import Extrenum from Excel", command=self._oacd_import_extrenum).pack(fill='x', pady=(0,8))
        ttk.Button(button_frame, text="Generate OACD Table", command=self._oacd_generate_table, style="Accent.TButton").pack(fill='x', pady=(0,8))
        ttk.Button(button_frame, text="Export as Excel", command=self._oacd_export_table).pack(fill='x', pady=(0,8))
        
        # --- OACD Table Display ---
        table_disp_frame = ttk.LabelFrame(right_frame, text="Generated OACD Table", padding=(8,6,8,8))
        table_disp_frame.pack(fill='both', expand=True)
        x_scroll = ttk.Scrollbar(table_disp_frame, orient='horizontal')
        x_scroll.pack(side='bottom', fill='x')
        y_scroll = ttk.Scrollbar(table_disp_frame, orient='vertical')
        y_scroll.pack(side='right', fill='y')
        self.oacd_table_tree = ttk.Treeview(table_disp_frame, show='headings', xscrollcommand=x_scroll.set, yscrollcommand=y_scroll.set, height=15)
        self.oacd_table_tree.pack(fill='both', expand=True)
        x_scroll.config(command=self.oacd_table_tree.xview)
        y_scroll.config(command=self.oacd_table_tree.yview)

    def _oacd_update_extrenum_table(self):
        # Clear previous widgets
        for widget in self.oacd_extrenum_table_frame.winfo_children():
            widget.destroy()
        self.oacd.extrenum_vars = []
        n = self.oacd.factor_num if getattr(self.oacd, 'factor_num', None) is not None else 2
        # Header
        ttk.Label(self.oacd_extrenum_table_frame, text="Factor", width=8).grid(row=0, column=0, padx=2, pady=2)
        ttk.Label(self.oacd_extrenum_table_frame, text="Min", width=8).grid(row=0, column=1, padx=2, pady=2)
        ttk.Label(self.oacd_extrenum_table_frame, text="Max", width=8).grid(row=0, column=2, padx=2, pady=2)
        ttk.Label(self.oacd_extrenum_table_frame, text="Limit", width=8).grid(row=0, column=3, padx=2, pady=2)
        for i in range(n):
            ttk.Label(self.oacd_extrenum_table_frame, text=f"c{i+1}", width=8).grid(row=i+1, column=0, padx=2, pady=2)
            min_var = tk.DoubleVar(value=0.0)
            max_var = tk.DoubleVar(value=0.0)
            min_entry = tk.Entry(self.oacd_extrenum_table_frame, textvariable=min_var, width=8, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
            max_entry = tk.Entry(self.oacd_extrenum_table_frame, textvariable=max_var, width=8, font=self.entry_font, insertwidth=1, insertontime=500, insertofftime=500)
            
            min_entry.grid(row=i+1, column=1, padx=2, pady=2)
            max_entry.grid(row=i+1, column=2, padx=2, pady=2)
            self.oacd.extrenum_vars.append((min_var, max_var))
            
        self._oacd_update_limit_factor_selection()

    def _oacd_apply_all_min(self):
        for min_var, _ in self.oacd.extrenum_vars:
            min_var.set(self.oacd_set_all_min.get())

    def _oacd_apply_all_max(self):
        for _, max_var in self.oacd.extrenum_vars:
            max_var.set(self.oacd_set_all_max.get())

    def _oacd_apply_max_nonzero(self):
        value = self.oacd_max_nonzero_var.get()
        if value <= 0:
            tk.messagebox.showwarning("Invalid Value", "Please enter a positive integer for max nonzero factors.")
            return
        self.oacd.max_nonzero = value
        self.oacd.reduce_levels()
        self._oacd_display_table()

    def _oacd_generate_table(self):
        # Set up OACD object
        print(self.oacd.limits)
        n = self.oacd.factor_num
        # Build extrenum DataFrame from UI
        extrenum = np.zeros((n,2))
        for i, (min_var, max_var) in enumerate(self.oacd.extrenum_vars):
            extrenum[i,0] = min_var.get()
            extrenum[i,1] = max_var.get()
        self.oacd.set_factor_extrenum(pd.DataFrame(extrenum))
        result = self.oacd.build_table()
        if result != 1:
            messagebox.showerror("Error", "Failed to build OACD table. Check factor number and table size.")
            return
        self.oacd.normalize_table()
        self._oacd_display_table()

    def _oacd_display_table(self):
        # Display the table
        for col in self.oacd_table_tree.get_children():
            self.oacd_table_tree.delete(col)
        columns = ["Run"] + [f"c{i+1}" for i in range(self.oacd.table.shape[1])]
        self.oacd_table_tree['columns'] = columns
        for i, col in enumerate(columns):
            self.oacd_table_tree.heading(col, text=col)
            self.oacd_table_tree.column(col, width=80, anchor='center')
        for idx, row in self.oacd.table.iterrows():
            values = [str(idx+1)] + [f"{x:.4f}" for x in row.values]
            self.oacd_table_tree.insert('', 'end', values=values)

    def _oacd_export_table(self):
        if self.oacd.table is None:
            messagebox.showwarning("No Table", "Please generate the OACD table first.")
            return
        file_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[("Excel files", "*.xlsx")], title="Save OACD Table As")
        if not file_path:
            return
        try:
            self.oacd.table.to_excel(file_path, index=False)
            messagebox.showinfo("Exported", f"OACD table exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export table:\n{str(e)}")
            
    def _oacd_import_extrenum(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")], title="Select Extrenum Excel File")
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path, header=None)
            n = self.oacd.factor_num
            if df.shape != (n, 2):
                messagebox.showerror("Import Error", f"Extrenum table must have {n} rows and 2 columns (min, max).\nImported shape: {df.shape}")
                return
            # Update UI variables
            for i in range(n):
                self.oacd.extrenum_vars[i][0].set(df.iloc[i,0])
                self.oacd.extrenum_vars[i][1].set(df.iloc[i,1])
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import extrenum table:\n{str(e)}")
            
    def _oacd_update_limit_factor_selection(self):
        """Update the factor selection checkboxes for limits"""
        # Check if the limit factor frame exists (UI might not be fully initialized yet)
        if not hasattr(self, 'oacd_extrenum_table_frame') or self.oacd_extrenum_table_frame is None:
            return
        
        n = self.oacd.factor_num if getattr(self.oacd, 'factor_num', None) is not None else 2
        
        # Clear existing comboboxes if they exist
        if hasattr(self, 'factor_combo'):
            for combo in self.factor_combo:
                if combo.winfo_exists():
                    combo.destroy()
        
        self.factor_combo = []
        
        # Create comboboxes for each factor
        for i in range(n):
            # Create combobox for factor limit selection
            combo = ttk.Combobox(self.oacd_extrenum_table_frame, state="readonly", font=self.entry_font, width=20)
            combo['values'] = list(self.oacd.limit_names.values())
            combo.grid(row=i+1, column=3, padx=(2,5))
            combo.current(0)
            
            # Bind the selection event with the correct factor index
            combo.bind('<<ComboboxSelected>>', lambda e, factor_idx=i: self._oacd_combo_select(factor_idx))
            
            # Add to the list of comboboxes
            self.factor_combo.append(combo)
    
    def _oacd_combo_select(self, i):
        # Get the combobox for the specific factor
        if i < len(self.factor_combo):
            combo = self.factor_combo[i]
            
            limit_key = next((key for key, value in self.oacd.limit_names.items() if value == combo.get()), None)
            if limit_key is None:
                name_entry = None
                index_entry = None
            else:
                name_entry = float(limit_key.split("_")[0])
                index_entry = int(limit_key.split("_")[1])
            self.oacd.add_limit([i], name_entry, index=index_entry)
    
    def _oacd_add_limit(self):
        """Add a new limit with selected factors"""
            
        limit_value = self.oacd_limit_value.get()
        if limit_value <= 0:
            messagebox.showwarning("Invalid Value", "Limit value must be positive.")
            return
            
        # Add limit to OACD object
        self.oacd.add_limit([], limit_value)
        
        # Update limits display
        self._oacd_update_limit_factor_selection()
        self._oacd_update_limits_display()
    
    def _oacd_remove_limit(self):
        """Remove the selected limit"""
        # Check if limits listbox exists
        if not hasattr(self, 'oacd_limits_listbox'):
            messagebox.showwarning("UI Not Ready", "Please wait for the UI to fully load.")
            return
            
        selection = self.oacd_limits_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a limit to remove.")
            return
            
        # Get the limit key from the listbox
        limit_name = list(self.oacd.limit_names.keys())[selection[0] + 1]
        
        self.oacd.remove_limit(limit_name)
            
        # Update limits display
        self._oacd_update_limits_display()
        self._oacd_update_limit_factor_selection()
    
    def _oacd_clear_limits(self):
        """Clear all limits"""
        if messagebox.askyesno("Clear All Limits", "Are you sure you want to clear all limits?"):
            # Clear all limit columns
            self.oacd.remove_all_limits()
            self._oacd_update_limits_display()
            self._oacd_update_limit_factor_selection()
    
    def _oacd_update_limits_display(self):
        """Update the limits listbox display"""
        # Check if limits listbox exists
        if not hasattr(self, 'oacd_limits_listbox'):
            return
            
        self.oacd_limits_listbox.delete(0, tk.END)
        
        limits_names = list(self.oacd.limit_names.values())
        for name in limits_names:
            self.oacd_limits_listbox.insert(tk.END, name)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSRApp(root)
    root.mainloop()
