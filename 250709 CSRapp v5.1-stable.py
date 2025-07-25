import os
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

        placeholder_label_tab3 = ttk.Label(self.tab3, text="OACD Table Builder - To be implemented", font=self.title_font)
        placeholder_label_tab3.pack(padx=10, pady=20, anchor='center')

    def create_csr_integration_tab(self):
        self.main_paned = tk.PanedWindow(self.tab1, orient=tk.HORIZONTAL, sashrelief=tk.GROOVE, sashwidth=8, background="#D0D0D0", bd=0)
        self.main_paned.pack(fill=tk.BOTH, expand=True)

        left_frame = ttk.Frame(self.main_paned, padding=15, style="App.TFrame")
        self.main_paned.add(left_frame, width=430, minsize=400, sticky="nsew")

        center_frame = ttk.Frame(self.main_paned, padding=(5,15,15,15), style="App.TFrame")
        self.main_paned.add(center_frame, width=650, minsize=500, sticky="nsew")

        right_paned_container = ttk.Frame(self.main_paned, style="App.TFrame", padding=0)
        self.main_paned.add(right_paned_container, width=450, minsize=420, sticky="nsew")

        right_paned = tk.PanedWindow(right_paned_container, orient=tk.VERTICAL, sashrelief=tk.GROOVE, sashwidth=8, background="#D0D0D0", bd=0)
        right_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=0)

        # === Left Panel Contents ===
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


        ttk.Label(left_frame, text="Normalization Standard:").pack(anchor='w', padx=5)
        self.norm_select = ttk.Combobox(left_frame, values=["[-1, 1]", "[0, 1]"], state="readonly", font=self.entry_font)
        self.norm_select.pack(fill='x', padx=5, pady=(2,10))
        self.norm_select.current(1)  # Default to [0, 1]


        reg_frame = ttk.LabelFrame(left_frame, text="Regularization Parameter", padding=(10,5,10,10))
        reg_frame.pack(fill='x', pady=10, padx=5)

        # Create the label first
        self.alpha_value_label = ttk.Label(reg_frame, text="1.00", width=7, anchor='e')
        self.alpha_value_label.grid(row=0, column=2, sticky='e', padx=(5,0), pady=3)

        # Then create the slider that references it
        self.alpha_slider = ttk.Scale(reg_frame, from_=-4, to=0, orient=tk.HORIZONTAL, command=self.update_alpha_value)
        self.alpha_slider.set(0)
        self.alpha_slider.grid(row=0, column=1, sticky='ew', padx=(10,0), pady=3)

        ttk.Label(reg_frame, text="Alpha (penalty):").grid(row=0, column=0, sticky='w', pady=3)
        reg_frame.columnconfigure(1, weight=1)

        ttk.Button(left_frame, text="Run Fitting Process", command=self.run_fitting).pack(pady=15, padx=5, fill='x', ipady=5)

        # Updated: Add frames for Function and Factor Definitions
        # Function Definition Frame
        function_def_frame = ttk.LabelFrame(left_frame, text="CSR Function", padding=(10, 5, 10, 10))
        function_def_frame.pack(fill='x', pady=(0, 5), padx=5, expand=True)

        eqn_text_frame = ttk.Frame(function_def_frame, style="App.TFrame")
        eqn_text_frame.pack(fill='both', expand=True, padx=0, pady=(0, 5)) # Adjusted pady

        self.equation_text = tk.Text(eqn_text_frame, height=4, wrap=tk.WORD, state='disabled',
                                     font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=8, pady=5,spacing1=2, spacing2=2, spacing3=2)
        scrollbar_eqn = ttk.Scrollbar(eqn_text_frame, command=self.equation_text.yview, orient=tk.VERTICAL)
        self.equation_text['yscrollcommand'] = scrollbar_eqn.set
        scrollbar_eqn.pack(side=tk.RIGHT, fill=tk.Y)
        self.equation_text.pack(side=tk.LEFT, fill='both', expand=True)

        # Factor Definition Frame
        factor_def_frame = ttk.LabelFrame(left_frame, text="Factor Definitions", padding=(10, 5, 10, 10))
        factor_def_frame.pack(fill='x', pady=(0, 5), padx=5, expand=True)

        factor_def_text_frame = ttk.Frame(factor_def_frame, style="App.TFrame")
        factor_def_text_frame.pack(fill='both', expand=True, padx=0, pady=(0, 5))

        self.factor_definitions_text = tk.Text(factor_def_text_frame, height=3, wrap=tk.WORD, state='disabled',
                                               font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=8, pady=5,spacing1=2, spacing2=2, spacing3=2)
        scrollbar_factor_def = ttk.Scrollbar(factor_def_text_frame, command=self.factor_definitions_text.yview, orient=tk.VERTICAL)
        self.factor_definitions_text['yscrollcommand'] = scrollbar_factor_def.set
        scrollbar_factor_def.pack(side=tk.RIGHT, fill=tk.Y)
        self.factor_definitions_text.pack(side=tk.LEFT, fill='both', expand=True)

        # === Center Panel Contents ===
        center_container = ttk.Frame(center_frame, style="App.TFrame")
        center_container.pack(fill='both', expand=True, pady=(0,5))
        
        # Create a scrollable frame for the factor selection
        scroll_frame = ttk.Frame(center_container, style="App.TFrame")
        scroll_frame.pack(fill='both', expand=True)
        
        # Create a canvas and scrollbar
        canvas = tk.Canvas(scroll_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack the canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Factor selection frame
        factor_frame = ttk.LabelFrame(scrollable_frame, text="Select Factors to Include", padding=10)
        factor_frame.pack(fill='x', pady=(0,10))
        
        self.factor_checkbox_frame = ttk.Frame(factor_frame, style="App.TFrame")
        self.factor_checkbox_frame.pack(fill='both', expand=True)
        
        # Initialize checkbox dictionary
        self.factor_checkboxes = {}
        
        # Add a label for the checkboxes
        ttk.Label(self.factor_checkbox_frame, text="Select factors to include in analysis:", 
                font=self.label_font).grid(row=0, column=0, sticky='w', pady=(0,10))
        
        # Key Results Frame (keep this the same)
        results_frame = ttk.LabelFrame(scrollable_frame, text="Key Results", padding=(10,5,10,10))
        results_frame.pack(fill='x', pady=(10,0), expand=True)
        
        # Create a canvas and scrollbar for the results frame
        results_canvas = tk.Canvas(results_frame, highlightthickness=0)
        scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=results_canvas.yview)
        scrollable_results_frame = ttk.Frame(results_canvas)
        
        scrollable_results_frame.bind(
            "<Configure>",
            lambda e: results_canvas.configure(
                scrollregion=results_canvas.bbox("all")
            )
        )
        
        results_canvas.create_window((0, 0), window=scrollable_results_frame, anchor="nw")
        results_canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack the canvas and scrollbar
        results_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Extremum section
        extremum_frame = ttk.LabelFrame(scrollable_results_frame, text="Extremum (Original Scale)", padding=(10,5,10,10))
        extremum_frame.pack(fill='x', pady=(0,10))
        
        # Individual results frame for comprehensive optimization
        self.individual_results_frame = ttk.LabelFrame(scrollable_results_frame, text="Individual Results at Extremum", padding=(10,5,10,10))
        self.individual_results_frame.pack(fill='x', pady=(0,10))
        
        # R² frame
        self.r2_frame = ttk.LabelFrame(scrollable_results_frame, text="Model Fit (R²)", padding=(10,5,10,10))
        self.r2_frame.pack(fill='x', pady=(0,10))
        
        # Initialize all text widgets (keep this the same)
        self.factors_text = tk.Text(extremum_frame, height=1, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
        self.factors_text.grid(row=0, column=1, sticky='ew', pady=3)
        
        self.value_text = tk.Text(extremum_frame, height=1, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
        self.value_text.grid(row=1, column=1, sticky='ew', pady=3)
        
        # Initialize R² text widget
        self.r2_text = tk.Text(self.r2_frame, height=1, wrap=tk.NONE, state='disabled',
                            font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
        self.r2_text.pack(fill='x', padx=5, pady=5)
        
        # Configure grid weights
        extremum_frame.columnconfigure(1, weight=1)
        self.r2_frame.columnconfigure(0, weight=1)


        # === Right Panel Contents (PanedWindow) ===
        # plot1_frame and plot2_frame added to right_paned
        plot1_frame = ttk.LabelFrame(right_paned, text="Actual vs. Predicted Values", padding=10)
        right_paned.add(plot1_frame)
        right_paned.paneconfig(plot1_frame) # Configure weight after adding

        self.figure1 = Figure(figsize=(5, 4), dpi=100, facecolor='#F0F0F0')
        self.figure1.subplots_adjust(bottom=0.18, left=0.18, top=0.9, right=0.95)
        self.canvas1 = FigureCanvasTkAgg(self.figure1, master=plot1_frame)
        self.canvas1.get_tk_widget().pack(fill='both', expand=True, padx=5, pady=5)
        toolbar1 = NavigationToolbar2Tk(self.canvas1, plot1_frame)
        toolbar1.update()
        toolbar1.configure(background='#F0F0F0') # Match background
        for child_widget in toolbar1.winfo_children(): child_widget.configure(background='#F0F0F0')


        plot2_frame = ttk.LabelFrame(right_paned, text="CSR Response Surface Plot", padding=10)
        right_paned.add(plot2_frame) # Removed weight here
        right_paned.paneconfig(plot2_frame) # Configure weight after adding

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

            ttk.Label(controls_frame, text="Evaluate Term Contributions at Factor Levels:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
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
            self.definitions_frame = ttk.LabelFrame(top_frame, text="Factor Definitions", padding=10)
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
            self.definitions_text.insert(tk.END, "Factor definitions will appear here after fitting")
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

    def select_file(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
            if not file_path:
                return

            # Clear previous state completely
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
        for widget in self.factor_checkbox_frame.winfo_children():
            widget.destroy()
        
        self.factor_checkboxes = {}
        
        if self.df is None or self.df.empty:
            return
        
        # Get all column names in order
        all_columns = list(self.df.columns)
        
        # Add checkboxes for all features
        for i, col in enumerate(all_columns):
            display_name = self.col_name_mapping.get(col, col)
            
            # Create frame for this factor's controls
            cb_frame = ttk.Frame(self.factor_checkbox_frame, style="App.TFrame")
            cb_frame.pack(fill='x', pady=2)
            
            # Add checkbox - only check the last box by default
            var = tk.BooleanVar(value=(i == len(all_columns)-1))  # True for last column
            cb = ttk.Checkbutton(cb_frame, variable=var, text=display_name)
            cb.pack(side='left', padx=(0,10))
            self.factor_checkboxes[col] = var

    def clear_results_and_plots(self):
        # Clear text widgets safely
        text_widgets = [
            'equation_text', 
            'factor_definitions_text',
            'factors_text', 
            'value_text', 
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



    def update_max_iter_value(self, value_str):
        try:
            value = float(value_str)
            max_iter = int(10 ** value)
            self.max_iter_value_label.config(text=f"{max_iter}")
        except ValueError:
            pass

    def update_alpha_value(self, value):
        """Update the alpha value label when the slider is moved"""
        try:
            alpha = 10 ** float(value)
            if alpha < 0.01:
                # Use scientific notation for small values
                alpha_str = "{:.2e}".format(alpha)
                # Format to remove leading zero in exponent (e.g., 1.00e-02 → 1.00e-2)
                alpha_str = alpha_str.replace("e-0", "e-").replace("e+0", "e+")
            else:
                # Regular decimal format for larger values
                alpha_str = f"{alpha:.2f}"
            self.alpha_value_label.config(text=alpha_str)
        except ValueError:
            self.alpha_value_label.config(text="1.00")

    def run_fitting(self):
        try:
            if self.df is None:
                raise ValueError("Please load data first.")
                
            # Store current checkbox states before updating
            current_states = {}
            if hasattr(self, 'factor_checkboxes'):
                current_states = {col: var.get() for col, var in self.factor_checkboxes.items()}
                
            # Get selected columns - use current_states if available, otherwise default to last column
            if current_states:
                self.result_cols = [col for col, state in current_states.items() 
                                  if state and col != "residual"]
            else:
                # Default to just the result column if no checkboxes exist yet
                self.result_cols = ["result"]
            
            if not self.result_cols:
                raise ValueError("Please select at least one result factor.")

            if len(self.result_cols) == 1:
                self._run_single_result_fitting(self.result_cols[0])
            else:
                self._run_comprehensive_fitting()
                
            # Update table view while preserving checkbox states
            self.update_table_view()
            
            # Restore checkbox states
            if hasattr(self, 'factor_checkboxes'):
                for col, var in self.factor_checkboxes.items():
                    if col in current_states:
                        var.set(current_states[col])
            
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
                           if col.startswith('factor') or col == result_col]
            
            # Create a working dataframe with just the needed columns
            working_df = self.df[working_cols].copy()
            
            # Rename the result column to 'result' for consistency
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

            self.X = X_fit

            X_design = self.create_design_matrix(self.X, self.bits_array)
            if X_design.shape[1] == 0:
                messagebox.showerror("Error", "Design matrix has no terms.")
                return

            alpha_val = 10 ** float(self.alpha_slider.get())

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

            if norm_type == "[-1, 1]":
                bounds_opt = [(-1, 1)] * n_factors
                x0_opt = np.zeros(n_factors)
            elif norm_type == "[0, 1]":
                bounds_opt = [(0, 1)] * n_factors
                x0_opt = np.full(n_factors, 0.5)
            else:
                bounds_opt = [(self.X_original_scale[:,i].min(), self.X_original_scale[:,i].max()) for i in range(n_factors)]
                x0_opt = np.mean(self.X_original_scale, axis=0)

            extremum_type_str = self.weight_combo.get().lower().replace(" ", "_")
            extremum_result_in_fitting_space = self.find_extremum(self.coefficients, self.bits_array,
                                               bounds_opt, x0_opt, extremum_type_str, self.X)

            self.extremum_point = extremum_result_in_fitting_space

            # Generate equation and definitions
            equation_str, function_defs_str, factor_defs_str = self.generate_equation_and_definitions(
                self.coefficients, self.bits_array, n_factors)

            # Clear individual results frame for single optimization
            if hasattr(self, 'individual_results_frame'):
                for widget in self.individual_results_frame.winfo_children():
                    widget.destroy()
            
            # Clear R² frame for single optimization
            if hasattr(self, 'r2_frame'):
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
                
                # Create a simple frame for R² display
                r2_frame = ttk.Frame(self.r2_frame)
                r2_frame.pack(fill='x', pady=(0,5))
                
                ttk.Label(r2_frame, text="R²:", font=self.label_font).grid(row=0, column=0, sticky='w')
                
                r2_text = tk.Text(r2_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                r2_text.grid(row=0, column=1, padx=(5,0), sticky='w')
                r2_text.config(state='normal')
                r2_text.insert(tk.END, f"{train_r2:.4f}")
                r2_text.config(state='disabled')
                
                # Add interpretation
                interpretation = ""
                if train_r2 >= 0.9:
                    interpretation = "Excellent fit"
                elif train_r2 >= 0.7:
                    interpretation = "Good fit"
                elif train_r2 >= 0.5:
                    interpretation = "Moderate fit"
                else:
                    interpretation = "Poor fit"
                
                interp_text = tk.Text(r2_frame, height=1, width=20, wrap=tk.NONE, state='disabled',
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                interp_text.grid(row=0, column=2, padx=(5,0), sticky='w')
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

        except Exception as e:
            raise Exception(f"Single result fitting failed: {str(e)}")

    def _run_comprehensive_fitting(self):
        """Handle fitting for multiple result columns (comprehensive optimization)"""
        try:
            # First fit individual CSR functions for each result
            self.result_functions = {}
            result_min_max = {}
            
            # Initialize residuals column
            self.df['residual'] = 0.0
            
            for result_col in self.result_cols:
                # Create temporary df with this result column
                temp_df = self.df[self.factor_cols + [result_col]].copy()
                temp_df.columns = [f"factor{i+1}" for i in range(len(self.factor_cols))] + ["result"]
                
                # Store min/max for normalization
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
                alpha_val = 10 ** float(self.alpha_slider.get())
                model = Ridge(alpha=alpha_val, fit_intercept=False)
                model.fit(X_design, y_fit)
                
                # Calculate predictions and residuals
                y_pred = model.predict(X_design)
                residuals = y_fit - y_pred
                
                # Store the function and update residuals
                self.result_functions[result_col] = {
                    'coefficients': model.coef_,
                    'bits_array': bits_array,
                    'x_min': x_min if norm_type != "None" else None,
                    'x_max': x_max if norm_type != "None" else None,
                    'norm_type': norm_type,
                    'model': model,
                    'X_design': X_design,
                    'y': y_fit,
                    'y_pred': y_pred,
                    'residuals': residuals,
                    'min_val': result_min_max[result_col]['min'],
                    'max_val': result_min_max[result_col]['max']
                }
                
                # Add residuals to main dataframe (average if multiple results)
                self.df['residual'] += residuals / len(self.result_cols)
            
            # Create comprehensive function
            def comprehensive_func(x):
                total = 0
                for result_col, func_data in self.result_functions.items():
                    # Evaluate individual function
                    if func_data['norm_type'] == "[-1, 1]":
                        x_norm = 2 * (x - func_data['x_min']) / (func_data['x_max'] - func_data['x_min']) - 1
                    elif func_data['norm_type'] == "[0, 1]":
                        x_norm = (x - func_data['x_min']) / (func_data['x_max'] - func_data['x_min'])
                    else:
                        x_norm = x
                        
                    design_row = self.create_design_matrix(x_norm.reshape(1, -1), func_data['bits_array'])
                    val = np.dot(design_row[0], func_data['coefficients'])
                    
                    # Normalize and add to total
                    norm_val = (val - func_data['min_val']) / (func_data['max_val'] - func_data['min_val'])
                    total += norm_val
                return total
            
            self.comprehensive_function = comprehensive_func
            
            # Find extremum of comprehensive function - use original scale bounds
            n_factors = len(self.factor_cols)
            norm_type = self.norm_select.get()
            
            # Set bounds in original scale regardless of normalization
            bounds_opt = [(self.df[col].min(), self.df[col].max()) for col in self.factor_cols]
            x0_opt = np.array([(min_val + max_val)/2 for min_val, max_val in bounds_opt])
            
            extremum_type_str = self.weight_combo.get().lower().replace(" ", "_")
            self.extremum_point = self.find_extremum_comprehensive(bounds_opt, x0_opt, extremum_type_str)
            
            # Update displays
            self.update_comprehensive_results_display(result_min_max)
            
            # Set up factor combos for plotting
            display_factor_names = [self.col_name_mapping.get(f, f) for f in self.factor_cols]
            self.x_factor_combo['values'] = display_factor_names
            self.y_factor_combo['values'] = display_factor_names
            if display_factor_names:
                self.x_factor_combo.current(0)
                self.y_factor_combo.current(min(1, len(display_factor_names)-1))
            
            # Update the table view with the new data
            self.update_table_view()
            
            self.plot_results(n_factors)

        except Exception as e:
            raise Exception(f"Comprehensive fitting failed: {str(e)}")

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
                print(f"Error updating factor definitions: {e}")

        # Update Extremum Factors Display
        if hasattr(self, 'factors_text'):
            try:
                self.factors_text.config(state='normal')
                self.factors_text.delete(1.0, tk.END)
                
                if self.extremum_point and 'x' in self.extremum_point:
                    if hasattr(self, 'comprehensive_function'):
                        # For comprehensive optimization, extremum point is already in original scale
                        extremum_x = self.extremum_point['x']
                    else:
                        # For single optimization, we need to unnormalize
                        extremum_x = self._unnormalize_point(self.extremum_point['x'])
                    
                    if extremum_x is not None:
                        self.factors_text.insert(tk.END, ", ".join([f"{val:.4f}" for val in extremum_x]))
                
                self.factors_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating factors text: {e}")

        # Update Extremum Value Display
        if hasattr(self, 'value_text'):
            try:
                self.value_text.config(state='normal')
                self.value_text.delete(1.0, tk.END)
                
                if self.extremum_point and 'value' in self.extremum_point:
                    self.value_text.insert(tk.END, f"{self.extremum_point['value']:.4f}")
                
                self.value_text.config(state='disabled')
            except tk.TclError as e:
                print(f"Error updating value text: {e}")
            
    def clear_results_and_plots(self):
        # Clear text widgets safely
        text_widgets = [
            'equation_text', 
            'factor_definitions_text',
            'factors_text', 
            'value_text', 
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
        # Clear previous widgets if they exist
        if hasattr(self, 'individual_results_frame'):
            try:
                for widget in self.individual_results_frame.winfo_children():
                    widget.destroy()
            except tk.TclError:
                pass
        
        if hasattr(self, 'r2_frame'):
            try:
                for widget in self.r2_frame.winfo_children():
                    widget.destroy()
            except tk.TclError:
                pass

        # Rest of the method remains the same...
        # Generate equation string for each individual result
        equation_parts = []
        for result_col in self.result_functions:
            # Get the CSR coefficients and bits array for this result
            func_data = self.result_functions[result_col]
            beta = func_data['coefficients']
            bits_array = func_data['bits_array']
            n_factors = bits_array.shape[1]
            
            # Generate the equation string for this result
            eq_str = self._generate_single_result_equation(beta, bits_array, n_factors)
            
            # Add the normalization part
            min_val = result_min_max[result_col]['min']
            max_val = result_min_max[result_col]['max']
            short_name = self.col_name_mapping.get(result_col, result_col)
            equation_parts.append(f"({eq_str}-{min_val:.4f})/({max_val-min_val:.4f})")
        
        # Combine all equations
        if not equation_parts:
            equation_str = "Comprehensive CSR = (No results)"
        else:
            equation_str = "Comprehensive CSR = " + " + ".join(equation_parts)
        
        # Update Equation Text
        self.equation_text.config(state='normal')
        self.equation_text.delete(1.0, tk.END)
        self.equation_text.insert(tk.END, equation_str)
        self.equation_text.config(state='disabled')
        
        # Update factor definitions
        factor_definitions = []
        if hasattr(self, 'factor_cols') and self.factor_cols:
            for i, factor in enumerate(self.factor_cols):
                short_name = f"f{i+1}"
                original_name = self.col_name_mapping.get(factor, f"Factor {i+1}")
                factor_definitions.append(f"{short_name} = {original_name}")
        
        factor_defs_str = "\n".join(factor_definitions)
        
        self.factor_definitions_text.config(state='normal')
        self.factor_definitions_text.delete(1.0, tk.END)
        self.factor_definitions_text.insert(tk.END, factor_defs_str)
        self.factor_definitions_text.config(state='disabled')
        
        # Update extremum display - show real factor values
        self.factors_text.config(state='normal')
        self.factors_text.delete(1.0, tk.END)
        self.value_text.config(state='normal')
        self.value_text.delete(1.0, tk.END)

        if self.extremum_point and 'x' in self.extremum_point and self.extremum_point['x'] is not None:
            # For comprehensive optimization, extremum point is already in original scale
            extremum_x_original_scale = self.extremum_point['x']
            if extremum_x_original_scale is not None and len(extremum_x_original_scale) > 0:
                self.factors_text.insert(tk.END, ", ".join([f"{val:.4f}" for val in extremum_x_original_scale]))
            self.value_text.insert(tk.END, f"{self.extremum_point['value']:.4f}")

        self.factors_text.config(state='disabled')
        self.value_text.config(state='disabled')

        # Add individual results
        if hasattr(self, 'result_functions'):
            # Create a grid layout for individual results
            row_num = 0
            for result_col, func_data in self.result_functions.items():
                display_name = self.col_name_mapping.get(result_col, result_col)
                
                # Result frame for each individual CSR
                result_frame = ttk.Frame(self.individual_results_frame)
                result_frame.grid(row=row_num, column=0, sticky='ew', pady=(0,5))
                
                # Display name label
                ttk.Label(result_frame, text=f"{display_name}:", font=self.label_font).grid(row=0, column=0, sticky='w')
                
                # Calculate extremum value for this individual CSR
                x_norm = None
                if func_data['norm_type'] == "[-1, 1]":
                    x_norm = 2 * (self.extremum_point['x'] - func_data['x_min']) / (func_data['x_max'] - func_data['x_min']) - 1
                elif func_data['norm_type'] == "[0, 1]":
                    x_norm = (self.extremum_point['x'] - func_data['x_min']) / (func_data['x_max'] - func_data['x_min'])
                else:
                    x_norm = self.extremum_point['x']
                
                design_row = self.create_design_matrix(x_norm.reshape(1, -1), func_data['bits_array'])
                val = np.dot(design_row[0], func_data['coefficients'])
                
                # Value at extremum
                val_text = tk.Text(result_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                 font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                val_text.grid(row=0, column=1, padx=(5,15), sticky='w')
                val_text.config(state='normal')
                val_text.insert(tk.END, f"{val:.4f}")
                val_text.config(state='disabled')
                
                row_num += 1
            
            # Add R² values in a similar grid layout
            r2_row_num = 0
            for result_col, func_data in self.result_functions.items():
                display_name = self.col_name_mapping.get(result_col, result_col)
                r2 = func_data['model'].score(func_data['X_design'], func_data['y'])
                
                r2_frame = ttk.Frame(self.r2_frame)
                r2_frame.grid(row=r2_row_num, column=0, sticky='ew', pady=(0,5))
                
                ttk.Label(r2_frame, text=f"{display_name}:", font=self.label_font).grid(row=0, column=0, sticky='w')
                
                r2_text = tk.Text(r2_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                                font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                r2_text.grid(row=0, column=1, padx=(5,0), sticky='w')
                r2_text.config(state='normal')
                r2_text.insert(tk.END, f"{r2:.4f}")
                r2_text.config(state='disabled')
                
                # Add interpretation
                interpretation = ""
                if r2 >= 0.9:
                    interpretation = "Excellent fit"
                elif r2 >= 0.7:
                    interpretation = "Good fit"
                elif r2 >= 0.5:
                    interpretation = "Moderate fit"
                else:
                    interpretation = "Poor fit"
                
                interp_text = tk.Text(r2_frame, height=1, width=20, wrap=tk.NONE, state='disabled',
                                    font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
                interp_text.grid(row=0, column=2, padx=(5,0), sticky='w')
                interp_text.config(state='normal')
                interp_text.insert(tk.END, interpretation)
                interp_text.config(state='disabled')
                
                r2_row_num += 1
            
            # Calculate and display average R²
            avg_r2 = np.mean([func['model'].score(func['X_design'], func['y']) for func in self.result_functions.values()])
            
            avg_frame = ttk.Frame(self.r2_frame)
            avg_frame.grid(row=r2_row_num, column=0, sticky='ew', pady=(10,0))
            
            ttk.Label(avg_frame, text="Average R²:", font=(self.label_font.cget("family"), 
                     self.label_font.cget("size"), "bold")).grid(row=0, column=0, sticky='w')
            
            avg_text = tk.Text(avg_frame, height=1, width=15, wrap=tk.NONE, state='disabled',
                             font=self.text_widget_font, relief=tk.SOLID, borderwidth=1, padx=5)
            avg_text.grid(row=0, column=1, padx=(5,0), sticky='w')
            avg_text.config(state='normal')
            avg_text.insert(tk.END, f"{avg_r2:.4f}")
            avg_text.config(state='disabled')

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
                fixed_factor_values_original_scale = self.extremum_point['x']

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

            # Plot extremum point if available
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                ax2.scatter([self.extremum_point['x'][x_idx]], 
                        [self.extremum_point['x'][y_idx]], 
                        [self.extremum_point['value']], 
                        c='gold', s=200, marker='*', edgecolor='black', 
                        linewidth=1, label='Extremum', depthshade=True, zorder=10)

            # Set axis labels
            x_axis_name = self.col_name_mapping.get(self.factor_cols[x_idx], self.factor_cols[x_idx])
            y_axis_name = self.col_name_mapping.get(self.factor_cols[y_idx], self.factor_cols[y_idx])
            result_axis_name = "Combined Result" if hasattr(self, 'comprehensive_function') else self.col_name_mapping.get("result", "Result")
            
            ax2.set_xlabel(f"\n{x_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_ylabel(f"\n{y_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_zlabel(f"\n{result_axis_name}", fontsize=9, fontweight='bold', linespacing=2)
            ax2.set_title('CSR Response Surface', fontsize=11, fontweight='bold', y=1.02)

            # Add legend if extremum point is plotted
            if (self.extremum_point and self.extremum_point['x'] is not None and 
                len(self.extremum_point['x']) == n_factors):
                ax2.legend(fontsize=8, facecolor='#F0F0F0', framealpha=0.8)

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

        res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, method='L-BFGS-B')
        return {'x': res.x, 'value': objective_func_val(res.x)}


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

        # Perform the optimization - note bounds are in original scale
        res = minimize(objective_to_minimize, x0_for_opt, bounds=bounds_for_opt, method='L-BFGS-B')
        
        return {
            'x': res.x,  # Already in original scale
            'value': comprehensive_func_for_optimizer(res.x)
        }

    def _normalize_point(self, x_point_original_scale):
        if x_point_original_scale is None: return None
        x_original_np = np.array(x_point_original_scale, dtype=float)
        if self.norm_select.get() == "None" or self.norm_x_min is None or self.norm_x_max is None:
            return x_original_np
        if len(x_original_np) != len(self.norm_x_min): return None
        range_val = self.norm_x_max - self.norm_x_min
        range_val[range_val == 0] = 1
        if self.norm_select.get() == "[-1, 1]":
            return 2 * (x_original_np - self.norm_x_min) / range_val - 1
        elif self.norm_select.get() == "[0, 1]":
            return (x_original_np - self.norm_x_min) / range_val
        return x_original_np

    def _unnormalize_point(self, x_point_fitting_scale):
        if x_point_fitting_scale is None: return None
        x_fitting_np = np.array(x_point_fitting_scale, dtype=float)
        if self.norm_select.get() == "None" or self.norm_x_min is None or self.norm_x_max is None:
            return x_fitting_np
        if len(x_fitting_np) != len(self.norm_x_min): return None
        range_val = self.norm_x_max - self.norm_x_min
        if self.norm_select.get() == "[-1, 1]":
            return (x_fitting_np + 1) / 2 * range_val + self.norm_x_min
        elif self.norm_select.get() == "[0, 1]":
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
                short_name = f"f{i+1}"
                original_name = self.col_name_mapping.get(
                    self.factor_cols[i] if i < len(self.factor_cols) else f"Factor {i+1}",
                    f"Factor {i+1}"
                )
                factor_definitions.append(f"{short_name} = {original_name}")
        elif n_factors > 0:
            for i in range(n_factors):
                factor_definitions.append(f"f{i+1} = Factor {i+1}")

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
                term_parts.append(f"f{j+1}")
            elif bits[j] == 2:
                term_parts.append(f"f{j+1}²")
        
        # Special handling for interaction terms to ensure proper ordering
        if len(term_parts) == 2 and sum(bits) == 2 and 2 not in bits:
            # Sort interaction terms (f1×f2, not f2×f1)
            term_parts.sort(key=lambda x: int(x.replace('f', '').replace('²', '')))
        
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
                            c=colors[i], label=self.col_name_mapping.get(result_col, result_col))
            
            min_val = min(min(func['y'].min(), func['model'].predict(func['X_design']).min()) 
                       for func in self.result_functions.values())
            max_val = max(max(func['y'].max(), func['model'].predict(func['X_design']).max()) 
                       for func in self.result_functions.values())
            
            ax1.plot([min_val, max_val], [min_val, max_val], 'r--', lw=1.5, label='Perfect Fit')
            ax1.legend(fontsize=9)
        elif self.y is not None and self.y_pred is not None:
            # Single result case
            ax1.scatter(self.y, self.y_pred, alpha=0.7, edgecolors='#333333', s=40, c="#007ACC")
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
                short_name = f"f{i+1}"
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
                name = f"f{idx+1}"
                term_contributions_map['linear_total'] += contribution_val
                term_contributions_map['linear_factors'][name] = term_contributions_map['linear_factors'].get(name, 0) + contribution_val
            elif sum_of_powers == 2 and 2 in unique_powers:
                idx = np.where(bits_def == 2)[0][0]
                name = f"f{idx+1}²"
                term_contributions_map['quadratic_total'] += contribution_val
                term_contributions_map['quadratic_factors'][name] = term_contributions_map['quadratic_factors'].get(name, 0) + contribution_val
            elif sum_of_powers == 2 and unique_powers.issubset({0, 1}):
                idxs = np.where(bits_def == 1)[0]
                if len(idxs) == 2:
                    idx1, idx2 = sorted(idxs)
                    name1 = f"f{idx1+1}"
                    name2 = f"f{idx2+1}"
                    d_name = f"{name1}×{name2}"
                    term_contributions_map['interaction_total'] += contribution_val
                    term_contributions_map['interaction_factors'][d_name] = term_contributions_map['interaction_factors'].get(d_name, 0) + contribution_val

        # Sort terms for consistent display
        def sort_linear_key(item):
            return int(item[0].replace('f', ''))
        
        def sort_quadratic_key(item):
            return int(item[0].replace('f', '').replace('²', ''))
        
        def sort_interaction_key(item):
            parts = item[0].split('×')
            return (int(parts[0].replace('f', '')), int(parts[1].replace('f', '')))

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

if __name__ == "__main__":
    root = tk.Tk()
    app = CSRApp(root)
    root.mainloop()
