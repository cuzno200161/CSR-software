import pandas as pd
import numpy as np
from scipy.optimize import minimize
from sklearn.linear_model import Ridge
from sklearn.model_selection import cross_val_score
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.figure import Figure

class CSRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Python CSR Integration App")
        self.root.geometry("1400x800")
        
        self.df = None
        self.coefficients = None
        self.bits_array = None
        self.factor_cols = []
        self.X = None
        self.y = None
        self.y_pred = None
        self.col_name_mapping = {}
        self.original_col_names = []
        self.extremum_point = None
        
        self.create_widgets()
    
    def create_widgets(self):
        """Create all GUI components with balanced layout"""
        self.notebook = ttk.Notebook(self.root)
        
        # Create tabs
        self.tab1 = ttk.Frame(self.notebook)
        self.tab2 = ttk.Frame(self.notebook)
        self.tab3 = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab1, text='CSR Integration')
        self.notebook.add(self.tab2, text='Coefficient Analysis')
        self.notebook.add(self.tab3, text='OACD Table Builder')
        self.notebook.pack(expand=1, fill='both')
        
        # Main paned window (left, center, right)
        self.main_paned = tk.PanedWindow(self.tab1, orient=tk.HORIZONTAL, sashrelief=tk.RAISED, sashwidth=8)
        self.main_paned.pack(fill=tk.BOTH, expand=True)
        
        # Left panel (25%) - Controls
        left_frame = tk.Frame(self.main_paned, width=350)
        self.main_paned.add(left_frame)
        
        # Center panel (50%) - Table
        center_frame = tk.Frame(self.main_paned, width=700)
        self.main_paned.add(center_frame)
        
        # Right panel (25%) - Plots (split vertically)
        right_paned = tk.PanedWindow(self.main_paned, orient=tk.VERTICAL, sashrelief=tk.RAISED, sashwidth=8)
        self.main_paned.add(right_paned)
        
        # Configure pane weights
        self.main_paned.paneconfig(left_frame, minsize=350, width=350)
        self.main_paned.paneconfig(center_frame, minsize=700, width=700)
        
        # === Left Panel Contents ===
        # File selection
        tk.Button(left_frame, text="Select File", command=self.select_file).pack(pady=5)
        
        # Project name
        tk.Label(left_frame, text="Project Name").pack()
        self.project_name_entry = tk.Entry(left_frame)
        self.project_name_entry.pack(fill='x')
        
        # Fitting parameters
        tk.Label(left_frame, text="Weight Selection").pack()
        self.weight_combo = ttk.Combobox(left_frame, 
                                      values=["Maximum", "Minimum", 
                                             "Maximum absolute value", 
                                             "Minimum absolute value"])
        self.weight_combo.pack(fill='x')
        
        tk.Label(left_frame, text="Normalization Standard").pack()
        self.norm_select = ttk.Combobox(left_frame, values=["None", "[-1, 1]", "[0, 1]"])
        self.norm_select.pack(fill='x')
        
        # Fitting button
        tk.Button(left_frame, text="Start Fitting", command=self.run_fitting).pack(pady=10)
        
        # Results display
        tk.Label(left_frame, text="CSR Function").pack()
        
        eqn_frame = tk.Frame(left_frame)
        eqn_frame.pack(fill='both', expand=True)
        scrollbar = tk.Scrollbar(eqn_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.equation_text = tk.Text(eqn_frame, height=8, width=30, wrap=tk.WORD,
                                   yscrollcommand=scrollbar.set, state='disabled')
        self.equation_text.pack(side=tk.LEFT, fill='both', expand=True)
        scrollbar.config(command=self.equation_text.yview)
        
        # Extremum and R² display (updated for copyable content)
        results_frame = tk.Frame(left_frame)
        results_frame.pack(fill='x', pady=5)
        
        tk.Label(results_frame, text="Extremum Point:").pack(anchor='w')
        
        # Frame for factors display
        factors_frame = tk.Frame(results_frame)
        factors_frame.pack(fill='x', padx=5, pady=2)
        tk.Label(factors_frame, text="Factors:").pack(side='left')
        self.factors_text = tk.Text(factors_frame, height=1, width=25, wrap=tk.NONE)
        self.factors_text.pack(side='left', fill='x', expand=True)
        self.factors_text.config(state='disabled')
        
        # Frame for value display
        value_frame = tk.Frame(results_frame)
        value_frame.pack(fill='x', padx=5, pady=2)
        tk.Label(value_frame, text="Value:").pack(side='left')
        self.value_text = tk.Text(value_frame, height=1, width=25, wrap=tk.NONE)
        self.value_text.pack(side='left', fill='x', expand=True)
        self.value_text.config(state='disabled')
        
        # Frame for R² display
        r2_frame = tk.Frame(results_frame)
        r2_frame.pack(fill='x', padx=5, pady=2)
        tk.Label(r2_frame, text="R²:").pack(side='left')
        self.r2_text = tk.Text(r2_frame, height=1, width=25, wrap=tk.NONE)
        self.r2_text.pack(side='left', fill='x', expand=True)
        self.r2_text.config(state='disabled')
        
        # === Center Panel Contents ===
        table_container = tk.Frame(center_frame)
        table_container.pack(fill='both', expand=True)
        
        # Scrollbars
        x_scroll = ttk.Scrollbar(table_container, orient='horizontal')
        x_scroll.pack(side='bottom', fill='x')
        y_scroll = ttk.Scrollbar(table_container)
        y_scroll.pack(side='right', fill='y')
        
        # Treeview table
        self.table = ttk.Treeview(table_container,
                                xscrollcommand=x_scroll.set,
                                yscrollcommand=y_scroll.set,
                                height=20)
        self.table.pack(side='left', fill='both', expand=True)
        x_scroll.config(command=self.table.xview)
        y_scroll.config(command=self.table.yview)
        
        # Style
        style = ttk.Style()
        style.configure("Treeview", rowheight=25, font=('Arial', 10))
        style.configure("Treeview.Heading", font=('Arial', 10, 'bold'))
        
        # Regularization controls
        reg_frame = tk.LabelFrame(left_frame, text="Regularization Parameters")
        reg_frame.pack(fill='x', pady=5)

        # Alpha (penalty) slider
        tk.Label(reg_frame, text="Alpha (penalty):").grid(row=0, column=0, sticky='w')
        self.alpha_slider = tk.Scale(
            reg_frame, 
            from_=-4,  # 10^-4 = 0.0001
            to=2,      # 10^2 = 100
            resolution=0.01,  # Fine-grained control
            orient=tk.HORIZONTAL,
            command=self.update_alpha_value
        )
        self.alpha_slider.set(0)  # Default alpha=10^0=1.0
        self.alpha_value_label = tk.Label(reg_frame, text="1.00")
        self.alpha_value_label.grid(row=0, column=2, sticky='w')
        self.alpha_slider.grid(row=0, column=1, sticky='ew')

        # Max iterations slider
        tk.Label(reg_frame, text="Max iterations:").grid(row=1, column=0, sticky='w')
        self.max_iter_slider = tk.Scale(reg_frame, from_=2, to=5, 
                                       resolution=0.01, orient=tk.HORIZONTAL,
                                       command=self.update_max_iter_value)
        self.max_iter_slider.set(np.log10(1000))  # Default value 1000 (10^3)
        self.max_iter_value_label = tk.Label(reg_frame, text="1000")
        self.max_iter_value_label.grid(row=1, column=2, sticky='w')
        self.max_iter_slider.grid(row=1, column=1, sticky='ew')

        # Early stopping checkbox
        tk.Label(reg_frame, text="Early stopping:").grid(row=2, column=0, sticky='w')
        self.early_stop_var = tk.BooleanVar(value=True)
        tk.Checkbutton(reg_frame, variable=self.early_stop_var).grid(row=2, column=1, sticky='w')

        reg_frame.columnconfigure(1, weight=1)  # Make slider fields expandable

        # === Right Panel Contents ===
        # Top plot (50% height)
        plot1_frame = tk.Frame(right_paned)
        right_paned.add(plot1_frame)
        tk.Label(plot1_frame, text="Interaction Surface").pack()
        self.figure1 = Figure(figsize=(6, 4), dpi=100)
        self.canvas1 = FigureCanvasTkAgg(self.figure1, master=plot1_frame)
        self.canvas1.get_tk_widget().pack(fill='both', expand=True)
        toolbar1 = NavigationToolbar2Tk(self.canvas1, plot1_frame)
        toolbar1.update()
        
        # Bottom plot (50% height)
        plot2_frame = tk.Frame(right_paned)
        right_paned.add(plot2_frame)
        
        # Factor selection controls
        factor_control_frame = tk.Frame(plot2_frame)
        factor_control_frame.pack(fill='x', pady=5)
        tk.Label(factor_control_frame, text="X-axis Factor:").pack(side='left', padx=5)
        self.x_factor_combo = ttk.Combobox(factor_control_frame, state='readonly')
        self.x_factor_combo.pack(side='left', padx=5)
        tk.Label(factor_control_frame, text="Y-axis Factor:").pack(side='left', padx=5)
        self.y_factor_combo = ttk.Combobox(factor_control_frame, state='readonly')
        self.y_factor_combo.pack(side='left', padx=5)
        tk.Button(factor_control_frame, text="Update Plot", 
                command=self.update_3d_plot).pack(side='left', padx=10)
        
        tk.Label(plot2_frame, text="CSR Surface").pack()
        self.figure2 = Figure(figsize=(6, 4), dpi=100)
        self.canvas2 = FigureCanvasTkAgg(self.figure2, master=plot2_frame)
        self.canvas2.get_tk_widget().pack(fill='both', expand=True)
        toolbar2 = NavigationToolbar2Tk(self.canvas2, plot2_frame)
        toolbar2.update()
        
    def select_file(self):
        """Load experimental data from file while preserving custom column names"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        
        try:
            # Read the file with headers
            self.df = pd.read_excel(file_path, header=0)
            
            # Store original column names for display
            self.original_col_names = list(self.df.columns)
            
            # Rename columns to standard format internally
            n_cols = len(self.df.columns)
            new_col_names = [f"factor{i+1}" for i in range(n_cols-1)] + ["result"]
            self.col_name_mapping = dict(zip(new_col_names, self.df.columns))
            self.df.columns = new_col_names
            
            # Verify we have data
            if self.df.empty:
                raise ValueError("No data found in the file")
                
            # Ensure numeric data and clean
            self.df = self.df.apply(pd.to_numeric, errors='coerce').dropna()
            
            # Add residual column
            self.df["residual"] = 0
            
            # Update display
            self.project_name_entry.delete(0, tk.END)
            self.project_name_entry.insert(0, file_path)
            self.update_table_view()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")

    def update_table_view(self):
        """Update table view with original column names"""
        self.table.delete(*self.table.get_children())
        
        if self.df is None or self.df.empty:
            return

        # Configure columns with original names
        display_cols = [self.col_name_mapping.get(col, col) for col in self.df.columns]
        self.table["columns"] = list(self.df.columns)
        self.table.column("#0", width=0, stretch=tk.NO)
        
        # Set column headings to show original names
        for internal_name, display_name in zip(self.df.columns, display_cols):
            self.table.heading(internal_name, text=display_name)
            self.table.column(internal_name, width=100, anchor='center', stretch=tk.NO)

        # Insert data with formatting
        for _, row in self.df.iterrows():
            formatted_values = [
                f"{x:.4f}" if isinstance(x, (float, np.floating)) else str(x)
                for x in row
            ]
            self.table.insert("", "end", values=formatted_values)

    def update_alpha_value(self, value):
        """Convert log-scale slider value to actual alpha with proper formatting"""
        try:
            alpha = 10 ** float(value)
            if alpha < 0.01:
                # Use scientific notation for very small values
                self.alpha_value_label.config(text=f"{alpha:.1e}")
            else:
                # Show 4 decimal places for larger values
                self.alpha_value_label.config(text=f"{alpha:.2f}")
        except:
            self.alpha_value_label.config(text="Error")

    def update_max_iter_value(self, value):
        """Update max iterations value display when slider moves"""
        try:
            max_iter = int(10 ** float(value))
            self.max_iter_value_label.config(text=f"{max_iter}")
        except:
            pass

    def run_fitting(self):
        """Perform CSR fitting with Ridge Regression and cross-validation"""
        if self.df is None:
            messagebox.showerror("Error", "Please load data first")
            return

        try:
            # Get factors and result
            self.factor_cols = [col for col in self.df.columns if col.startswith('factor')]
            if not self.factor_cols:
                messagebox.showerror("Error", "No factor columns found")
                return
                
            X = self.df[self.factor_cols].values
            y = self.df['result'].values
            n_factors = len(self.factor_cols)
            
            # Normalize data if requested
            if self.norm_select.get() == "[-1, 1]":
                self.x_min = X.min(axis=0)
                self.x_max = X.max(axis=0)
                X = 2 * (X - self.x_min) / (self.x_max - self.x_min) - 1
            elif self.norm_select.get() == "[0, 1]":
                self.x_min = X.min(axis=0)
                self.x_max = X.max(axis=0)
                X = (X - self.x_min) / (self.x_max - self.x_min)
            
            # Generate binary combinations matrix
            self.bits_array = self.generate_bits_array(n_factors)
            if self.bits_array.size == 0:
                messagebox.showerror("Error", "Could not generate terms matrix")
                return
            
            # Create design matrix
            X_design = self.create_design_matrix(X, self.bits_array)
            
            # Get user-selected regularization parameters
            try:
                alpha = 10 ** float(self.alpha_slider.get())
                max_iter = int(10 ** float(self.max_iter_slider.get()))
            except:
                messagebox.showerror("Error", "Invalid regularization parameters")
                return

            # Initialize and fit Ridge model
            model = Ridge(alpha=alpha, max_iter=max_iter, fit_intercept=False, random_state=42)
            
            # Perform 5-fold cross-validation
            cv_scores = cross_val_score(model, X_design, y, cv=5, scoring='r2')
            avg_r2 = np.mean(cv_scores)
            std_r2 = np.std(cv_scores)
            
            # Fit model on full data
            model.fit(X_design, y)
            self.coefficients = model.coef_
            
            # Calculate predictions and residuals
            y_pred = model.predict(X_design)
            residuals = y - y_pred
            self.df['residual'] = residuals
            
            # Calculate training R-squared
            train_r2 = model.score(X_design, y)
            
            # Find extremum point
            extremum_type = self.weight_combo.get().lower().replace(" ", "_")
            self.extremum_point = self.find_extremum(self.coefficients, self.bits_array, 
                                                   X.min(axis=0), X.max(axis=0), 
                                                   extremum_type)
            
            # Generate equation string
            eqn = self.generate_equation(self.coefficients, self.bits_array, n_factors)
            
            # Update results display
            self.update_results_display(eqn, train_r2, avg_r2, std_r2)
            
            # Update table with residuals
            self.update_table_view()
            
            # Update factor selection comboboxes
            display_names = [self.col_name_mapping.get(f"factor{i+1}", f"factor{i+1}") 
                           for i in range(len(self.factor_cols))]
            self.x_factor_combo['values'] = display_names
            self.y_factor_combo['values'] = display_names
            self.x_factor_combo.current(0)
            self.y_factor_combo.current(1 if len(display_names) > 1 else 0)
            
            # Store data for plotting
            self.X = X
            self.y = y
            self.y_pred = y_pred
            
            # Plot results with default factors
            self.plot_results(X, y, y_pred, n_factors, 0, 1 if n_factors > 1 else 0)
          
        except Exception as e:
            messagebox.showerror("Error", f"Fitting failed: {str(e)}")

    def update_results_display(self, eqn, train_r2, avg_r2, std_r2):
        """Update all result displays"""
        self.equation_text.config(state='normal')
        self.equation_text.delete(1.0, tk.END)
        self.equation_text.insert(tk.END, eqn)
        self.equation_text.config(state='disabled')

        # Update extremum point display
        self.factors_text.config(state='normal')
        self.factors_text.delete(1.0, tk.END)
        if self.extremum_point:
            self.factors_text.insert(tk.END, ", ".join([f"{x:.4f}" for x in self.extremum_point['x']]))
        self.factors_text.config(state='disabled')

        self.value_text.config(state='normal')
        self.value_text.delete(1.0, tk.END)
        if self.extremum_point:
            self.value_text.insert(tk.END, f"{self.extremum_point['value']:.4f}")
        self.value_text.config(state='disabled')

        # Update R² display with both training and CV R²
        self.r2_text.config(state='normal')
        self.r2_text.delete(1.0, tk.END)
        self.r2_text.insert(tk.END, f"Train: {train_r2:.3f}\nCV: {avg_r2:.3f}±{std_r2:.3f}")
        self.r2_text.config(state='disabled')

    def update_3d_plot(self):
        """Update 3D plot based on selected factors"""
        if self.df is None or self.coefficients is None:
            return
            
        x_factor = self.x_factor_combo.get()
        y_factor = self.y_factor_combo.get()
        
        if not x_factor or not y_factor:
            return
            
        try:
            # Get display names mapping
            display_to_internal = {v: k for k, v in self.col_name_mapping.items()}
            
            # Get factor indices
            x_internal = display_to_internal.get(x_factor, x_factor)
            y_internal = display_to_internal.get(y_factor, y_factor)
            
            x_idx = int(x_internal.replace("factor", "")) - 1
            y_idx = int(y_internal.replace("factor", "")) - 1
            
            # Replot with selected factors
            self.plot_results(self.X, self.y, self.y_pred, len(self.factor_cols), x_idx, y_idx)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update plot: {str(e)}")

    def generate_bits_array(self, n_factors):
        """Generate binary combinations matrix for constant, linear, and quadratic terms"""
        if n_factors == 0:
            return np.zeros((1, 0))  # Handle empty case
        
        bits = []
        
        # 1. Constant term (all zeros)
        bits.append(np.zeros(n_factors))
        
        # 2. Linear terms
        bits.extend(np.eye(n_factors))
        
        # 3. Quadratic and interaction terms
        for i in range(n_factors):
            for j in range(i, n_factors):
                term = np.zeros(n_factors)
                if i == j:
                    term[i] = 2  # Quadratic term
                else:
                    term[i] = 1
                    term[j] = 1  # Interaction term
                bits.append(term)
        
        return np.array(bits)

    def create_design_matrix(self, X, bits_array):
        """Create design matrix with constant term in first column"""
        if X.size == 0 or bits_array.size == 0:
            return np.zeros((0, 0))
        
        n_samples = X.shape[0]
        X_design = np.ones((n_samples, bits_array.shape[0]))  # First column is 1s for intercept
        
        for i, bits in enumerate(bits_array):
            term = np.ones(n_samples)
            for j in range(len(bits)):
                if bits[j] == 1:
                    term *= X[:,j]
                elif bits[j] == 2:
                    term *= X[:,j] * X[:,j]
            X_design[:,i] = term
        
        return X_design

    def find_extremum(self, beta, bits_array, x_min, x_max, extremum_type):
        """Find extremum of CSR function using optimization"""
        def csr_func(x):
            result = 0
            for i, bits in enumerate(bits_array):
                term = 1
                for j in range(len(bits)):
                    if bits[j] == 1:
                        term *= x[j]
                    elif bits[j] == 2:
                        term *= x[j] * x[j]
                result += beta[i] * term
            return result
        
        # Define objective based on extremum type
        if extremum_type == 'maximum':
            objective = lambda x: -csr_func(x)
        elif extremum_type == 'minimum':
            objective = csr_func
        elif extremum_type == 'maximum_absolute_value':
            objective = lambda x: -abs(csr_func(x))
        elif extremum_type == 'minimum_absolute_value':
            objective = lambda x: abs(csr_func(x))
        else:
            objective = csr_func  # Default to minimization
        
        # Initial guess (center of bounds)
        x0 = (x_min + x_max) / 2
        
        # Run optimization
        bounds = [(low, high) for low, high in zip(x_min, x_max)]
        res = minimize(objective, x0, bounds=bounds)
        
        return {'x': res.x, 'value': csr_func(res.x)}

    def generate_equation(self, beta, bits_array, n_factors):
        """Generate human-readable equation string with proper formatting"""
        if bits_array.size == 0 or len(beta) == 0:
            return "y = 0"  # Default equation if no terms
        
        terms = []
        for i, (coef, bits) in enumerate(zip(beta, bits_array)):
            if i >= len(bits_array) or i >= len(beta):
                continue  # Skip if indices are out of bounds
            if abs(coef) < 1e-6:  # Skip near-zero terms
                continue
                
            term_parts = []
            for j in range(min(len(bits), n_factors)):  # Ensure we don't go out of bounds
                if bits[j] == 1:
                    term_parts.append(f"x{j+1}")
                elif bits[j] == 2:  # Quadratic term
                    term_parts.append(f"x{j+1}²")
            
            # Handle sign display
            sign = "+" if coef >= 0 else "-"
            abs_coef = abs(coef)
            
            if not term_parts:  # Constant term
                term = f"{sign} {abs_coef:.4f}"
            else:
                # For quadratic terms we don't need the * symbol
                if len(term_parts) == 1 and 2 in bits:
                    term = f"{sign} {abs_coef:.4f}{term_parts[0]}"
                else:
                    term = f"{sign} {abs_coef:.4f}*{'*'.join(term_parts)}"
            
            terms.append(term)
        
        # Build final equation
        if not terms:
            return "y = 0"
        
        equation = "y = " + " ".join(terms).replace("+ -", "- ")
        
        # Clean up leading operator
        if equation.startswith("y = +"):
            equation = equation[5:]
            equation = "y = " + equation
        return equation

    def calculate_r2(self, y_true, y_pred):
        """Calculate R-squared value"""
        ss_res = np.sum((y_true - y_pred) ** 2)
        ss_tot = np.sum((y_true - np.mean(y_true)) ** 2)
        return 1 - (ss_res / ss_tot) if ss_tot != 0 else 0

    def plot_results(self, X_original, y, y_pred, n_factors, x_idx=0, y_idx=1):
        """Plot only the CSR surface and extremum point"""
        self.figure1.clf()
        self.figure2.clf()

        # 1. Actual vs Predicted plot (unchanged)
        ax1 = self.figure1.add_subplot(111)
        ax1.scatter(y, y_pred, alpha=0.6, label='Data Points')
        ax1.plot([y.min(), y.max()], [y.min(), y.max()], 'r--', label='Perfect Fit')
        ax1.set_xlabel('Actual Values', fontsize=10)
        ax1.set_ylabel('Predicted Values', fontsize=10)
        ax1.set_title('Actual vs Predicted', fontsize=12)
        ax1.grid(True)
        ax1.legend()
        self.canvas1.draw()

        # 2. Clean 3D Surface plot (no data points)
        if n_factors >= 2:
            ax2 = self.figure2.add_subplot(111, projection='3d')

            # Create grid based on original ranges
            x1 = np.linspace(X_original[:,x_idx].min(), X_original[:,x_idx].max(), 20)
            x2 = np.linspace(X_original[:,y_idx].min(), X_original[:,y_idx].max(), 20)
            x1_grid, x2_grid = np.meshgrid(x1, x2)

            # Get extremum values (original scale)
            extremum_original = self.extremum_point['x'] if self.extremum_point else np.mean(X_original, axis=0)

            # Evaluate CSR function with un-normalization
            z_grid = np.zeros_like(x1_grid)
            for i in range(x1_grid.shape[0]):
                for j in range(x1_grid.shape[1]):
                    point_normalized = np.zeros(n_factors)
                    # Normalize the grid point temporarily to use the fitted model
                    if hasattr(self, 'x_min') and hasattr(self, 'x_max') and self.norm_select.get() == "[-1, 1]":
                        point_normalized[x_idx] = 2 * (x1_grid[i, j] - self.x_min[x_idx]) / (self.x_max[x_idx] - self.x_min[x_idx]) - 1
                        point_normalized[y_idx] = 2 * (x2_grid[i, j] - self.x_min[y_idx]) / (self.x_max[y_idx] - self.x_min[y_idx]) - 1
                        for k in range(n_factors):
                            if k != x_idx and k != y_idx:
                                point_normalized[k] = 2 * (extremum_original[k] - self.x_min[k]) / (self.x_max[k] - self.x_min[k]) - 1
                    elif hasattr(self, 'x_min') and hasattr(self, 'x_max') and self.norm_select.get() == "[0, 1]":
                        point_normalized[x_idx] = (x1_grid[i, j] - self.x_min[x_idx]) / (self.x_max[x_idx] - self.x_min[x_idx])
                        point_normalized[y_idx] = (x2_grid[i, j] - self.x_min[y_idx]) / (self.x_max[y_idx] - self.x_min[y_idx])
                        for k in range(n_factors):
                            if k != x_idx and k != y_idx:
                                point_normalized[k] = (extremum_original[k] - self.x_min[k]) / (self.x_max[k] - self.x_min[k])
                    else:
                        point_normalized[x_idx] = x1_grid[i, j]
                        point_normalized[y_idx] = x2_grid[i, j]
                        for k in range(n_factors):
                            if k != x_idx and k != y_idx:
                                point_normalized[k] = extremum_original[k]

                    z_grid[i, j] = self.evaluate_csr(point_normalized, X_original, y)

            # Plot surface only
            surf = ax2.plot_surface(x1_grid, x2_grid, z_grid,
                                     cmap='viridis', alpha=0.8,
                                     linewidth=0, antialiased=True)

            # Plot extremum point only if available
            if self.extremum_point:
                extremum_x = self.extremum_point['x']
                extremum_val = self.extremum_point['value']
                ax2.scatter([extremum_x[x_idx]], [extremum_x[y_idx]], [extremum_val],
                            c='gold', s=100, marker='*', edgecolor='black',
                            linewidth=1, label='Extremum')

            # Labels
            x_name = self.col_name_mapping.get(f"factor{x_idx+1}", f"factor{x_idx+1}")
            y_name = self.col_name_mapping.get(f"factor{y_idx+1}", f"factor{y_idx+1}")
            result_name = self.col_name_mapping.get("result", "result")

            ax2.set_xlabel(x_name, fontsize=9)
            ax2.set_ylabel(y_name, fontsize=9)
            ax2.set_zlabel(result_name, fontsize=9)
            ax2.set_title('CSR Response Surface', fontsize=11)
            if self.extremum_point:
                ax2.legend()

            self.figure2.colorbar(surf, ax=ax2, shrink=0.5, aspect=5)
            ax2.view_init(elev=30, azim=45)
            self.canvas2.draw()

    def evaluate_csr(self, x_normalized, X_original, y):
        """Evaluate CSR function using fitted coefficients with un-normalization"""
        if not hasattr(self, 'coefficients') or not hasattr(self, 'bits_array'):
            return 0

        n_factors = len(self.factor_cols)
        x_original = np.zeros(n_factors)

        if self.norm_select.get() == "[-1, 1]":
            x_min = self.x_min if hasattr(self, 'x_min') else X_original.min(axis=0)
            x_max = self.x_max if hasattr(self, 'x_max') else X_original.max(axis=0)
            x_original = (x_normalized + 1) / 2 * (x_max - x_min) + x_min
        elif self.norm_select.get() == "[0, 1]":
            x_min = self.x_min if hasattr(self, 'x_min') else X_original.min(axis=0)
            x_max = self.x_max if hasattr(self, 'x_max') else X_original.max(axis=0)
            x_original = x_normalized * (x_max - x_min) + x_min
        else:
            x_original = x_normalized # No normalization

        # Create the design matrix row for the un-normalized point
        design_row = []
        for bits in self.bits_array:
            term = 1
            for j in range(len(bits)):
                if bits[j] == 1:
                    term *= x_original[j]
                elif bits[j] == 2:
                    term *= x_original[j] * x_original[j]
            design_row.append(term)

        return np.dot(design_row, self.coefficients)

if __name__ == "__main__":
    root = tk.Tk()
    app = CSRApp(root)
    root.mainloop()