class Integration:
    def __init__(self, root):
        self.root = root
        
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
        
    
        
    