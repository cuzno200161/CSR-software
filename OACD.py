class OACD:
    def __init__(self, root):
        self.root = root
        self.root.title("OACD Table Builder")
        self.root.geometry("800x600")
        
        self.factor_num = None;
        self.table_size = None;
        
    def set_factor_num(self, factor_num):
        self.factor_num = factor_num;
        
    def set_table_size(self, table_size):
        self.table_size = table_size;
        
    def build_table(self):
        pass;
    
    