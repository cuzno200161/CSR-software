import os
import sys
import pandas as pd
import tkinter as tk
import numpy as np

class OACD:
    def __init__(self):
        
        self.factor_num = 2 # 2, 3, 4, 5, 6, 7, 8, 9, 10
        self.table_size = "Small" # Small, Medium, Large
        self.table = None # Pandas DataFrame object
        self.factor_extrenum = None # First column is minimum, second column is maximum; each row is a factor
        self.max_nonzero = None; # Max number of non-zero factors
        self.limits = pd.DataFrame({'limits': [None] * self.factor_num}) # each row is a factor; each factor can be assigned to a limit or None
        self.limit_names = {None : None} # Dictionary of limit names and their display
        
    def __str__(self):
        string = ""
        for row in self.table:
            for cell in row:
                string += cell + " "
            string += "\n"
        return string;
        
    def set_factor_num(self, factor_num):
        self.factor_num = factor_num
        self.factor_extrenum = pd.DataFrame(np.ones((self.factor_num, 2)))
        self.factor_extrenum.iloc[:, 1] = -1; 
        self.limits = pd.DataFrame({'limits': [None] * self.factor_num}) 
        
    def set_table_size(self, table_size):
        self.table_size = table_size
        
    def set_factor_extrenum(self, new_extrenum):
        # Set the entire extrenum table (can be used if importing a table of extrenum)
        self.factor_extrenum = new_extrenum
        
    def set_extrenum(self, value, factor, extrenum):
        # Set a singular extrenum value (can be used for manual input of extrenum)
        if extrenum == "min":
            self.factor_extrenum.iloc[factor, 0] = value;
        elif extrenum == "max":
            self.factor_extrenum.iloc[factor, 1] = value;    
        
    def _resource_path(self, relative_path):
        """
        Resolve resource paths for both development and PyInstaller bundles.
        """
        if hasattr(sys, "_MEIPASS"):
            # PyInstaller onefile
            base_path = sys._MEIPASS
        elif getattr(sys, 'frozen', False):
            # PyInstaller onedir - MacOS .app places resources in Contents/Resources
            # But the executable runs from Contents/MacOS
            # Check if we are in a .app bundle
            exe_dir = os.path.dirname(sys.executable)
            if "Contents/MacOS" in exe_dir:
                # Look in ../Resources relative to the executable
                base_path = os.path.join(os.path.dirname(exe_dir), "Resources")
            else:
                # Standard onedir (e.g. Windows/Linux or non-bundle macOS)
                base_path = exe_dir
        else:
            # Development
            base_path = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(base_path, relative_path)

    def excel_to_python(self, filepath):
        resolved_path = self._resource_path(filepath)
        print(f"DEBUG: Opening Excel file: {filepath} -> {resolved_path}")
        try:
            df = pd.read_excel(resolved_path, header=None)
            return df
        except Exception as e:
            print(f"ERROR: Failed to open {resolved_path}: {e}")
            raise e    
         
    def build_table(self):
        
        # Returns -1 if there table isn't built / error in factor_num
        if(self.table_size != "Small" and self.table_size != "Medium" and self.table_size != "Large" and self.factor_num != None):
            return -1
        
        # Creates a table using a combination of an orthogonal array and central composite design
        match(self.factor_num):
            case 2:
                self.table = self.excel_to_python("OACD_tables/23_OA.xlsx")
            case 3:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/2_3-1.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa9.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :3]], ignore_index=True)
                    case "Medium" | "Large":
                        t1 = self.excel_to_python("OACD_tables/23.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa9.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :3]], ignore_index=True)
            case 4:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/2_4-1.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa9.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :4]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/pb12.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa9.xlsx")
                        t2 = t2.iloc[:, [0, 3, 2, 1]]
                        self.table = pd.concat([t1.iloc[:, :4], t2.iloc[:, :4]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/24.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa9.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :4]], ignore_index=True)
            case 5:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/2_5-2.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2.iloc[:, 0] = t2.iloc[:, 1]
                        t2.iloc[:, 1] = t2.iloc[:, 2]
                        t2.iloc[:, 2] = t2.iloc[:, 3]
                        t2.iloc[:, 3] = t2.iloc[:, 5]
                        self.table = pd.concat([t1, t2.iloc[:, :5]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/pb12.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2.iloc[:, 0] = t2.iloc[:, 1]
                        t2.iloc[:, 1] = t2.iloc[:, 4]
                        t2.iloc[:, 4] = t2.iloc[:, 5]
                        self.table = pd.concat([t1.iloc[:, :5], t2.iloc[:, :5]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_5-1.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2.columns = [col - 1 for col in t2.columns] # Shift columns names to left by 1
                        self.table = pd.concat([t1, t2.iloc[:, 1:6]], ignore_index=True)
            case 6:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/pb12.xlsx")
                        t1.iloc[:, 5] = t1.iloc[:, 6]
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 1]
                        t2r.iloc[:, 1] = t2.iloc[:, 4]
                        t2r.iloc[:, 4] = t2.iloc[:, 5]
                        t2r.iloc[:, 5] = t2.iloc[:, 0]
                        self.table = pd.concat([t1.iloc[:, :6], t2r.iloc[:, :6]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/pb20.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2r = t2
                        t2r.iloc[:, 1] = t2.iloc[:, 3]
                        t2r.iloc[:, 2] = t2.iloc[:, 5]
                        t2r.iloc[:, 3] = t2.iloc[:, 2]
                        t2r.iloc[:, 4] = t2.iloc[:, 1]
                        t2r.iloc[:, 5] = t2.iloc[:, 4]
                        self.table = pd.concat([t1.iloc[:, :6], t2r.iloc[:, :6]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_6-1.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :6]], ignore_index=True)
                        
            case 7:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/pb20.xlsx")
                        t1.iloc[:, 5] = t1.iloc[:, 12]
                        t1.iloc[:, 6] = t1.iloc[:, 15]
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 2]
                        t2r.iloc[:, 1] = t2.iloc[:, 0]
                        t2r.iloc[:, 2] = t2.iloc[:, 4]
                        t2r.iloc[:, 3] = t2.iloc[:, 6]
                        t2r.iloc[:, 4] = t2.iloc[:, 3]
                        t2r.iloc[:, 5] = t2.iloc[:, 1]
                        t2r.iloc[:, 6] = t2.iloc[:, 5]
                        self.table = pd.concat([t1.iloc[:, :7], t2r.iloc[:, :7]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/2_6-2.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2r = t2
                        t2r.iloc[:, 2] = t2.iloc[:, 4]
                        t2r.iloc[:, 3] = t2.iloc[:, 2]
                        t2r.iloc[:, 4] = t2.iloc[:, 3]
                        t2r.iloc[:, 5] = t2.iloc[:, 6]
                        t2r.iloc[:, 6] = t2.iloc[:, 5]
                        self.table = pd.concat([t1, t2r.iloc[:, :7]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_7-1.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :7]], ignore_index=True)
            case 8:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/pb20.xlsx")
                        t1.iloc[:, 5] = t1.iloc[:, 12]
                        t1.iloc[:, 6] = t1.iloc[:, 15]
                        t1.iloc[:, 7] = t1.iloc[:, 14]
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 5]
                        t2r.iloc[:, 1] = t2.iloc[:, 2]
                        t2r.iloc[:, 2] = t2.iloc[:, 7]
                        t2r.iloc[:, 4] = t2.iloc[:, 1]
                        t2r.iloc[:, 5] = t2.iloc[:, 0]
                        t2r.iloc[:, 7] = t2.iloc[:, 4]
                        self.table = pd.concat([t1.iloc[:, :8], t2r.iloc[:, :8]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/2_8-3.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 1] = t2.iloc[:, 2]
                        t2r.iloc[:, 2] = t2.iloc[:, 3]
                        t2r.iloc[:, 3] = t2.iloc[:, 4]
                        t2r.iloc[:, 4] = t2.iloc[:, 1]
                        t2r.iloc[:, 5] = t2.iloc[:, 6]
                        t2r.iloc[:, 6] = t2.iloc[:, 7]
                        t2r.iloc[:, 7] = t2.iloc[:, 5]
                        self.table = pd.concat([t1, t2r.iloc[:, :8]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_8-2.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :8]], ignore_index=True)
            case 9:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/2_9-4.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 4]
                        t2r.iloc[:, 1] = t2.iloc[:, 5]
                        t2r.iloc[:, 2] = t2.iloc[:, 0]
                        t2r.iloc[:, 3] = t2.iloc[:, 6]
                        t2r.iloc[:, 4] = t2.iloc[:, 1]
                        t2r.iloc[:, 5] = t2.iloc[:, 3]
                        t2r.iloc[:, 6] = t2.iloc[:, 8]
                        t2r.iloc[:, 7] = t2.iloc[:, 2]
                        t2r.iloc[:, 8] = t2.iloc[:, 7]
                        self.table = pd.concat([t1, t2r.iloc[:, :9]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/2_9-3.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 1] = t2.iloc[:, 2]
                        t2r.iloc[:, 2] = t2.iloc[:, 7]
                        t2r.iloc[:, 3] = t2.iloc[:, 1]
                        t2r.iloc[:, 4] = t2.iloc[:, 5]
                        t2r.iloc[:, 5] = t2.iloc[:, 6]
                        t2r.iloc[:, 6] = t2.iloc[:, 4]
                        t2r.iloc[:, 7] = t2.iloc[:, 3]
                        self.table = pd.concat([t1, t2r.iloc[:, :9]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_9-2.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :9]], ignore_index=True)
            case 10:
                match(self.table_size):
                    case "Small":
                        t1 = self.excel_to_python("OACD_tables/2_10-5.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 6]
                        t2r.iloc[:, 1] = t2.iloc[:, 5]
                        t2r.iloc[:, 3] = t2.iloc[:, 1]
                        t2r.iloc[:, 4] = t2.iloc[:, 8]
                        t2r.iloc[:, 5] = t2.iloc[:, 0]
                        t2r.iloc[:, 6] = t2.iloc[:, 9]
                        t2r.iloc[:, 8] = t2.iloc[:, 4]
                        t2r.iloc[:, 9] = t2.iloc[:, 3]
                        self.table = pd.concat([t1, t2r.iloc[:, :10]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/2_10-4.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        t2r = t2
                        t2r.iloc[:, 0] = t2.iloc[:, 4]
                        t2r.iloc[:, 1] = t2.iloc[:, 6]
                        t2r.iloc[:, 2] = t2.iloc[:, 7]
                        t2r.iloc[:, 3] = t2.iloc[:, 1]
                        t2r.iloc[:, 4] = t2.iloc[:, 2]
                        t2r.iloc[:, 5] = t2.iloc[:, 3]
                        t2r.iloc[:, 6] = t2.iloc[:, 9]
                        t2r.iloc[:, 7] = t2.iloc[:, 6]
                        t2r.iloc[:, 8] = t2.iloc[:, 8]
                        t2r.iloc[:, 9] = t2.iloc[:, 0]
                        self.table = pd.concat([t1, t2r.iloc[:, :10]], ignore_index=True)
                    case "Large":
                        t1 = self.excel_to_python("OACD_tables/2_10-3.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa27.xlsx")
                        self.table = pd.concat([t1, t2.iloc[:, :10]], ignore_index=True)
            case _:
                # Return -2 if factor_num is not in valid range
                return -2;
        
        # Change the three-level design into the min, max, and average of each factor
        for factor in range(self.factor_num):
            min_val = self.factor_extrenum.iloc[factor, 0]
            max_val = self.factor_extrenum.iloc[factor, 1]
            avg_val = (min_val + max_val) / 2
            # Replace both int and float -1, 0, 1
            self.table.iloc[:, factor] = self.table.iloc[:, factor].map(
                lambda x: min_val if x == -1 or x == -1.0 else (max_val if x == 1 or x == 1.0 else (avg_val if x == 0 or x == 0.0 else x))
            )
        
        # self.limits = pd.DataFrame({'limits': [None] * self.factor_num})
            
        return 1;
    
    def reduce_levels(self):
        """
        Remove rows from self.table where the number of nonzero factors exceeds self.max_nonzero.
        Only applies if self.max_nonzero is not None and self.table is not None.
        """
        if self.max_nonzero is None or self.table is None:
            return
        # Count nonzero values per row (across all factor columns)
        # Assume all columns are factors
        mask = (self.table != 0).sum(axis=1) <= self.max_nonzero
        self.table = self.table[mask].reset_index(drop=True)
        
    def add_limit(self, factors, limit, index=None):
        """
        Add an overall limiting category that applies to 2+ factors
        
        index allows existing limits to be applied to other factors
        """
        str_index = 0;
        
        # Check if changing limit on factor to None
        if limit is None and factors:
            self.limits.iloc[factors, 0] = None;
            return
        
        while (str(limit) + "_" + str(str_index)) in self.limit_names.keys():
            str_index += 1
        if index != None and index.is_integer():
            str_index = index
        
        # Check if limit was added, but not to any factors
        # If so, will just be added to limit names but not applied to any factors
        if factors:
            # Apply limit to factors
            self.limits.iloc[factors, 0] = str(limit) + "_" + str(str_index);
        self.limit_names.update({str(limit) + "_" + str(str_index) : "Limit: " + str(limit) + " | Entry " + str(str_index)})
    
    def remove_limit(self, limit_name):
        """
        Remove a limit by its name
        """
        if limit_name in self.limit_names.keys():
            del self.limit_names[limit_name]
        
        if limit_name in self.limits['limits'].values:
            self.limits['limits'] = self.limits['limits'].apply(lambda x: None if x == limit_name else x)
            
    def remove_all_limits(self):
        self.limits = pd.DataFrame({'limits': [None] * self.factor_num})
        self.limit_names = {None : None}
            
    def find_limit(self, limit_name):
        """
        Returns the factors that are under the limit
        """
        return self.limits.loc[self.limits['limits'].eq(limit_name)].index.to_list()
    
    def normalize_table(self):
        """
        Normalize the table based on imposed limits
        """        
        # Create the temporary normalized_table
        normalized_table = self.table.copy()
        print(self.limits)
        
        # Iterate through each entry in table
        for run in range(self.table.shape[0]):
            for factor in range(self.table.shape[1]):
                # Find if the row has a limit applied to it
                limit_name = self.limits.iloc[factor, 0]
                if limit_name is not None:
                    # Apply normalization to the values if they go above the limit
                    limit = float(self.limits.iloc[factor, 0].split("_")[0])
                    factors_to_limit = self.find_limit(limit_name)
                    old_total = self.table.iloc[run, factors_to_limit].sum()
                    normalized_value = self.table.iloc[run, factor]
                    if old_total != 0:
                        normalized_value = limit * (self.table.iloc[run, factor] / old_total)
                    normalized_table.iloc[run, factor] = normalized_value
                    
        # Remove duplicate rows
        normalized_table = normalized_table.drop_duplicates(ignore_index=True)
        print("normalized")
        self.table = normalized_table
    

if __name__ == "__main__":
    # oacd = OACD();
    # oacd.set_table_size("Small");
    # oacd.set_factor_num(3);
    # oacd.set_extrenum(0, 0, "min")
    # oacd.set_extrenum(100, 0, "max")
    # oacd.set_extrenum(0, 1, "min")
    # oacd.set_extrenum(100, 1, "max")
    # oacd.set_extrenum(0, 2, "min")
    # oacd.set_extrenum(100, 2, "max")
    # oacd.add_limit([0, 1, 2], 100)
    # oacd.add_limit([0, 1, 2], 100)
    # oacd.add_limit([0, 2], 100, 0)
    # oacd.build_table()
    # print(oacd.table)
    # # oacd.remove_limit("100_0")
    # oacd.normalize_table()
    # print(oacd.table)
    # print(oacd.limit_names)
    pass
    