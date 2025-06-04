import pandas as pd
import tkinter as tk
import oapackage
import numpy as np

class OACD:
    def __init__(self):
        
        # Variables for OACD Tables
        self.factor_num = None # 2, 3, 4, 5, 6, 7, 8, 9, 10
        self.table_size = None # Small, Medium, Large
        self.table = None # Pandas DataFrame object
        self.factor_extrenum = None # First column is minimum, second column is maximum; each row is a factor
        
    def __str__(self):
        string = ""
        for row in self.table:
            for cell in row:
                string += cell + " "
            string += "\n"
        return string;
        
    def set_factor_num(self, factor_num):
        self.factor_num = factor_num
        
        # When factor number is changed, the extrenum table is automatically changed too
        self.factor_extrenum = pd.DataFrame(np.ones((self.factor_num, 2)))
        self.factor_extrenum.iloc[:, 1] = -1; 
        
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
        
    def excel_to_python(self, filepath):
        df = pd.read_excel(filepath, header=None)
        return df    
         
    def build_table(self):
        
        # Returns -1 if there table isn't built / error in factor_num
        if(self.table_size != "Small" and self.table_size != "Medium" and self.table_size != "Large"):
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
            self.table.iloc[:, factor].replace(1, self.factor_extrenum.iloc[factor, 1])
            self.table.iloc[:, factor].replace(-1, self.factor_extrenum.iloc[factor, 0])
            self.table.iloc[:, factor].replace(0, (self.factor_extrenum.iloc[factor, 1] + self.factor_extrenum.iloc[factor, 0])/2)
        
        return 1;
    
if __name__ == "__main__":
    oacd = OACD();
    oacd.set_factor_num(10);
    oacd.set_table_size("Large");
    built = oacd.build_table();

    print(oacd.table)
    