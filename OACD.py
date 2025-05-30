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
        
    def set_table_size(self, table_size):
        self.table_size = table_size
        
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
                        t2 = t2.iloc[:, [1, 2, 3, 5]]
                        self.table = pd.concat([t1, t2.iloc[:, :5]], ignore_index=True)
                    case "Medium":
                        t1 = self.excel_to_python("OACD_tables/pb12.xlsx")
                        t2 = self.excel_to_python("OACD_tables/oa18.xlsx")
                        t2 = t2.iloc[:, [1, 4, 2, 3, 5]]
                        self.table = pd.concat([t1, t2.iloc[:, :5]], ignore_index=True)
                    case "Large":
                        pass;
            case 6:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 7:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 8:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 9:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 10:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case _:
                # Return -2 if factor_num is not in valid range
                return -2;
        
        return 1;
                
    
if __name__ == "__main__":
    oacd = OACD();
    oacd.set_factor_num(5);
    oacd.set_table_size("Small");
    built = oacd.build_table();

    print(oacd.table)
    # array = df.values.tolist()
    # print(array[:][3])
    