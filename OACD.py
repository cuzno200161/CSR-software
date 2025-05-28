import pandas as pd
import tkinter as tk
import oapackage

class OACD:
    def __init__(self):
        
        # Variables for OACD Tables
        self.factor_num = None; # 2, 3, 4, 5, 6, 7, 8, 9, 10
        self.table_size = None; # Small, Medium, Large
        self.table = None;
        self.factor_extrenum = None; # First column is minimum, second column is maximum
        
    def __str__(self):
        string = "";
        for row in self.table:
            for cell in row:
                string += cell + " ";
            string += "\n";
        return string;
        
    def set_factor_num(self, factor_num):
        self.factor_num = factor_num;
        
    def set_table_size(self, table_size):
        self.table_size = table_size;
        
    def excel_to_python(self, filepath):
        pass;
            
        
    def build_table(self):
        # Returns -1 if there table isn't built / error in factor_num
        if(self.table_size != "Small" or self.table_size != "Medium" or self.table_size != "Large"):
            return -1;
        
        match(self.factor_num):
            case 2:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 3:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 4:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
                    case "Large":
                        pass;
            case 5:
                match(self.table_size):
                    case "Small":
                        pass;
                    case "Medium":
                        pass;
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
    oacd.set_factor_num(2);
    oacd.set_table_size(10);
    oacd.build_table();
    
    factors = 3;
    strength = 2;
    levels = 3;
    runs = 9;
    
    array=oapackage.arraydata_t(levels, runs, strength, factors)
    ll2 = [array.create_root()]
    ll2[0].showarraycompact()   
    
    