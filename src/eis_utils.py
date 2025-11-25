import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from pathlib import Path

def load_all_excel_sheets(folder_path):
    folder = Path(folder_path)
    excel_files = list(folder.glob("*.xlsx"))
    
    for file in excel_files:
        try:
            excel_obj = pd.ExcelFile(file)
            print(f"\n File: {file.name}")
            for sheet in excel_obj.sheet_names:
                print(f" Sheet: {sheet}")
        except Exception as e:
            print(f"Error reading {file.name}: {e}")
            
def extract_tables(file_path, sheet_name, table_ranges):
    sheet = pd.read_excel(file_path, sheet_name=sheet_name)
    tables = []
    for i, r in enumerate(table_ranges):
        df = sheet.iloc[r['start_row']:r['end_row'], r['start_col']:r['end_col']]
        df.columns = df.iloc[0]
        df = df[1:].reset_index(drop=True)  # Remove the header row from data
        tables.append(df)
        
        # Print the table
        print(f"\nTable {i+1}:")
        print(df)
        print("-" * 40)
    return tables

def plot_eis_table(table):
    #Choose and plot EIS graphs from a single table: Nyquist Plot, Bode Magnitude Plot, Bode Phase Plot
    print("\nChoose plot type:")
    print("1 = Nyquist Plot (Zre vs -Zim)")
    print("2 = Bode Magnitude |Z| vs Frequency")
    print("3 = Bode Phase vs Frequency")
    choice = input("Enter plot number (1, 2, or 3): ").strip()
        
    #Plot 1: Nyquist
    if choice == "1":
        x = "Zre (ohms)"
        y = "Zim (ohms)"
        title = "Nyquist Plot"

        if x not in table.columns or y not in table.columns:
            print(f"Missing required columns: '{x}' or '{y}'")
            return
        
        table = table.copy()
        table[y] = table[y] * -1

        # Scatter plot
        plt.figure(figsize=(8, 8))  
        sns.scatterplot(data=table, x=x, y=y)

        # Set X and Y axis limits to be equal, based on the min and max of both axes
        min_val = min(table[x].min(), table[y].min())
        max_val = max(table[x].max(), table[y].max(), table[x].max()*1.05)  # 5% larger than Zre max
        plt.xlim(min_val, max_val)
        plt.ylim(min_val, max_val)
        plt.gca().set_aspect('equal', adjustable='box')
        plt.xlabel("Zre (ohms)")
        plt.ylabel("-Zim (ohms)")
        plt.title(title)
        plt.grid(True)
    #Plot 2: Bode Magnitude
    elif choice == "2":
        x = "Frequency (Hz)"
        y = "|Z| (ohms)"
        title = "Bode Magnitude Plot"
        if x not in table.columns or y not in table.columns:
            print(f"Missing required columns: '{x}' or '{y}'")
            return
        # Line plot
        plt.figure(figsize=(10, 6))
        sns.lineplot(data=table, x=x, y=y)
        plt.xscale('log')
        plt.xlabel(x)
        plt.ylabel(y)
        plt.title(title)
        plt.grid(True, which='both', ls='--')
        
    #Plot 3: Bode Phase
    elif choice == "3":
        x = "Frequency (Hz)"
        y = "Phase of Z (deg)"
        title = "Bode Phase Plot"
        if x not in table.columns or y not in table.columns:
            print(f"Missing required columns: '{x}' or '{y}'")
            return
        
        table_plot = table.copy()
        table_plot["-Phase"] = -table_plot[y]
        plt.figure(figsize=(10, 6))
        sns.lineplot(data=table_plot, x=x, y="-Phase")
        plt.xscale('log')
        plt.xlabel(x)
        plt.ylabel("-Phase of Z (deg)")
        plt.title(title)
        plt.grid(True, which='both', ls='--')
       
    else:
        print("Invalid choice.")
        return

    plt.tight_layout()
    plt.show()