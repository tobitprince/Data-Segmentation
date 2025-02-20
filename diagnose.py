import pandas as pd

def diagnose_excel():
    # Load Irene's campaign list and check each sheet
    print("Analyzing Irene Email campaign list.xlsx...")
    excel_file = pd.ExcelFile("Irene Email campaign list.xlsx")
    
    print("\nSheets found:", excel_file.sheet_names)
    
    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        print(f"\nSheet: {sheet_name}")
        print("Columns:", df.columns.tolist())
        print(f"Number of rows: {len(df)}")
        # Check if 'email' column exists and show sample
        if 'email' in df.columns:
            print("Sample emails:", df['email'].head(3).tolist())
        elif 'Email' in df.columns:
            print("Sample emails:", df['Email'].head(3).tolist())

if __name__ == "__main__":
    diagnose_excel()