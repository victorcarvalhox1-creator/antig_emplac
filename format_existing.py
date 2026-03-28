import os
import glob
import pandas as pd

def apply_formatting(file_path):
    print(f"Formatting {file_path} ...")
    try:
        df = pd.read_excel(file_path, header=1)
        if len(df.columns) >= 10:
            new_cols = [
                df.columns[0], # A
                df.columns[5], # F
                df.columns[7], # H
                df.columns[1], # B
                df.columns[9], # J
                df.columns[8], # I
                df.columns[3], # D
                df.columns[6]  # G
            ]
            df_formatted = df[new_cols].copy()
            df_formatted.columns = ['Data', 'Chassis', 'Fabricante', 'Modelo', 'Municipio', 'UF', 'CNPJ', 'Placa']
            df_formatted.to_excel(file_path, index=False)
            print(f"Successfully formatted: {file_path}")
        else:
            print(f"Skipped {file_path}: Doesn't match expected original structure.")
    except Exception as e:
         print(f"Failed {file_path}: {e}")

if __name__ == "__main__":
    files = glob.glob("downloads/*.xlsx")
    for f in files:
        apply_formatting(f)
