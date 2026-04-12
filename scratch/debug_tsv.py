import pandas as pd
import io
import warnings
import os

filepath = 'CONFERMESAP.xls'
if not os.path.exists(filepath):
    print(f"File {filepath} not found.")
    exit(1)

try:
    with open(filepath, 'rb') as f:
        raw_content = f.read()
        content = raw_content.decode('utf-16')
        lines = content.split('\n')
        skip = 0
        for i, line in enumerate(lines[:100]):
            tokens = [t.strip() for t in line.split('\t')]
            if any(x in tokens for x in ['Materiale', 'Material', 'MATNR', 'Componente']):
                skip = i
                print(f"Found header at line {i}: {tokens}")
                break
        
        df = pd.read_csv(io.StringIO(content), sep='\t', skiprows=skip)
        print("Columns found:")
        print(df.columns.tolist())
        print("\nFirst 5 rows:")
        print(df.head())
except Exception as e:
    print(f"Error: {e}")
