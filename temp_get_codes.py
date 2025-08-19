import pandas as pd

try:
    df = pd.read_excel("Prueba10_con_creditos.xlsx")
    # Drop duplicates to keep only one entry per course
    df_unique = df.drop_duplicates(subset=['asignatura'])
    # Create the dictionary mapping
    code_map = pd.Series(df_unique.codigo_asignatura.values, index=df_unique.asignatura).to_dict()
    print(code_map)
except FileNotFoundError:
    print("Error: Archivo 'Prueba10_con_creditos.xlsx' no encontrado.")
except Exception as e:
    print(f"Error: {e}")