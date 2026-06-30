import pandas as pd
import re

def limpiar_a_numeros(texto):
    """Extrae solo dígitos del texto."""
    if not isinstance(texto, str):
        texto = str(texto)
    numeros = re.findall(r'\d+', texto)
    return numeros[0] if numeros else ""

def procesar_archivo1(df, periodo_val):
    # 1) Renombrar primera columna a NOTA (columna original 0)
    nombres = df.columns.tolist()
    if len(nombres) == 0:
        raise ValueError("El archivo 1 no tiene columnas.")
    nombres[0] = "NOTA"
    df.columns = nombres

    # 2) Insertar columna PEDIDO vacía a la derecha de NOTA
    df.insert(1, "PEDIDO", "")

    # 3) Limpiar NOTA: dejar solo números
    df["NOTA"] = df["NOTA"].astype(str).apply(limpiar_a_numeros)

    # 4) Insertar PERIODO al inicio (columna 0) con el valor del UI
    df.insert(0, "PERIODO", periodo_val if periodo_val is not None else "")

    # Resultado: columnas → PERIODO, NOTA, PEDIDO, ...
    return df

def procesar_archivo2(df):
    # Renombrar columna 1 → PEDIDO y columna 3 → NOTA
    nombres = df.columns.tolist()
    if len(nombres) < 3:
        raise ValueError("El archivo 2 debe tener al menos 3 columnas.")
    nombres[0] = "PEDIDO"
    nombres[2] = "NOTA"
    df.columns = nombres
    return df

def realizar_lookup(df1, df2):
    # Higiene de nombres
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    # Convertir NOTA a texto limpio
    df1["NOTA"] = df1["NOTA"].astype(str).str.strip()
    df2["NOTA"] = df2["NOTA"].astype(str).str.strip()

    # Merge estilo XLOOKUP, evitando conflicto de nombres
    df_merge = df1.merge(
        df2[["NOTA", "PEDIDO"]],
        on="NOTA",
        how="left",
        suffixes=("", "_src")
    )

    # Llenar PEDIDO en df1 con los valores del archivo 2
    df1["PEDIDO"] = df_merge["PEDIDO_src"]
    return df1

def generar_output(file1, file2, output, periodo):
    """
    Función principal para usar desde el UI o CLI.
    - file1: ruta archivo 1 (Output)
    - file2: ruta archivo 2 (F)
    - output: nombre/ruta del archivo de salida
    - periodo: valor a insertar en columna PERIODO
    """
    # Leer archivos
    df1 = pd.read_excel(file1, engine="openpyxl")
    df2 = pd.read_excel(file2, engine="openpyxl")

    # Procesar
    df1 = procesar_archivo1(df1, periodo_val=periodo)
    df2 = procesar_archivo2(df2)

    # Lookup
    df_final = realizar_lookup(df1, df2)

    # Guardar
    df_final.to_excel(output, index=False)
    return output

if __name__ == "__main__":
    # Modo CLI opcional: python master_script.py archivo1 archivo2 output periodo
    import sys
    if len(sys.argv) < 5:
        print("Uso: python master_script.py archivo1.xlsx archivo2.xlsx output.xlsx PERIODO")
    else:
        _, a1, a2, out, per = sys.argv
        generar_output(a1, a2, out, per)
        print(f"Archivo final guardado como: {out}")