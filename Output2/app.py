# -*- coding: utf-8 -*-
"""
App GUI: Copiar Output2 y anteponer (A–C) las 3 primeras columnas de Output1
en las filas que coincidan por DOBLE llave extraída de la última columna de Output1: 'PRIMERA/SEGUNDA'

Llaves de unión:
  - PRIMERA (antes de '/') ↔ Columna H de Output2 (Wholesale Invoice Number)
      * normalizar a SOLO dígitos y sin ceros a la izquierda (maneja mezcla de formatos)
  - SEGUNDA (después de '/') ↔ Primera columna de Output2
      * quitar ceros a la izquierda

Salida:
  - Base: TODO Output2 (todas sus columnas y filas, sin perder nada)
  - Se ANTEPONEN (columnas A–C) las PRIMERAS 3 columnas de Output1 (por posición) para las filas que coinciden
    - Si algún nombre de esas 3 columnas ya existe en Output2, se agregan con sufijo "_O1" (no se sobreescribe nada)
  - Guardar en carpeta 'Resultado' junto a Output2, con el MISMO NOMBRE de Output2
    * Si Output2 es .xls, se guardará como .xlsx por compatibilidad

Registro (HKCU):
  - Guarda temporalmente rutas seleccionadas y se BORRA inmediatamente al finalizar el proceso.

Autor: M365 Copilot
"""

import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# ---------------------------------
# Manejo de Registro (Windows HKCU)
# ---------------------------------
try:
    import winreg
except ImportError:
    winreg = None  # En sistemas no-Windows, el registro no está disponible

REG_PATH = r"Software\AA\OutputMatcher"

def reg_write_temp_paths(path1: str, path2: str):
    """Escribe rutas temporales en HKCU. Si no hay winreg, no hace nada."""
    if winreg is None:
        return
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        winreg.SetValueEx(key, "LastOutput1", 0, winreg.REG_SZ, path1 or "")
        winreg.SetValueEx(key, "LastOutput2", 0, winreg.REG_SZ, path2 or "")
        winreg.CloseKey(key)
    except Exception:
        pass  # no romper si falla el registro

def reg_read_temp_paths():
    """Lee rutas temporales desde HKCU. Retorna (output1, output2)."""
    if winreg is None:
        return ("", "")
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH)
        o1, _ = winreg.QueryValueEx(key, "LastOutput1")
        o2, _ = winreg.QueryValueEx(key, "LastOutput2")
        winreg.CloseKey(key)
        return (o1 or "", o2 or "")
    except Exception:
        return ("", "")

def reg_delete_key_recursive(root, sub_key):
    """Elimina una clave del Registro de forma recursiva (por si tiene subclaves)."""
    try:
        with winreg.OpenKey(root, sub_key, 0, winreg.KEY_READ | winreg.KEY_WRITE) as hkey:
            i = 0
            while True:
                try:
                    child = winreg.EnumKey(hkey, i)
                    reg_delete_key_recursive(root, sub_key + "\\" + child)
                    i += 1
                except OSError:
                    break
        winreg.DeleteKey(root, sub_key)
    except FileNotFoundError:
        pass
    except Exception:
        pass  # evitar que un error aquí rompa el proceso

def reg_delete_temp():
    """Borra la clave temporal HKCU\\Software\\AA\\OutputMatcher inmediatamente tras el proceso."""
    if winreg is None:
        return
    reg_delete_key_recursive(winreg.HKEY_CURRENT_USER, REG_PATH)

# -------------------------
# Utilidades de limpieza
# -------------------------
def to_str_series(s: pd.Series) -> pd.Series:
    """Convierte a str, normaliza NBSP, colapsa espacios múltiples y recorta."""
    s = s.astype(str)
    s = s.str.replace("\u00a0", " ", regex=False)  # NBSP -> espacio normal
    s = s.str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.replace({"nan": None, "NaN": None, "None": None})
    return s

def normalize_nozeros_digits(s: pd.Series) -> pd.Series:
    """
    Normaliza a SOLO DÍGITOS sin ceros a la izquierda:
      - Extrae únicamente dígitos [0-9]
      - Elimina ceros a la izquierda
      - Si queda vacío, retorna "0" (estandariza cero)
    Útil para comparar números de factura mezclados (texto/número).
    """
    s = to_str_series(s)
    s = s.apply(lambda x: re.sub(r"\D+", "", x) if isinstance(x, str) else x)
    def _strip_zeros(val):
        if val is None:
            return None
        v = val.lstrip("0")
        return v if v != "" else "0"
    return s.apply(_strip_zeros)

def normalize_strip_leading_zeros(s: pd.Series) -> pd.Series:
    """
    Remueve ceros a la izquierda preservando el resto de caracteres.
    Si queda cadena vacía, devuelve None.
    """
    s = to_str_series(s)
    def _lstrip0(val):
        if val is None:
            return None
        v = val.lstrip("0")
        return v if v != "" else None
    return s.apply(_lstrip0)

def split_primera_segunda(serie_ultima_col: pd.Series):
    """
    Divide la última columna de Output1 en PRIMERA y SEGUNDA por el primer '/'.
    Si no hay '/', PRIMERA = valor completo, SEGUNDA = None.
    """
    s = to_str_series(serie_ultima_col)
    partes = s.str.split("/", n=1, expand=True)
    if partes.shape[1] == 1:
        primera = partes[0]
        segunda = pd.Series([None]*len(primera), index=primera.index)
    else:
        primera = partes[0]
        segunda = partes[1]
    return primera, segunda

def detectar_columna_H(df: pd.DataFrame):
    """
    Detecta la columna 'Wholesale Invoice Number' de Output2:
      1) Busca nombre que contenga 'wholesale' y 'invoice' (case-insensitive)
      2) Si no encuentra, usa la 8va columna (índice 7)
    """
    for c in df.columns:
        c_low = str(c).lower()
        if "wholesale" in c_low and "invoice" in c_low:
            return c
    if df.shape[1] >= 8:
        return df.columns[7]
    raise ValueError("No se pudo detectar la columna H (Wholesale Invoice Number) en Output2.")

# -------------------------
# Lógica principal
# -------------------------
def procesar_archivos(ruta_output1: tk.StringVar, ruta_output2: tk.StringVar):
    """Procesa los archivos usando las rutas de los StringVar. Borra HKCU al terminar."""
    reg_write_temp_paths(ruta_output1.get(), ruta_output2.get())

    try:
        if not ruta_output1.get() or not ruta_output2.get():
            raise ValueError("Debes seleccionar los dos archivos (Output1 y Output2).")

        # Leer como texto para evitar problemas de tipos y preservar ceros a la izquierda
        df1 = pd.read_excel(ruta_output1.get(), dtype=str)
        df2 = pd.read_excel(ruta_output2.get(), dtype=str)

        # Validaciones mínimas
        if df1.shape[1] < 4:
            raise ValueError("Output1 debe tener al menos 4 columnas (3 primeras para la salida y una última con 'PRIMERA/SEGUNDA').")
        if df2.shape[1] < 2:
            raise ValueError("Output2 debe tener al menos 2 columnas.")

        # ---- OUTPUT1: llaves y columnas de salida ----
        cols_out1_first3 = df1.columns[:3].tolist()  # Usar por posición (p.ej., PERIODO, PEDIDO, NOTA)
        ultima_columna = df1.columns[-1]
        primera_raw, segunda_raw = split_primera_segunda(df1[ultima_columna])

        # Llave K1 (PRIMERA contra H) -> solo dígitos sin ceros a la izquierda
        k1_output1 = normalize_nozeros_digits(primera_raw)
        # Llave K2 (SEGUNDA contra col1) -> quitar ceros a la izquierda
        k2_output1 = normalize_strip_leading_zeros(segunda_raw)

        df1_loc = df1.copy()
        df1_loc["K1_PRIMERA_NOZEROS"] = k1_output1
        df1_loc["K2_SEGUNDA_NOZEROS"] = k2_output1

        # Filas con ambas llaves presentes
        df1_valid = df1_loc[df1_loc["K1_PRIMERA_NOZEROS"].notna() & df1_loc["K2_SEGUNDA_NOZEROS"].notna()].copy()

        # Tabla de aporte desde Output1 (una por combinación de llaves)
        df1_keys = df1_valid[["K1_PRIMERA_NOZEROS", "K2_SEGUNDA_NOZEROS"] + cols_out1_first3].copy()
        df1_keys = df1_keys.drop_duplicates(subset=["K1_PRIMERA_NOZEROS", "K2_SEGUNDA_NOZEROS"], keep="first")

        # ---- OUTPUT2: llaves ----
        col1_out2 = df2.columns[0]         # primera columna
        colH_out2 = detectar_columna_H(df2)  # columna H
        k1_out2 = normalize_nozeros_digits(df2[colH_out2])
        k2_out2 = normalize_strip_leading_zeros(df2[col1_out2])

        df2_loc = df2.copy()
        df2_loc["K1_PRIMERA_NOZEROS"] = k1_out2
        df2_loc["K2_SEGUNDA_NOZEROS"] = k2_out2

        # ---- Resolver colisiones de nombres ----
        o1_cols_renamed = []
        rename_map = {}
        for c in cols_out1_first3:
            if c in df2_loc.columns:
                new_c = f"{c}_O1"
                rename_map[c] = new_c
                o1_cols_renamed.append(new_c)
            else:
                o1_cols_renamed.append(c)
        if rename_map:
            df1_keys = df1_keys.rename(columns=rename_map)

        # ---- Merge LEFT: mantener TODO Output2 y anexar columnas de Output1 cuando coincida ----
        df_merged = pd.merge(
            df2_loc,  # base
            df1_keys[["K1_PRIMERA_NOZEROS", "K2_SEGUNDA_NOZEROS"] + o1_cols_renamed],
            on=["K1_PRIMERA_NOZEROS", "K2_SEGUNDA_NOZEROS"],
            how="left",
            validate="many_to_one"  # cada combinación de llaves en df1_keys debe ser única
        )

        # Quitar columnas de llaves auxiliares del resultado final
        df_merged = df_merged.drop(columns=["K1_PRIMERA_NOZEROS", "K2_SEGUNDA_NOZEROS"])

        # ---- Reordenar columnas: anteponer A–C con las 3 columnas aportadas de Output1 ----
        # Mantener orden original del Output2 para el resto
        cols_current = list(df_merged.columns)
        # Primero, asegurar lista de columnas a anteponer en el orden correcto
        prepend_cols = [c for c in o1_cols_renamed if c in cols_current]
        # Luego, el resto de columnas (Output2) en su orden original
        rest_cols = [c for c in cols_current if c not in prepend_cols]
        final_cols = prepend_cols + rest_cols
        df_final = df_merged[final_cols]

        # ---- Guardado: MISMO NOMBRE que Output2 dentro de /Resultado ----
        carpeta_salida = os.path.join(os.path.dirname(ruta_output2.get()), "Resultado")
        os.makedirs(carpeta_salida, exist_ok=True)

        nombre_output2 = os.path.basename(ruta_output2.get())
        base, ext = os.path.splitext(nombre_output2)
        ext = ext.lower()

        changed_ext = False
        if ext == ".xlsx":
            ruta_salida = os.path.join(carpeta_salida, nombre_output2)
        else:
            # Por compatibilidad, guardar como .xlsx
            ruta_salida = os.path.join(carpeta_salida, base + ".xlsx")
            changed_ext = True

        df_final.to_excel(ruta_salida, index=False)

        # Mensaje
        info_cols = f"Columnas totales: {len(df_final.columns)}  (primeras 3 son de Output1{' _renombradas con _O1' if rename_map else ''})"
        if changed_ext:
            messagebox.showinfo(
                "Éxito",
                (
                    f"Archivo generado (se cambió extensión a .xlsx por compatibilidad):\n"
                    f"{ruta_salida}\n\nFilas: {len(df_final)} | {info_cols}"
                )
            )
        else:
            messagebox.showinfo(
                "Éxito",
                (
                    f"Archivo generado:\n{ruta_salida}\n\nFilas: {len(df_final)} | {info_cols}"
                )
            )

    except Exception as e:
        messagebox.showerror("Error", str(e))

    finally:
        # Borrar inmediatamente el registro HKCU al terminar (éxito o error)
        reg_delete_temp()

# -------------------------
# GUI (Tkinter)
# -------------------------
def construir_gui():
    ventana = tk.Tk()
    ventana.title("Copiar Output2 y ANTEPONER 3 columnas de Output1 (doble llave)")
    ventana.resizable(False, False)

    ruta_output1 = tk.StringVar(master=ventana)
    ruta_output2 = tk.StringVar(master=ventana)

    frm = tk.Frame(ventana, padx=10, pady=10)
    frm.pack(fill="both", expand=True)

    lbl_info = tk.Label(
        frm,
        text=(
            "Flujo:\n"
            "• Output1 (última col): 'PRIMERA/SEGUNDA'\n"
            "   - PRIMERA (10 dígitos → solo dígitos sin ceros) ↔ Columna H (Wholesale Invoice #) de Output2\n"
            "   - SEGUNDA (sin ceros a la izquierda) ↔ Primera columna de Output2\n"
            "• Se conserva TODO Output2; se ANTEPONEN (A–C) las primeras 3 columnas de Output1 cuando coincida\n"
            "• Resultado: misma ruta de Output2 dentro de carpeta 'Resultado', con el MISMO nombre de archivo\n"
            "• El registro HKCU es temporal y se BORRA al finalizar"
        ),
        justify="left"
    )
    lbl_info.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

    def seleccionar_output1():
        archivo = filedialog.askopenfilename(
            title="Seleccionar Output1",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if archivo:
            ruta_output1.set(archivo)
            reg_write_temp_paths(ruta_output1.get(), ruta_output2.get())

    def seleccionar_output2():
        archivo = filedialog.askopenfilename(
            title="Seleccionar Output2",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if archivo:
            ruta_output2.set(archivo)
            reg_write_temp_paths(ruta_output1.get(), ruta_output2.get())

    tk.Label(frm, text="Output1:").grid(row=1, column=0, sticky="e", padx=(0, 6))
    tk.Entry(frm, textvariable=ruta_output1, width=60).grid(row=1, column=1, sticky="we")
    tk.Button(frm, text="Buscar", command=seleccionar_output1).grid(row=1, column=2, padx=(6, 0))

    tk.Label(frm, text="Output2:").grid(row=2, column=0, sticky="e", padx=(0, 6), pady=(6, 0))
    tk.Entry(frm, textvariable=ruta_output2, width=60).grid(row=2, column=1, sticky="we", pady=(6, 0))
    tk.Button(frm, text="Buscar", command=seleccionar_output2).grid(row=2, column=2, padx=(6, 0), pady=(6, 0))

    tk.Button(
        frm,
        text="Procesar",
        command=lambda: procesar_archivos(ruta_output1, ruta_output2),
        bg="#c6f6c6"
    ).grid(row=3, column=0, columnspan=3, pady=(12, 0), sticky="we")

    # Precargar rutas si existieran (se borrarán al finalizar)
    o1, o2 = reg_read_temp_paths()
    if o1:
        ruta_output1.set(o1)
    if o2:
        ruta_output2.set(o2)

    ventana.mainloop()

# -------------------------
# Main
# -------------------------
if __name__ == "__main__":
    construir_gui()
