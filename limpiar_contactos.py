#!/usr/bin/env python3
import argparse
from pathlib import Path
import sys

try:
    import pandas as pd
except Exception:
    print("ERROR: falta 'pandas'. Instálalo con: py -m pip install pandas openpyxl")
    sys.exit(1)

def read_table(path):
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"No existe: {p}")
    if p.suffix.lower() in (".xls", ".xlsx"):
        return pd.read_excel(p)
    else:
        # Intentar CSV con separador automático
        return pd.read_csv(p)

def detect_email_column(df):
    candidates = ["email", "e-mail", "correo", "mail", "address", "dirección", "direccion"]
    for col in df.columns:
        name = str(col).strip().lower()
        if name in candidates or "email" in name or "correo" in name or "e-mail" in name:
            return col
    return None

def normalize_series(s):
    return s.astype(str).str.strip().str.lower().replace({"nan": ""})

def main():
    parser = argparse.ArgumentParser(description="Quita de tus contactos los emails que están en la lista de rebotes.")
    parser.add_argument("contacts", help="Archivo de contactos (.xlsx o .csv)")
    parser.add_argument("rebotes", nargs="?", default="correos_rebotados.xlsx", help="Archivo de rebotes (por defecto: correos_rebotados.xlsx)")
    parser.add_argument("-c", "--email-column", help="Nombre exacto de la columna de emails en el archivo de contactos (opcional)")
    parser.add_argument("-o", "--output", help="Archivo de salida (por defecto: mismo nombre que contactos + _limpios)")
    args = parser.parse_args()

    # Cargar archivos
    try:
        df_contacts = read_table(args.contacts)
    except Exception as e:
        print("ERROR al leer contactos:", e)
        sys.exit(1)
    try:
        df_rebotes = read_table(args.rebotes)
    except Exception as e:
        print("ERROR al leer rebotes:", e)
        sys.exit(1)

    # Detectar columna de email en contactos
    email_col = args.email_column
    if email_col:
        # aceptar nombre insensible a mayúsculas
        matches = [c for c in df_contacts.columns if str(c).strip().lower() == email_col.strip().lower()]
        if matches:
            email_col = matches[0]
        elif email_col not in df_contacts.columns:
            print("ERROR: la columna indicada no existe en el archivo de contactos.")
            print("Columnas disponibles:", list(df_contacts.columns))
            sys.exit(1)
    else:
        auto = detect_email_column(df_contacts)
        if auto is None:
            print("No pude detectar automáticamente la columna de emails.")
            print("Columnas encontradas en el archivo de contactos:")
            for c in df_contacts.columns:
                print(" -", c)
            print("Vuelve a ejecutar indicando --email-column \"NOMBRE_DE_COLUMNA\"")
            sys.exit(1)
        email_col = auto

    # Preparar sets normalizados
    df_contacts[email_col] = df_contacts[email_col].fillna("").astype(str)
    contactos_norm = normalize_series(df_contacts[email_col])
    df_contacts = df_contacts[contactos_norm != ""]  # eliminar filas sin email
    before = len(df_contacts)

    # Rebotes: detectar columna de email o usar la primera
    reb_col = detect_email_column(df_rebotes) or df_rebotes.columns[0]
    df_rebotes[reb_col] = df_rebotes[reb_col].fillna("").astype(str)
    rebotes_norm = set(normalize_series(df_rebotes[reb_col]).tolist())

    # Filtrar
    mask_keep = ~normalize_series(df_contacts[email_col]).isin(rebotes_norm)
    df_clean = df_contacts[mask_keep].copy()
    after = len(df_clean)
    removed = before - after

    # Guardar en mismo formato que el archivo de contactos (o en salida indicada)
    in_path = Path(args.contacts)
    if args.output:
        out_path = Path(args.output)
    else:
        out_path = in_path.parent / f"{in_path.stem}_limpios{in_path.suffix}"

    if out_path.suffix.lower() in (".xls", ".xlsx"):
        df_clean.to_excel(out_path, index=False)
    else:
        df_clean.to_csv(out_path, index=False)

    print(f"Hecho. Eliminados: {removed}. Guardado: {out_path.resolve()}")

if __name__ == "__main__":
    main()
