#!/usr/bin/env python3
"""
extraer_correos.py
Extrae emails que aparecen tras "Final-Recipient: rfc822;" desde:
  - una carpeta (procesada recursivamente), o
  - un zip (se extrae en un tmp dir).
Genera un XLSX con columna "Email".
Uso:
  py extraer_correos.py <ruta_carpeta_o_zip> [-o salida.xlsx]
Si no se pasa ruta, usa el directorio actual.
"""
import argparse
import zipfile
import re
import tempfile
from pathlib import Path
import sys

try:
    import pandas as pd
except Exception:
    print("ERROR: falta 'pandas'. Instálalo con: py -m pip install pandas openpyxl")
    sys.exit(1)

# --- utilidades ---
def read_text_file(path: Path):
    # intentamos varias codificaciones comunes
    for enc in ("utf-8", "utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=enc)
        except Exception:
            continue
    return None

def find_emails_in_text(text: str):
    # Unfold headers: reemplaza CRLF + SP/HT por espacio (maneja líneas partidas)
    unfolded = re.sub(r'\r?\n[ \t]+', ' ', text)
    # Regex para capturar la dirección justo después del marcador
    pattern = re.compile(
        r'Final-Recipient:\s*rfc822;\s*<?\s*([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})\s*>?',
        re.IGNORECASE
    )
    return pattern.findall(unfolded)

def process_folder(folder: Path):
    emails = set()
    for p in folder.rglob('*'):
        if not p.is_file():
            continue
        txt = read_text_file(p)
        if not txt:
            continue
        found = find_emails_in_text(txt)
        for e in found:
            emails.add(e.strip().lower())   # deduplicamos ignorando mayúsc/minúsc
    return emails

# --- main ---
def main():
    parser = argparse.ArgumentParser(description="Extrae emails tras 'Final-Recipient: rfc822;' (carpeta o zip) y guarda XLSX.")
    parser.add_argument("path", nargs="?", default=".", help="Ruta a carpeta o ZIP (por defecto: directorio actual).")
    parser.add_argument("-o", "--output", default="correos_rebotados.xlsx", help="Archivo de salida (.xlsx o .csv).")
    args = parser.parse_args()

    input_path = Path(args.path).expanduser()
    if not input_path.exists():
        print(f"ERROR: no existe la ruta: {input_path}")
        sys.exit(1)

    # Si es ZIP, extraer en tmpdir; si es carpeta, procesar directamente
    if input_path.is_file() and zipfile.is_zipfile(input_path):
        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)
            with zipfile.ZipFile(input_path, 'r') as z:
                z.extractall(td_path)
            print(f"Procesando ZIP (temporal): {input_path} ...")
            emails = process_folder(td_path)
    elif input_path.is_dir():
        print(f"Procesando carpeta: {input_path} ...")
        emails = process_folder(input_path)
    else:
        print("ERROR: la ruta no es ni carpeta ni ZIP válido.")
        sys.exit(1)

    emails = sorted(emails)
    if not emails:
        print("No se encontraron direcciones tras 'Final-Recipient: rfc822;'.")
        sys.exit(0)

    out_path = Path(args.output)
    if out_path.suffix.lower() in (".xlsx", ".xls"):
        df = pd.DataFrame(emails, columns=["Email"])
        df.to_excel(out_path, index=False)
    else:
        # .csv u otros
        df = pd.DataFrame(emails, columns=["Email"])
        df.to_csv(out_path, index=False)

    print(f"Hecho: {len(emails)} dirección(es) guardadas en: {out_path.resolve()}")

if __name__ == "__main__":
    main()
