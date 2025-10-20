import pandas as pd
import json
import os
from pathlib import Path

def excel_to_json_dynamic(path):
    # Carica il file Excel
    df = pd.read_excel(path)
    df = df.dropna(axis=1, how="all")  # rimuove colonne completamente vuote
    df = df.ffill()     # riempie le celle vuote con il valore precedente
    df.columns = [str(c).strip() for c in df.columns]  # pulizia nomi colonne

    # Identifica le colonne chiave
    step_cols = [c for c in df.columns if "step" in c.lower() or "result" in c.lower()]
    id_col = next((c for c in df.columns if c.lower() == "id"), None)

    if not id_col:
        raise ValueError("❌ Nessuna colonna 'ID' trovata nel file Excel!")

    # Tutte le altre colonne tranne gli step
    meta_cols = [c for c in df.columns if c not in step_cols]

    # Dizionario finale
    tests = {}

    for _, row in df.iterrows():
        test_id = str(row[id_col]).strip()
        if test_id not in tests:
            # Crea dinamicamente il dizionario dei metadati
            meta_data = {col: row.get(col, "") for col in meta_cols if col != id_col}
            meta_data["Steps"] = []
            tests[test_id] = meta_data

        # Costruisce lo step
        step_data = {col: row.get(col, "") for col in step_cols}
        tests[test_id]["Steps"].append(step_data)

    # Esporta in JSON
    # Directory di output (stessa cartella dello script)
    output_dir = os.path.join(os.path.dirname(__file__), "outputs")
    output_path = os.path.join(output_dir, "tests_output.json")

    # Scrive il file JSON
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(tests, f, indent=2, ensure_ascii=False)

    print(f"✅ JSON creato in: {output_path}")
    return tests


# Esempio di uso:
excel_to_json_dynamic(Path(__file__).parent / "input" / "tests_cases.xlsx")
