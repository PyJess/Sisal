import json
import os
import pandas as pd
from collections import defaultdict
import subprocess
import re
from pathlib import Path

def load_file(filepath:str):
    with open(filepath, encoding="utf-8") as f:
        return f.read()


def load_json(filepath:str):
    with open(filepath, encoding="utf-8") as f:
        return json.load(f)


def save_json(data, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def get_user_path(user_id: str, subfolder: str = "") -> str:
    """Get the user-specific path."""
    base_dir = os.path.join(os.path.dirname(__file__),"outputs", user_id, subfolder)
    os.makedirs(base_dir, exist_ok=True)
    return base_dir


def excel_to_json(path):
    df = pd.read_excel(path)
    df = df.dropna(axis=1, how="all")  
    df = df.ffill()     
    df.columns = [str(c).strip() for c in df.columns]  

    step_cols = [c for c in df.columns if "step" in c.lower() or "result" in c.lower()]
    id_col = next((c for c in df.columns if c.lower() == "id"), None)

    if not id_col:
        raise ValueError("❌ Nessuna colonna 'ID' trovata nel file Excel!")

    # Tutte le altre colonne tranne gli step
    meta_cols = [c for c in df.columns if c not in step_cols]

    tests = {}

    for _, row in df.iterrows():
        test_id = str(row[id_col]).strip()
        if test_id not in tests:
            meta_data = {col: row.get(col, "") for col in meta_cols if col != id_col}
            meta_data["Steps"] = []
            tests[test_id] = meta_data

        step_data = {col: row.get(col, "") for col in step_cols}
        tests[test_id]["Steps"].append(step_data)

    output_dir = os.path.join(os.path.dirname(__file__),"..", "input")
    output_path = os.path.join(output_dir, "tests_output.json")

    # Scrive il file JSON
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(tests, f, indent=2, ensure_ascii=False)

    print(f"✅ JSON creato in: {output_path}")
    return tests



def group_by_funzionalita(data):
    grouped = defaultdict(dict)

    for key, value in data.items():
        funzionalita = value.get("Funzionalità", "Unknown")
        grouped[funzionalita][key] = value

    # Print nicely
    #print(json.dumps(grouped, indent=2, ensure_ascii=False))
    return (json.dumps(grouped, indent=2, ensure_ascii=False))


PANDOC_EXE = "pandoc" 
def process_docx(docx_path, output_base):
    """
    Process a DOCX file using Pandoc and split it into sections based on Markdown headers (#, ##, etc.).
    """
    
    txt_output_path = os.path.join(output_base, Path(docx_path).stem + ".txt")
    
    docx_path = os.path.normpath(docx_path)
    txt_output_path = os.path.normpath(txt_output_path)
    os.makedirs(output_base, exist_ok=True)
    
    # Convert in md
    command = [
        PANDOC_EXE,
        "-s", docx_path,
        "--columns=120",
        "-t", "markdown",
        "-o", txt_output_path
    ]
    
    try:
        subprocess.run(command, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=30)
    except Exception as e:
        print(f"[ERROR] Pandoc conversion failed: {e}")
        return [], os.path.basename(docx_path), []
    
    with open(txt_output_path, "r", encoding="utf-8") as f:
        text_lines = f.read().splitlines()
    
    headers = []
    heading_list = []
    
    for index, line in enumerate(text_lines):
        if line.startswith("#"):
            level = line.count("#")
            clean_name = line.replace("#", "").strip()
            headers.append([clean_name, index, level])
            heading_list.append([clean_name, level])
    
    headers.insert(0, ["== first line ==", 0, 0])
    headers.append(["== last line ==", len(text_lines), 0])
    heading_list.insert(0, ["== first line ==", 0])
    heading_list.append(["== last line ==", 0])
    
    chunks = []
    for i in range(len(headers) - 1):
        start_idx = headers[i][1]
        end_idx = headers[i + 1][1]
        section_lines = text_lines[start_idx:end_idx]
        chunk_text = "\n".join(section_lines).strip()
        
        header_cleaned = re.sub(r"\s*\{.*?\}", "", headers[i][0])
        header_cleaned = header_cleaned.replace("--", "–").strip(" *[]\n")
        chunk_text = header_cleaned + "\n" + chunk_text
        
        chunks.append(chunk_text)
    
    return chunks



def fill_excel_file(test_cases: dict, output_path: str = None):
    """
    Salva i test case in un file Excel, mantenendo gli step su righe separate
    e applica i testi rossi dove necessario.

    Args:
        test_cases (dict): Dizionario contenente i test case da salvare.
        output_path (str, opzionale): Percorso personalizzato del file Excel da salvare.
                                      Se non fornito, salva in ../outputs/testbook_feedbackAI.xlsx
    """
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont
    import os, re

    def apply_red_text(cell):
        """Color text in [[RED]]...[[/RED]] red, preserving the rest."""
        text = str(cell.value)
        if "[[RED]]" not in text:
            return  

        parts = re.split(r'(\[\[RED\]\]|\[\[/RED\]\])', text)
        rich_text = CellRichText()

        red = False
        for part in parts:
            if part == "[[RED]]":
                red = True
            elif part == "[[/RED]]":
                red = False
            elif part:
                font = InlineFont(color="FF0000") if red else InlineFont(color="000000")
                rich_text.append(TextBlock(font, part))

        cell.value = rich_text

    field_mapping = {
        'Canale': 'Channel',
        'Dispositivo': 'Device',
        'Sistema di riferimento': 'Reference System',
        'Modalità Operativa': 'Execution Mode',
        'Funzionalità': 'Functionality',
        'Tipologia Test': 'Test Type',
        'Test di no regression': 'No Regression Test',
        'Automation': 'Automation',
        'Risultato Atteso': 'Expected Result',
        '_polarion': '_polarion'
    }

    columns = [
        'Title', 'ID', '#', 'Test Group', 'Channel', 'Device', 
        'Priority', 'Test Stage', 'Reference System', 
        'Preconditions', 'Execution Mode', 'Functionality', 
        'Test Type', 'No Regression Test', 'Automation',
        'Dataset', 'Expected Result', 
        'Step', 'Step Description', 'Step Expected Result',
        'Country', 'Project', 'Author', 'Assignee(s)', 'Type', 
        'Partial Coverage Description', '_polarion',
        'Analysis', 'Coverage', 'Dev Complexity', 'Execution Time', 
        'Volatility', 'Developed', 'Note', 'Team Ownership', 
        'Team Ownership Note', 'Requires Script Maintenance'
    ]

    rows = []
    for tc_id, tc_data in test_cases.items():
        steps = tc_data.get('Steps', [])
        if not steps:
            steps = [{}]
        
        first = True
        for step in steps:
            row = {}
            if first:
                for col in columns:
                    if col not in ['Step', 'Step Description', 'Step Expected Result']:
                        value = tc_data.get(col, '')
                        if not value:
                            italian_key = next((k for k, v in field_mapping.items() if v == col), None)
                            if italian_key:
                                value = tc_data.get(italian_key, '')
                        row[col] = value
                first = False
            else:
                for col in columns:
                    if col not in ['Step', 'Step Description', 'Step Expected Result']:
                        row[col] = ''

            row['Step'] = step.get('Step', '')
            row['Step Description'] = step.get('Step Description', '')
            row['Step Expected Result'] = step.get('Expected Result', '')
            rows.append(row)

    df = pd.DataFrame(rows, columns=columns)

    if output_path is None:
        output_path = os.path.join(os.path.dirname(__file__), "..", "outputs", "testbook_feedbackAI.xlsx")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    df.to_excel(output_path, index=False)

    wb = load_workbook(output_path)
    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value and isinstance(cell.value, str) and "[[RED]]" in cell.value:
                apply_red_text(cell)

    wb.save(output_path)
    print(f"✅ Excel salvato con testi rossi: {output_path}")
