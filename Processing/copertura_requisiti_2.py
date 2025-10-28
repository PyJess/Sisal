import pandas as pd
import sys
import os
from pathlib import Path
import json

project_root = Path(__file__).resolve().parent.parent
sys.path.append(str(project_root))
from utils.simple_functions import process_docx, AI_check_TC_requisiti, color_new_testcases_red, fill_excel_file_requisiti
from Input_extraction.extract_polarion_field_mapping import *

path_output = Path(__file__).parent.parent/"outputs"
test_case_input = Path(__file__).parent.parent/"input"/"generated_test_cases3_label_rimosse.xlsx"
documents_word = Path(__file__).parent.parent/"input"/"RU_SportsBookPlatform_SGP_Gen_FullResponsive_v1.1 - FE (2).docx"


df_testcase = pd.read_excel(test_case_input)
df_testcase_polarion = df_testcase['_polarion'].astype(str).str.lower().str.strip()
print(df_testcase_polarion.unique())

chunks,head= process_docx(documents_word,path_output)

mapping = extract_field_mapping()

# === Pipeline ===
new_testcases = []

for req in head:
    if "first line" in req.lower():
        continue

    req_norm = req.lower().strip()
    print(req_norm)
    
    if any(req_norm == p for p in df_testcase_polarion):
        print(f"✅ {req} già coperto")
        continue

    print(f"❌ {req} mancante → generazione AI...")
    context = chunks[head.index(req)]
    print("--------------------------")
    try:
        llm_new_tc = AI_check_TC_requisiti(req,context,mapping)
        print(llm_new_tc)
        if isinstance(llm_new_tc, dict) and "test_cases" in llm_new_tc:
            for tc in llm_new_tc["test_cases"]:
                tc["_polarion"] = req_norm
        elif isinstance(llm_new_tc, list):
                    for tc in llm_new_tc:
                        tc["_polarion"] = req_norm
        elif isinstance(llm_new_tc, dict):
                    llm_new_tc["_polarion"] = req_norm
            
        new_testcases.append(llm_new_tc)
        #vettorizzare file word per vector search con req e context mi restituscip outpit
        #vector search da context
        print(f"🧠 Generato test case per '{req}'")
        #req
    except Exception as e:
        print(f"⚠️ Errore generazione LLM per '{req}': {e}")

# === Salvataggio nuovi test case ===

structured_testcases = {}

counter = 1
for group in new_testcases:
    if isinstance(group, str):
        try:
            group = json.loads(group)
        except json.JSONDecodeError:
            print(f"⚠️ Gruppo non valido, salto.")
            continue

    if isinstance(group, dict) and "test_cases" in group:
        for single_tc in group["test_cases"]:
            structured_testcases[f"AI_TC_{counter}"] = single_tc
            counter += 1

    elif isinstance(group, dict):
        structured_testcases[f"AI_TC_{counter}"] = group
        counter += 1
    else:
        print(f"⚠️ Tipo non riconosciuto: {type(group)}")
print(f"\n✅ Riconosciuti {len(structured_testcases)} test case totali")
for k, v in list(structured_testcases.items())[:3]:
    print(f"{k} → {v.get('Title', '')}")

if new_testcases:
    df_new = fill_excel_file_requisiti(structured_testcases, base_columns=list(df_testcase.columns))

    output_excel = test_case_input.with_name(f"{test_case_input.stem}_feedbackAI_requisiti.xlsx")
    df_final = pd.concat([df_testcase, df_new], ignore_index=True)
    df_final.to_excel(output_excel, index=False)
    
    color_new_testcases_red(output_excel, len(df_new))

    print(f"✅ File aggiornato salvato in: {output_excel}")
else:
    print("Tutti i requisiti sono già coperti.")
    
        
        
    



