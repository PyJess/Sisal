import sys
import os
from pathlib import Path
from openpyxl import Workbook
import json
from docx import Document
import subprocess
from langchain_openai import OpenAIEmbeddings
from langchain.vectorstores import FAISS
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.simple_functions import group_by_funzionalita, load_json
from Processing.controllo_sintattico import *
import json

data= load_json("C:\\Users\\x.hita\\OneDrive - Reply\\Workspace\\Sisal\\Test_Design\\input\\tests_output.json")

embedding_model = "text-embedding-3-large"
PANDOC_EXE = "pandoc" 


import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import os

async def save_test_cases_to_excel(existing_tests, new_tests, output_path):
    """
    Salva i test case in Excel, mantenendo gli step su righe separate
    e colorando in rosso i nuovi test cases.
    """
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
    
    # Definisci colonne finali
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

    # Combina i test esistenti e nuovi
    all_tests = []
    
    # Aggiungi test esistenti
    for test in existing_tests:
        all_tests.append((test, False))  # False = non nuovo
    
    # Aggiungi test nuovi
    for test in new_tests:
        all_tests.append((test, True))  # True = nuovo
    
    rows = []
    row_metadata = []  # Track which rows belong to new tests
    
    for tc_data, is_new in all_tests:
        steps = tc_data.get('Steps', [])
        
        if not steps:
            steps = [{}]
        
        first = True
        for step in steps:
            row = {}

            if first:
                # Prima riga: tutti i dati del test case
                for col in columns:
                    if col not in ['Step', 'Step Description', 'Step Expected Result']:
                        value = tc_data.get(col, '')
                        
                        # Fallback per chiavi italiane
                        if not value:
                            italian_key = next((k for k, v in field_mapping.items() if v == col), None)
                            if italian_key:
                                value = tc_data.get(italian_key, '')
                        
                        row[col] = value
                first = False
            else:
                # Righe successive: solo gli step
                for col in columns:
                    if col not in ['Step', 'Step Description', 'Step Expected Result']:
                        row[col] = ''

            # Aggiungi dati dello step
            row['Step'] = step.get('Step', '')
            row['Step Description'] = step.get('Step Description', '')
            row['Step Expected Result'] = step.get('Step Expected Result', step.get('Expected Result', ''))
            
            rows.append(row)
            row_metadata.append(is_new)  # Traccia se questa riga è nuova

    df = pd.DataFrame(rows, columns=columns)

    # Crea directory se non esiste
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Salva in Excel
    df.to_excel(output_path, index=False)

    # Applica formattazione rossa ai nuovi test
    wb = load_workbook(output_path)
    ws = wb.active
    
    red_font = Font(color="FF0000")

    for idx, is_new in enumerate(row_metadata, start=2):  # start=2 per saltare l'header
        if is_new:
            for cell in ws[idx]:  # Colora tutte le celle della riga
                if cell.value:  # Solo se c'è un valore
                    cell.font = red_font

    wb.save(output_path)
    print(f"✅ Excel salvato con nuovi TC in rosso: {output_path}")
    print(f"   - TC esistenti: {len([t for t, is_new in all_tests if not is_new])}")
    print(f"   - TC nuovi (in rosso): {len([t for t, is_new in all_tests if is_new])}")


async def prepare_prompt_application(input: str, mapping: str = None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""
    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_applicativi", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_applicativi", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output_applicativi.json"))

    user_prompt = user_prompt.replace("{input}", input)

    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema


async def AI_check_applications(input: str, mapping: str = None) -> Dict:
    messages, schema = await prepare_prompt_application(input)
    print("starting calling llm")
    #print(f"{messages}")
    response = await a_invoke_model(messages, schema, model="gpt-4.1")
    return response


async def prepare_prompt(input: str, context:str = None, mapping: str = None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""
    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "test_design", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "test_design", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output.json"))

    user_prompt = user_prompt.replace("{input}", input)
    print(f"input {input}")
    user_prompt = user_prompt.replace("{context}", context )
    mapping= mapping.to_json() 
    user_prompt = user_prompt.replace("{mapping}", mapping )
    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema


async def AI_gen_TC(input: str, context:str, mapping: str = None) -> Dict:
    messages, schema = await prepare_prompt(input, context, mapping)
    print("starting calling llm")
    #print(f"{messages}")
    response = await a_invoke_model(messages, schema, model="gpt-4.1")
    return response

async def main():
    excel_path = os.path.join(os.path.dirname(__file__), "..", "outputs", "generated_test_cases3_withoutDHW.xlsx")
    dic = excel_to_json(excel_path) 
    print("finishing excel to json")

    mapping = extract_field_mapping()
    # print("finishing mapping")
    
    application_list = []
    input_path=os.path.join(os.path.dirname(__file__), "..", "input", "RU_SportsBookPlatform_SGP_Gen_FullResponsive_v1.1 - FE (2).docx")
    requirements, name = process_docx(input_path, os.path.dirname(input_path))
    
    # rag_path= input_path=os.path.join(os.path.dirname(__file__), "..", "input", "RU_SportsBookPlatform_SGP_Gen_FullResponsive_v1.1 - FE (2).docx")
    # chunks, _ = process_docx(input_path, os.path.dirname(rag_path))
    # embeddings = OpenAIEmbeddings(model=embedding_model)
    # chunks= chunks + requirements
    # vectorstore = FAISS.from_texts(chunks, embeddings)
    
    tasks = [AI_check_applications(input=req) for req in requirements]
    application_list.append( await asyncio.gather(*tasks))

    # Eliminates duplicates based on application_name
    unique_apps = {}
    for sublist in application_list:
        for result in sublist:
            for app in result.get("applications", []):
                name = app.get("application_name")
                text = app.get("specific_text")
                if name: 
                    if name not in unique_apps:
                        unique_apps[name] = {
                            "application_name": name.strip(),
                            "specific_text": text.strip() if text else ""
                        }

    application_list = [
        {"applications": list(unique_apps.values())}
    ]

    print("Applications found: \n\n")
    for app in application_list[0]["applications"]:
        print(app.get("application_name"))


    for result in application_list:
        new_TC=[]
        for app in result.get("applications", []):
            app_name = app.get("application_name", "").strip()
            print(f"\n{'='*50}")
            print(f"Checking application: '{app_name}'")
            present = False
            
            test_cases = dic.values() 
            
            print(f"Total test cases to check: {len(list(dic.values()))}")
            
            for tc_id, test_case in dic.items(): 
                if not isinstance(test_case, dict):
                    print(f"  {tc_id}: Not a dict, skipping")
                    continue
                
                # Controlla Title
                title = test_case.get("Title", "")
                
                if app_name.lower() in title.lower():
                    print(f"  {tc_id}: FOUND in title!")
                    present = True
                    break
                
                # Controlla tutti gli step
                steps = test_case.get("Steps") or test_case.get("Step")
                if steps:
                    for j, step in enumerate(steps):
                        for field in ["Step Description", "Expected Result"]:
                            value = step.get(field, "")
                            if app_name.lower() in str(value).lower():
                                print(f"  {tc_id}, Step {j}: FOUND in {field}!")
                                present = True
                                break
                        if present:
                            break
                if present:
                    break
            
            if present:
                print(f"{app_name}: presente")
            else:
                print(f"{app_name}: non presente")
                app_text= app.get("specific_text", "")
                #fare il retrived del context in base al app_text
                context=""
                context= str(context)
                new_TC.append(await AI_gen_TC(app_text, context, mapping))


    combined_test_cases= list(dic.values()) + new_TC
    print(f"\nTotal existing TCs: {len(dic)}")
    print(f"Total new TCs generated: {len(new_TC)}")
    print(f"Total combined TCs: {len(combined_test_cases)}")
    #print(combined_test_cases)
    output_excel = os.path.join(os.path.dirname(__file__), "..", "outputs", "combined_test_cases.xlsx")
    await save_test_cases_to_excel(dic.values(), new_TC, output_excel)


    
    

if __name__ == "__main__":
    asyncio.run(main())

