import os
import sys
from pathlib import Path
import asyncio

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.simple_functions import process_docx, excel_to_json
from Input_extraction.extract_polarion_field_mapping import *
from utils.simple_functions import *
from llm.llm import a_invoke_model
#from Processing.controllo_sintattico import prepare_prompt
from utils.simple_functions import fill_excel_file_progettazione,color_new_testcases_red

from typing import List, Dict, Any, Tuple
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS

embedding_model = "text-embedding-3-large"

async def prepare_prompt(input: str,excel: Dict, mapping: str = None, title: str=None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""
    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_progettazione", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_progettazione", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output.json"))

    user_prompt = user_prompt.replace("{input}", str(input))
    user_prompt = user_prompt.replace("{title}", title if title else "No title available")

    mapping_as_string = mapping.to_json() 
    user_prompt = user_prompt.replace("{mapping}", mapping_as_string)
    user_prompt = user_prompt.replace("{TC}", json.dumps(excel))

    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema


async def gen_new_TC(paragraph, title, results, mapping):
    print("finishing mapping")
    
    paragraph_text = paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph)
    
    # 🔹 Passiamo anche il titolo al prompt
    messages, schema = await prepare_prompt(paragraph_text, results, mapping, title)
    
    print("starting calling llm")
    print(f"{messages}")
    
    response = await a_invoke_model(messages, schema, model="gpt-4.1")

    if not response or response == []:
        print(f"⚠️ Nessun nuovo test generato per: {title}")
    else:
        print(f"🧠 Generati nuovi test per: {title}")

    return response

def add_new_TC(new_TC, original_excel):

    new_TC_list = [tc for tc in new_TC if tc is not None]

    if not new_TC_list:
        print("Nessun nuovo TC da aggiungere (tutti i requirement sono già coperti)")
        return original_excel

    field_mapping = {
        'Title': 'Title',
        'Test Group': 'Test Group',
        'Channel': 'Canale',
        'Device': 'Dispositivo',
        'Priority': 'Priority',
        'Test Stage': 'Test Stage',
        'Reference System': 'Sistema di riferimento',
        'Preconditions': 'Precondizioni',
        'Execution Mode': 'Modalità Operativa',
        'Functionality': 'Funzionalità',
        'Test Type': 'Tipologia Test',
        'Dataset': 'Dataset',
        'Expected Result': 'Risultato Atteso',
        'Country': 'Country',
        'Type': 'Type',
        '_polarion': '_polarion'
    }

    all_columns = set()
    for test_data in original_excel.values():
        all_columns.update(test_data.keys())

    max_number = max(
        (int(test_data.get('#', 0)) for test_data in original_excel.values() 
         if '#' in test_data and isinstance(test_data['#'], (int, float))),
        default=0
    )

    for new_test in new_TC:
        test_id = new_test.get('ID', '')
        max_number += 1

        new_test_case = {}
        for col in all_columns:
            if col == 'Steps':
                new_test_case[col] = []
            else:
                new_test_case[col] = ''
        new_test_case['#'] = max_number

        for ai_field, json_field in field_mapping.items():
            if ai_field in new_test and new_test[ai_field] is not None and new_test[ai_field] != '':
                new_test_case[json_field] = new_test[ai_field]

        if 'Steps' in new_test and new_test['Steps']:
                new_test_case['Steps'] = new_test['Steps']
        original_excel[test_id] = new_test_case
    return original_excel


def research_vectordb(paragraph, excel, k=20, similarity_threshold=0.65):
        test_texts = prepare_test_texts(excel)
        embeddings = OpenAIEmbeddings(model=embedding_model)
        vectorstore = FAISS.from_texts(test_texts, embeddings)
        num_documents = len(vectorstore.index_to_docstore_id)
        print(f"Numero di documenti salvati nel vectorstore: {num_documents}")
        docs_found = vectorstore.similarity_search_with_score(paragraph, k)
        print(f"Documents found: {len(docs_found)}")
        closest_doc, score = docs_found[0]
        print(f"Score: {1 - (score / 2)}")

        matching_docs = []

        for doc, score in docs_found:
            normalized_score = 1 - (score / 2)  # normalizzazione
            if normalized_score >= similarity_threshold:
                matching_docs.append({
                    "content": doc.page_content,
                    "score": normalized_score
                })

        


        result = {
            "Paragraph": paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph),
            "Closest_Test": matching_docs if matching_docs else None
        }
        return result


async def main():

    input_path = Path(__file__).parent.parent / "input" / "RU_SportsBookPlatform_SGP_Gen_FullResponsive_v1.1 - FE (2).docx"
    
    paragraphs,title=process_docx(input_path, os.path.dirname(input_path))

    # paragraphs = paragraphs
    # title = title
    print(title)    
    print("***************************")
    #print(paragraphs)
    print(f"Test limitato ai primi {len(paragraphs)} paragrafi")

    excel_path = Path(os.path.join(os.path.dirname(__file__), "..", "input", "generated_test_cases3_label_rimosse.xlsx"))
    
    dic = excel_to_json(excel_path) 
    print("finishing excel to json")
    mapping = extract_field_mapping()


    new_TC=[]
    
    for i, par in enumerate(paragraphs, 1):
        current_title = title[i-1] if i-1 < len(title) else "Untitled"
        print(f"\n--- Paragrafo {i}/{len(paragraphs)}: {current_title} ---")

        if "first line" in current_title.lower():
            print(f" Skipping title: {current_title}")
            continue

        # Cerca nel file Excel i test case che contengono questo titolo in _polarion
        matching_tests = [
            t for t in dic.get("test_cases", [])
            if current_title.lower() in str(t.get("_polarion", "")).lower()
        ]
        print(f" Mattching tests found: {matching_tests}")

        #result = research_vectordb(par, dic, k=5, similarity_threshold=0.75)

        # Genera nuovi test case (se necessari)
        # llm_new_tc = await gen_new_TC(par, current_title, matching_tests, mapping)

        # # Aggiungi il titolo corrente come valore del campo _polarion
        # if isinstance(llm_new_tc, dict) and "test_cases" in llm_new_tc:
        #     for tc in llm_new_tc["test_cases"]:
        #         tc["_polarion"] = current_title
        #     new_TC.append(llm_new_tc)
        # elif isinstance(llm_new_tc, list):
        #     for tc in llm_new_tc:
        #         tc["_polarion"] = current_title
        #     new_TC.extend(llm_new_tc)
        # elif isinstance(llm_new_tc, dict):
        #     llm_new_tc["_polarion"] = current_title
        #     new_TC.append(llm_new_tc)

        # print(new_TC)
    sys.exit(0)
    updated_json=add_new_TC(new_TC, dic)
    
    
    json_to_excel = fill_excel_file_progettazione(updated_json, excel_path.with_name(f"{excel_path.stem}_feedbackAI_testcase_progettazione.xlsx"))

        
    # 2Colora le ultime righe aggiunte (ad esempio len(new_TC))
    new_rows_count = sum(len(tc_block.get("test_cases", [])) for tc_block in new_TC if tc_block)
    color_new_testcases_red(excel_path.with_name(f"{excel_path.stem}_feedbackAI_testcase_progettazione.xlsx"), new_rows_count)
    
## PER TESTING JSON INTO EXCEL
# test = fill_excel_file(json_test,Path(__file__).parent.parent/"outputs"/"testcase_feedbackAAAAA.xlsx")
# # save_updated_json(updated_json, output_path='updated_test_cases.json')


if __name__ == "__main__":
    asyncio.run(main())
