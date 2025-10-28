from langchain_community.document_loaders import UnstructuredWordDocumentLoader
# from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS

# from langchain.document_loaders import UnstructuredWordDocumentLoader

# from langchain.text_splitter import RecursiveCharacterTextSplitter
# from langchain_openai import OpenAIEmbeddings
# from langchain.vectorstores import FAISS
import pandas as pd
from langchain_openai import ChatOpenAI
# from docx import Document
import sys
import os
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
import asyncio


sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from Input_extraction.extract_polarion_field_mapping import *
from utils.simple_functions import *
from llm.llm import a_invoke_model
from typing import Tuple, List, Dict, Any


embedding_model = "text-embedding-3-large"
PANDOC_EXE = "pandoc" 


async def prepare_prompt(input: Dict, results: Dict, mapping: str = None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""
    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_requisiti", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_requisiti", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output1.json"))

    user_prompt = user_prompt.replace("{input}", json.dumps(input))
    #mapping_as_string = mapping.to_json() 
    user_prompt = user_prompt.replace("{mapping}", str(mapping))
    user_prompt = user_prompt.replace("{TC}", str(results))

    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema

def prepare_test_texts(excel):
    test_texts = []
    for tc_id, tc in excel.items():
        steps_text = ""
        for step in tc.get("Steps", []):
            steps_text += f"Step {step['Step']}: {step['Step Description']}. Expected: {step['Expected Result']}. "
        combined_text = f"Title: {tc['Title']}. Functionality: {tc.get('Funzionalità','')}. Preconditions: {tc.get('Precondizioni','')}. Steps: {steps_text}"
        test_texts.append(combined_text)
    return test_texts


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



async def gen_TC(paragraph, results):
    if results["Closest_Test"] is not None:
        print(f" Requirement già coperto da: {results['Closest_Test'][:100]}...")
        return None
    else:
        print("This requirement has no TC")
        mapping = extract_field_mapping()
        print("finishing mapping")
        paragraph = paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph)
        messages, schema = await prepare_prompt(paragraph, mapping, results)
        print("starting calling llm")
        #print(f"{messages}")
        response = await a_invoke_model(messages, schema, model="gpt-4.1")
        print("File Excel generato con successo!")


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

            
def save_updated_json(updated_json, output_path='updated_test_cases.json'):
    """
    Salva il JSON aggiornato su file.
    
    Args:
        updated_json (dict): JSON con i test cases aggiornati
        output_path (str): Percorso del file di output
    """
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(updated_json, f, ensure_ascii=False, indent=2)
    print(f"JSON aggiornato salvato in: {output_path}")



async def main():
    

    input_path= os.path.join(os.path.dirname(__file__), "..", "input", "Esempio 2", "RU_Sportsbook_Platform_Fantacalcio_Prob. Form_v0.2 (1).docx")
    print(os.path.dirname(input_path))
    paragraphs, headers =process_docx(input_path, os.path.dirname(input_path))
    input_path = os.path.join(os.path.dirname(__file__), "..", "outputs", "generated_test_cases3 - Copy.xlsx")
    dic = excel_to_json(input_path) 
    print("finishing excel to json")

    new_TC=[]
    for i, par in enumerate(paragraphs, 1):
        print(f"\n--- Paragrafo {i}/{len(paragraphs)} ---")
        result = research_vectordb(par, dic, k=20, similarity_threshold=0.65)
        tc = await gen_TC(par, result)
        new_TC.append(tc)

    updated_json=add_new_TC(new_TC, dic)
    save_updated_json(updated_json, output_path='updated_test_cases1.json')



if __name__ == "__main__":
    asyncio.run(main())
