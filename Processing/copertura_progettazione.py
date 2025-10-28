import os
import sys
from pathlib import Path
import asyncio

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.simple_functions import process_docx, excel_to_json
from Processing.copertura_requisiti import research_vectordb
from Input_extraction.extract_polarion_field_mapping import *
from utils.simple_functions import *
from llm.llm import a_invoke_model
#from Processing.controllo_sintattico import prepare_prompt
from Processing.copertura_requisiti import add_new_TC, save_updated_json
from utils.simple_functions import fill_excel_file
from typing import List, Dict, Any, Tuple


async def prepare_prompt(input: Dict,excel: Dict, mapping: str = None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""
    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_progettazione", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "copertura_progettazione", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output.json"))

    user_prompt = user_prompt.replace("{input}", json.dumps(input))
    mapping_as_string = mapping.to_json() 
    user_prompt = user_prompt.replace("{mapping}", mapping_as_string)
    user_prompt = user_prompt.replace("{TC}", json.dumps(excel))

    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema


async def gen_new_TC(paragraph,title, results,mapping):
        print("finishing mapping")
        title = 
        paragraph = paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph)
        messages, schema = await prepare_prompt(paragraph,results, mapping)
        print("starting calling llm")
        print(f"{messages}")
        response = await a_invoke_model(messages, schema, model="gpt-4.1")
        return response

async def main():

    # input_path=r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\RU_ZENIT_V_0.4_FASE_1.docx"
    input_path = Path(__file__).parent.parent / "input" / "RU_ZENIT_V_0.4_FASE_1.docx"
    print(input_path)
    
    print(os.path.dirname(input_path))
    paragraphs,title=process_docx(input_path, os.path.dirname(input_path))
    excel_path = Path(os.path.join(os.path.dirname(__file__), "..", "input", "tests_cases.xlsx"))
    
    dic = excel_to_json(excel_path) 
    print("finishing excel to json")
    mapping = extract_field_mapping()


    new_TC=[]
    
    for i, par in enumerate(title, 1):
        print(f"\n--- Paragrafo {i}/{len(paragraphs)} ---")
        result = research_vectordb(par, dic, k=20, similarity_threshold=0.75)
        #cercare title nella colonna polarion -> prendere tutti test che hanno quel title
        #mettere  content_title in gen_new_tc
        #dire al prompt di confrontare i testcase presenti col contenuto del title  IF all case lasciare vuoto array return else genera
        new_TC.append(await gen_new_TC(par, result,mapping))
        print(new_TC)

    updated_json=add_new_TC(new_TC, dic)
    
    
    json_to_excel = fill_excel_file(updated_json, excel_path.with_name(f"{excel_path.stem}_feedbackAI_testcase_progettazione.xlsx"))


## PER TESTING JSON INTO EXCEL
# test = fill_excel_file(json_test,Path(__file__).parent.parent/"outputs"/"testcase_feedbackAAAAA.xlsx")

    
    # save_updated_json(updated_json, output_path='updated_test_cases.json')


# Se non ci sono lacune → creare JSON vuoto + Excel vuoto. SOLO SEABBIAMO UN EXCEL CHE FUNGE DA DB


# Gestire il caso “no merge” quando non ci sono nuovi test.


if __name__ == "__main__":
    asyncio.run(main())
