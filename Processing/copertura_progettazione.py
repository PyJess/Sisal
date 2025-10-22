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
from Processing.controllo_sintattico import prepare_prompt
from Processing.copertura_requisiti import add_new_TC, save_updated_json
from utils.simple_functions import fill_excel_file

async def gen_new_TC(paragraph, results):
        mapping = extract_field_mapping()
        print("finishing mapping")
        paragraph = paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph)
        messages, schema = await prepare_prompt(paragraph, mapping)
        print("starting calling llm")
        print(f"{messages}")
        response = await a_invoke_model(messages, schema)
        return response

async def main():

    # input_path=r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\RU_ZENIT_V_0.4_FASE_1.docx"
    input_path = Path(__file__).parent.parent / "input" / "RU_ZENIT_V_0.4_FASE_1.docx"
    print(input_path)
    print(os.path.dirname(input_path))
    paragraphs=process_docx(input_path, os.path.dirname(input_path))
    input_path = os.path.join(os.path.dirname(__file__), "..", "input", "tests_cases.xlsx")
    dic = excel_to_json(input_path) 
    print("finishing excel to json")


    new_TC=[]
    
    for i, par in enumerate(paragraphs, 1):
        print(f"\n--- Paragrafo {i}/{len(paragraphs)} ---")
        result = research_vectordb(par, dic, k=20, similarity_threshold=0.75)
        new_TC.append(await gen_new_TC(par, result))

    updated_json=add_new_TC(new_TC, dic)

    json_to_excel = fill_excel_file(updated_json)
json_test =load_json(r"C:\Users\WilliamBencich\Desktop\Sisal\updated_test_cases.json")
test = fill_excel_file(json_test,Path(__file__).parent.parent/"outputs"/"testcase_feedbackIAAAA.xlsx")

    
    # save_updated_json(updated_json, output_path='updated_test_cases.json')


#settare il path funzionante per salvalre file con come esetensione _feedbackAI
# Se non ci sono lacune → creare JSON vuoto + Excel vuoto. SOLO SEABBIAMO UN EXCEL CHE FUNGE DA DB


# Gestire il caso “no merge” quando non ci sono nuovi test.

# if __name__ == "__main__":
#     asyncio.run(main())