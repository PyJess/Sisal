import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.simple_functions import process_docx, excel_to_json
from Processing.copertura_requisiti import research_vectordb

from Input_extraction.extract_polarion_field_mapping import *
from utils.simple_functions import *
from llm.llm import a_invoke_model
from Processing.controllo_sintattico import prepare_prompt
from Processing.copertura_requisiti import add_new_TC, save_updated_json


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

    input_path=r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\RU_ZENIT_V_0.4_FASE_1.docx"
    print(os.path.dirname(input_path))
    paragraphs=process_docx(input_path, os.path.dirname(input_path))
    input_path = os.path.join(os.path.dirname(__file__), "..", "input", "tests_cases.xlsx")
    dic = excel_to_json(input_path) 
    print("finishing excel to json")


    new_TC=[]
    for i, par in enumerate(paragraphs, 1):
        print(f"\n--- Paragrafo {i}/{len(paragraphs)} ---")
        result = research_vectordb(par, dic, k=100 similarity_threshold=0.75)
        new_TC.append(gen_new_TC(par, result))

    updated_json=add_new_TC(new_TC, dic)
    save_updated_json(updated_json, output_path='updated_test_cases.json')