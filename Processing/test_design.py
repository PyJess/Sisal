from langchain.document_loaders import UnstructuredWordDocumentLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain.vectorstores import FAISS
import pandas as pd
from langchain_openai import ChatOpenAI
from docx import Document
import sys
import os
from typing import Dict, List, Tuple, Any
os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
import asyncio

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from Input_extraction.extract_polarion_field_mapping import *
from utils.simple_functions import *
from llm.llm import a_invoke_model
from Processing.copertura_requisiti import add_new_TC, save_updated_json

embedding_model = "text-embedding-3-large"
PANDOC_EXE = "pandoc" 


async def prepare_prompt(input: Dict, mapping: str = None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """Prepare prompt for the LLM"""

    system_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "test_design", "system_prompt.txt"))
    user_prompt = load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "test_design", "user_prompt.txt")) 
    schema = load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output.json"))

    user_prompt = user_prompt.replace("{input}", json.dumps(input))
    mapping_as_string = mapping.to_json() 
    user_prompt = user_prompt.replace("{mapping}", mapping_as_string)
    print("finishing prepare prompt")

    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema

async def gen_TC(paragraph):
        print("This requirement has no TC")
        mapping = extract_field_mapping()
        print("finishing mapping")
        paragraph = paragraph.page_content if hasattr(paragraph, 'page_content') else str(paragraph)
        messages, schema = await prepare_prompt(paragraph, mapping)
        print("starting calling llm")
        print(f"{messages}")
        response = await a_invoke_model(messages, schema)
        print("✅ File Excel generato con successo!")
        return response


def merge_TC(new_TC):
    """
    Merge all the test cases in one json

    """
    all_test_cases = []
    
    for tc in new_TC:
        if tc is None:
            continue
            
        if isinstance(tc, list):
            all_test_cases.extend(tc)
        
        elif isinstance(tc, dict) and "test_cases" in tc:
            test_cases = tc["test_cases"]
            if isinstance(test_cases, list):
                all_test_cases.extend(test_cases)
            else:
                all_test_cases.append(test_cases)
        
        elif isinstance(tc, dict):
            all_test_cases.append(tc)
    
    return {
        "test_cases": all_test_cases,
        "total_count": len(all_test_cases)
    }


async def main():

    input_path=r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\RU_ZENIT_V_0.4_FASE_1.docx"
    print(os.path.dirname(input_path))
    paragraphs=process_docx(input_path, os.path.dirname(input_path))

    new_TC=[]
    for i, par in enumerate(paragraphs, 1):
        print(f"\n--- Paragrafo {i}/{len(paragraphs)} ---")

        tc = await gen_TC(par)
        new_TC.append(tc)

    updated_json=merge_TC(new_TC)
    save_updated_json(updated_json, output_path='updated_test_cases.json')



if __name__ == "__main__":
    asyncio.run(main())





