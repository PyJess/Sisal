import sys
import os

from pathlib import Path
# from docx import Document
import subprocess

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from utils.simple_functions import group_by_funzionalita, load_json
from Processing.controllo_sintattico import *
import json

data= load_json(Path(__file__).parent.parent/"input/tests_output.json")

grouped_json=group_by_funzionalita(data)

PANDOC_EXE = "pandoc" 

def process_docx_sections(docx_path, output_base)-> list:
        """
        Process a DOCX file using Pandoc and split it into sections based on Markdown headers (#, ##, etc.).
        Also extracts embedded images and maps them to their respective sections.
        """

        #images_folder = os.path.join(output_base, "Images", Path(docx_path).stem)
        #os.makedirs(images_folder, exist_ok=True)
        txt_output_path = os.path.join(output_base, Path(docx_path).stem + ".txt")

        docx_path = os.path.normpath(docx_path)
        txt_output_path = os.path.normpath(txt_output_path)
        #images_folder = os.path.normpath(images_folder)

        # Step 1: Convert DOCX to Markdown using Pandoc
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
            return [], os.path.basename(docx_path), {}, []

        # Step 2: Read the Markdown output
        with open(txt_output_path, "r", encoding="utf-8") as f:
            text_lines = f.read().splitlines()

        # Step 3: Find headers and define section boundaries
        headers = []
        heading_list = []

        for index, line in enumerate(text_lines):
            if line.startswith("#"):
                level = line.count("#")
                clean_name = line.replace("#", "").strip()
                headers.append([clean_name, index, level])
                heading_list.append([clean_name, level])

        # Add artificial first/last headers
        headers.insert(0, ["== first line ==", 0, 0])
        headers.append(["== last line ==", len(text_lines), 0])
        heading_list.insert(0, ["== first line ==", 0])
        heading_list.append(["== last line ==", 0])

        # Step 4: Split text into section chunks
        chunks = []
        for i in range(len(headers) - 1):
            start_idx = headers[i][1]
            end_idx = headers[i + 1][1]
            section_lines = text_lines[start_idx:end_idx]
            chunk_text = "\n".join(section_lines).strip()

            # Optional: Clean up Pandoc artifacts
            header_cleaned = re.sub(r"\s*\{.*?\}", "", headers[i][0])
            header_cleaned = header_cleaned.replace("--", "–").strip(" *[]\n")
            chunk_text = header_cleaned + "\n" + chunk_text

            chunks.append(chunk_text)
            print(chunk_text)
        print(len(chunks))

        return chunks
doc_path = Path(__file__).parent.parent / "input" / "RU_ZENIT_V_0.4_FASE_1.docx"
output_base = Path(__file__).parent/"/outputs"

chunks=process_docx_sections(doc_path,output_base)

sys.exit()
async def main():
    input_path = os.path.join(os.path.dirname(__file__), "..", "input", "tests_cases.xlsx")
    dic = excel_to_json(input_path) 
    print("finishing excel to json")
    grouped_json=group_by_funzionalita(dic)

    mapping = extract_field_mapping()
    print("finishing mapping")
        
    requirements= process_docx_sections() #TODO
    tasks = [AI_check_TC(input=req, mapping=mapping) for req in requirements]
    results_list = await asyncio.gather(*tasks)
    print("finishing gpt call")
    
    merged_results = {tc["ID"]: tc for tc in results_list}


    print(f"Merged result: {merged_results}")

    fill_excel_file(merged_results)


if __name__ == "__main__":
    asyncio.run(main())

