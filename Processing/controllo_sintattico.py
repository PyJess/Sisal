import os
import pandas as pd
from llm.llm import a_invoke_model
from utils.simple_functions import *
import asyncio
from typing import Tuple, List, Dict, Any
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.cell.rich_text import CellRichText, TextBlock
import re


async def prepare_prompt(input:str, mapping:str =None) -> Tuple[List[Dict[str, str]], Dict[str, Any]]:
    """ Prepare prompt for the LLM"""

    system_prompt= load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "controllo_sintattico", "system_prompt.txt"))
    user_prompt= load_file(os.path.join(os.path.dirname(__file__), "..", "llm", "prompts", "controllo_sintattico", "user_prompt.txt")) 
    schema= load_json(os.path.join(os.path.dirname(__file__), "..", "llm", "schema", "schema_output.json"))

    user_prompt=user_prompt.replace(f"{input}", input)
    
    if mapping and "{mapping}" in user_prompt:
        user_prompt = user_prompt.replace("{mapping}", mapping)

    messages=[
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_prompt},
    ]

    return messages, schema


async def AI_check_TC(input:str, mapping:str =None) -> Dict:

    messages, schema= await prepare_prompt(input, mapping)
    response = await a_invoke_model(messages, schema)

    return response


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
                rich_text.append(TextBlock(Font(color="FF0000") if red else Font(color="000000"), part))

        cell.value = rich_text



def fill_excel_file(llm_response: Dict):
    """Create or append to the testbook Excel file and highlight [[RED]] text in red."""
    
    data = llm_response
    df = pd.DataFrame([data])

    excel_path = os.path.join(os.path.dirname(__file__), "..", "outputs", "testbook.xlsx")
    os.makedirs(os.path.dirname(excel_path), exist_ok=True)

    if os.path.exists(excel_path):
        df_existing = pd.read_excel(excel_path)
        df_combined = pd.concat([df_existing, df], ignore_index=True)
    else:
        df_combined = df

    df_combined.to_excel(excel_path, index=False)

    wb = load_workbook(excel_path)
    ws = wb.active

    last_row = ws.max_row
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=last_row, column=col)
        if isinstance(cell.value, str) and "[[RED]]" in cell.value:
            apply_red_text(cell)

    for col in ws.columns:
        max_length = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_length + 2, 80)

    wb.save(excel_path)
    print(f"Test case added with red highlights in {excel_path}")





