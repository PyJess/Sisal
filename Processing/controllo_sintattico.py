import os
from llm.llm import a_invoke_model
from utils.simple_functions import *
import asyncio
from typing import Tuple, List, Dict, Any


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


