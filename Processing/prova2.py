from dotenv import load_dotenv
import os
from openai import OpenAI
import asyncio
import base64
import sys
from typing import Dict, List, Tuple, Any
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from llm.llm import a_invoke_model_without_schema
from utils.simple_functions import load_file

load_dotenv()

client = OpenAI()

async def AI_check_TC(messages) -> Dict:
    """Call LLM to check test case syntax"""
    print("starting calling llm")
    print(f"{messages}")
    response = await a_invoke_model_without_schema(messages, model="gpt-4.1")
    return response


PDF_FOLDER = r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\Figma\Tabellone"


def encode_image(image_path: str) -> str:
    """
    Encode an image file as a base64 string.

        Parameters:
            image_path (str): The file path to the image to be encoded

        Returns:
            str: a base64 encoded string representation of the image
    """
    try:
        with open(image_path, "rb") as image_file:
            image_data = image_file.read()
            if not image_data:
                print(f"Il file {image_path} è vuoto")
                return None
            return base64.b64encode(image_data).decode('utf-8')
    except FileNotFoundError as e:
        print(f"Immagine  {image_path} non trovata: {e}")
        return None
    except IOError as e:
        print(f"Errore nella lettura dell'immagine {image_path}: {e}")
        return None


async def main(folder_path):
    pdf_files = []
    
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        file_ext = filename.lower()
        
        if file_ext.endswith(".pdf"):
            print(f"Caricamento file PDF: {filename} ...")
            pdf_obj = client.files.create(
                file=open(filepath, "rb"),
                purpose="assistants"
            )
            pdf_files.append(pdf_obj)

    print(f"{len(pdf_files)} file PDF caricati con successo.")

    # Carica il system prompt
    system_prompt_path = os.path.join(os.path.dirname(__file__), "system_prompt_figma_prova.txt")
    with open(system_prompt_path, "r", encoding="utf-8") as f:
        system_prompt = f.read()
    
    # Costruisci il message_content come prima
    message_content = [{"type": "text", "text": system_prompt}]
    
    # Aggiungi i PDF
    for pdf in pdf_files:
        message_content.append({"type": "file", "file": {"file_id": pdf.id}})
    
    # Aggiungi le immagini se presenti
    for filename in os.listdir(folder_path):
        filepath = os.path.join(folder_path, filename)
        file_ext = filename.lower()
        
        if file_ext.endswith((".png", ".jpeg", ".jpg")):
            print(f"Caricamento immagine: {filename} ...")
            base64_image = encode_image(filepath)
            if base64_image:
                # Determina il media type
                if file_ext.endswith(".png"):
                    media_type = "image/png"
                else:
                    media_type = "image/jpeg"
                
                message_content.append({
                    "type": "image_url",
                    "image_url": {
                        "url": f"data:{media_type};base64,{base64_image}"
                    }
                })

    response = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "user", "content": message_content}
        ]
    )

    print("\n=== RISPOSTA DEL MODELLO ===\n")
    
    # Crea la struttura messages per AI_check_TC
    messages = [
        {"role": "user", "content": message_content}
    ]
    
    response = await AI_check_TC(messages)
    print(response.content)


    message_content_without_system = message_content[1:]

    messages = [
        {"role": "user", "content": message_content_without_system}
    ]
    print(messages)


if __name__ == "__main__":
    asyncio.run(main(r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\Figma\Tabellone"))