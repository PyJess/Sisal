from dotenv import load_dotenv
load_dotenv()

import os

from openai import OpenAI
client = OpenAI()

def load_file(filepath:str):
    with open(filepath, encoding="utf-8") as f:
        return f.read()

# carica il pdf
pdf_file = client.files.create(
    file=open(r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\Figma\DOC1.pdf", "rb"),
    purpose="assistants"
)

system_prompt = load_file(os.path.join(os.path.dirname(__file__), "system_prompt_figma_prova.txt"))
# Chiedi al modello di processarlo
response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "user", "content": [
            {"type": "text", "text": system_prompt},
            {"type": "file", "file": {"file_id": pdf_file.id}}
        ]}
    ]
)

print(response.choices[0].message.content)
