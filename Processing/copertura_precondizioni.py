import pandas as pd
from pathlib import Path
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
import sys
import os
from dotenv import load_dotenv
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from utils.simple_functions import process_docx

import tempfile

temp_dir = tempfile.mkdtemp()
load_dotenv()

input_excel = Path(__file__).parent.parent/"outputs"/"testcase_feedbackAAAAA.xlsx"

output_excel = input_excel.with_name(f"{input_excel.stem}_feedbackAI_precondizioni.xlsx")

embeddings = OpenAIEmbeddings(model="text-embedding-3-large")

file_word_rag = "path for rag and info"

processed_file_word_rag = process_docx(file_word_rag,temp_dir)

vectorstore = FAISS.from_texts(processed_file_word_rag, embeddings)


df_excel = pd.read_excel(input_excel)

df_precondition_missing = df_excel["Preconditions"].isna()

for idx, sample in df_excel.iterrows():
    current_precond = sample.get("Preconditions", "")
    if pd.isna(current_precond) or str(current_precond).strip() == "":
        knowledge_base_word = "documento"
        # Combina tutti i campi non nulli in un unico testo leggibile
        sample_text = " ".join(
            f"{col}: {str(val)}"
            for col, val in sample.items()
            if pd.notna(val) and str(val).strip() != ""
        )
        results = vectorstore.similarity_search_with_score(sample_text, k=1)
        print(results)

        # new_precond = generate_precondition_llm(results_vector)
        # df_excel.at[idx, "Preconditions"] = new_precond
        # print(f" Added AI precondition to test {sample['Title']}")


# print(df_excel.iterrows())

# df_excel.to_excel(output_excel, index=False)
# print(f"✅ File aggiornato salvato in: {output_excel}")


 #applicativo zenith 
#carosello filtering deve esser epiu generale
#per ora nome paragrafo su id

#dove sta backend DESKTOP->testgroup E BACKOFFICE->canale (piu sensato al contrariio) quindi testgroup backoffice
#Geo sistema di backoffice