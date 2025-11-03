import fitz
from docx import Document
import re
if not hasattr(fitz.Rect, "get_area"):
    fitz.Rect.get_area = lambda self: self.width * self.height

from pdf2docx import Converter

pdf_path = r"input\PDF_test\20210930_REGOLAMENTO_SVT_2021.pdf"
docx_path = r"input\PDF_test\20210930_REGOLAMENTO_SVT_2021.docx"

# cv = Converter(pdf_path)
# cv.convert(docx_path, start=0, end=None)
# cv.close()


def parse_docx_structured(docx_path):
    doc = Document(docx_path)
    
    struttura = []
    current_section = {"title": None, "content": []}

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        
        if re.match(r"^(Articolo\s+\d+)", text, re.IGNORECASE):
            # Salva la sezione precedente
            if current_section["title"]:
                struttura.append(current_section)
                current_section = {"title": None, "content": []}
            
            current_section["title"] = text
        else:
            current_section["content"].append(text)
    
    # Aggiungi l'ultima sezione
    if current_section["title"]:
        struttura.append(current_section)

    return struttura

# --- ESEMPIO D'USO ---
sections = parse_docx_structured(docx_path)

for s in sections:
    print(f"\n--- {s['title']} ---")
    print("\n".join(s['content'][:3]))  
