import base64
from pathlib import Path
from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage
import fitz  # PyMuPDF
from PIL import Image
import io

def extract_pdf_content(pdf_path: str, save_images: bool = True, output_folder: str = "extracted_images"):
    """
    Estrae testo e immagini da un PDF.
    
    Args:
        pdf_path: Path al file PDF
        save_images: Se True, salva le immagini su disco
        output_folder: Cartella dove salvare le immagini
    
    Returns:
        dict con 'text', 'images' (list di base64) e 'image_paths' (paths salvati)
    """
    doc = fitz.open(pdf_path)
    
    full_text = ""
    images = []
    image_paths = []
    
    # Crea la cartella per le immagini se richiesto
    if save_images:
        output_path = Path(output_folder)
        output_path.mkdir(exist_ok=True)
        print(f"Cartella immagini: {output_path.absolute()}")
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Estrai testo
        full_text += f"\n--- Pagina {page_num + 1} ---\n"
        full_text += page.get_text()
        
        # Estrai immagini
        image_list = page.get_images(full=True)
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            
            # Converti in base64
            image_base64 = base64.b64encode(image_bytes).decode('utf-8')
            image_ext = base_image["ext"]
            
            # Salva l'immagine su disco se richiesto
            if save_images:
                image_filename = f"page{page_num + 1}_img{img_index + 1}.{image_ext}"
                image_path = output_path / image_filename
                
                with open(image_path, "wb") as img_file:
                    img_file.write(image_bytes)
                
                image_paths.append(str(image_path))
                print(f"  Salvata: {image_filename}")
            
            images.append({
                "data": image_base64,
                "format": image_ext,
                "page": page_num + 1,
                "filename": image_filename if save_images else None
            })
    
    doc.close()
    
    return {
        "text": full_text,
        "images": images,
        "image_paths": image_paths
    }


def process_pdf_with_langchain(pdf_path: str, query: str, api_key: str = None):
    """
    Processa un PDF (con testo e immagini) usando LangChain e OpenAI.
    
    Args:
        pdf_path: Path al file PDF
        query: Domanda o istruzione per l'LLM
        api_key: API key di OpenAI (opzionale, usa variabile ambiente OPENAI_API_KEY)
    
    Returns:
        Risposta dell'LLM
    """
    
    # 1. Estrai contenuto dal PDF
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF non trovato: {pdf_path}")
    
    print("Estraendo contenuto dal PDF...")
    content = extract_pdf_content(pdf_path, save_images=True, output_folder="extracted_images")
    
    # 2. Inizializza il modello
    llm = ChatOpenAI(
        model="gpt-4o",  # Supporta visione e testo
        api_key=api_key,
        temperature=0
    )
    
    # 3. Costruisci il messaggio con testo e immagini
    message_content = [
        {
            "type": "text",
            "text": f"{query}\n\n=== CONTENUTO TESTUALE DEL PDF ===\n{content['text']}"
        }
    ]
    
    # Aggiungi le immagini
    if content['images']:
        message_content.append({
            "type": "text",
            "text": f"\n\n=== IMMAGINI ESTRATTE ({len(content['images'])} totali) ==="
        })
        
        for idx, img in enumerate(content['images']):
            message_content.append({
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/{img['format']};base64,{img['data']}"
                }
            })
            message_content.append({
                "type": "text",
                "text": f"[Immagine {idx + 1} - dalla pagina {img['page']}]"
            })
    
    message = HumanMessage(content=message_content)
    
    # 4. Invia la richiesta
    print("Inviando richiesta all'LLM...")
    response = llm.invoke([message])
    
    return response.content


def process_pdf_text_only(pdf_path: str, query: str, api_key: str = None, save_images: bool = False):
    """
    Versione più semplice: processa solo il testo (più economica).
    """
    content = extract_pdf_content(pdf_path, save_images=save_images)
    
    llm = ChatOpenAI(
        model="gpt-4o-mini",  # Più economico per solo testo
        api_key=api_key,
        temperature=0
    )
    
    message = HumanMessage(
        content=f"{query}\n\n=== CONTENUTO DEL PDF ===\n{content['text']}"
    )
    
    response = llm.invoke([message])
    return response.content


# Esempio di utilizzo
if __name__ == "__main__":
    # Configura il path del PDF
    pdf_path = r"C:\Users\x.hita\OneDrive - Reply\Workspace\Sisal\Test_Design\input\Figma\UX_UI App SEVV - Agile II.pdf"  # Modifica con il tuo path
    
    # Definisci la tua domanda/richiesta
    query = """
    Analizza questo documento e fornisci:
    1. Un riassunto del contenuto testuale
    2. Una descrizione delle immagini presenti
    3. I punti chiave principali
    """
    
    try:
        # OPZIONE 1: Con testo e immagini (più costoso)
        print("=== PROCESSAMENTO COMPLETO (testo + immagini) ===\n")
        result = process_pdf_with_langchain(
            pdf_path=pdf_path,
            query=query,
            # api_key="sk-..."  # Opzionale se usi variabile ambiente
        )
        
        print("\n=== RISPOSTA ===")
        print(result)
        
        # Mostra dove sono state salvate le immagini
        # if content.get('image_paths'):
        #     print(f"\n=== IMMAGINI SALVATE ({len(content['image_paths'])}) ===")
        #     for img_path in content['image_paths']:
        #         print(f"  - {img_path}")
        
        # OPZIONE 2: Solo testo (più economico e veloce)
        # print("=== PROCESSAMENTO SOLO TESTO ===\n")
        # result = process_pdf_text_only(
        #     pdf_path=pdf_path,
        #     query=query
        # )
        # print("\n=== RISPOSTA ===")
        # print(result)
        
    except Exception as e:
        print(f"Errore: {e}")
        import traceback
        traceback.print_exc()