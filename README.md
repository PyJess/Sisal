Istruzioni per testare i file:

-per adesso si possono processare file EXCEL e file DOCX
 
CONTROLLO SINTATTICO:
Accedere alla cartella Processing, aprire file "controllo_sintattico.py", in main definire il path del file excel a cui vogliamo fare il controllo sintattico. Una volta definito il path in input_path si può runnare il codice direttamente.
 
 
TEST DESIGN:
Accedere alla cartella Processing, aprire file "test_design.py", in main definire il path di "input_path" (il file da cui vogliamo generate test cases), inoltre definire "rag_path" il file che vogliamo usare par fare rag. Si può runnare il codice direttamente.
 
 
COPERTURA APPLICATIVI:
Accedere alla cartella Processing, aprire file "copertura_applicativi.py", definire il path del file excel e del corrispettivo file di requirement. L'agente partendo dal requirement andrà ad inidividuare tutti i possibili applicativi, una volta individuati (possibile vederli dal terminal) va a cercare sul file excel se sono presenti, se sono presenti significa che c'è almeno un test case che tratta di quell'applicativo, altrimenti genera test case inerente a questo applicativo.
Definire "input_path", "rag_path" ed "excel_path", si può runnare il codice direttamente.
 
COPERTURA REQUISITI: 
Accedere alla cartella Processing, aprire file "copertura_requisiti.py", definire il path del file excel nella variabile "test_case_input" , ripetere lo stesso procedimento per il file docx nella variabile "documents_word". A partire da requirement viene eseguito un confronto dei codici "_polarion" presente nei testcase con quelli presenti nel file docx.
Qualora ci fossero dei requirement mancanti l'agente IA si attiva e procederà a generare i testcase a partire dalla documentazione fornita. L'output si trovera dentro la "input" assieme all'excel originale con lo stesso nome+_feedbackAI_requisiti.xlsx

COPERTURA PROGETTAZIONE:
Accedere alla cartella Processing, aprire file "copertura_proggetazione.py", definire il path del file excel nella variabile "excel_path" , ripetere lo stesso procedimento per il file docx nella variabile "input_path".  A partire da requirement viene eseguito un confronto dei codici "_polarion" presente nei testcase con quelli presenti nel file docx.
Qualora ci fossero dei requirement mancanti l'agente IA si attiva e procederà a generare i testcase o a modificarli per colmare delle lacune a partire dalla documentazione fornita. L'output si trovera dentro la "input" assieme all'excel originale con lo stesso nome+_feedbackAI_testcase_progettazione.xlsx

