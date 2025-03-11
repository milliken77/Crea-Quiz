# Generatore di Quiz

Questo progetto è un generatore di quiz che consente di creare versioni multiple di un quiz a partire da un insieme di domande e risposte formattate in modo specifico. Il programma produce due output principali:

- **Documento DOCX**: Un file Word che contiene le domande del quiz formattate in due colonne, complete di intestazione, dati di compilazione (nome, cognome, data) e una griglia di correzione.
- **File Excel (XLSX)**: Un foglio di calcolo che offre una scheda per l'inserimento delle risposte degli studenti e una scheda (nascosta) con le risposte corrette. La scheda per gli studenti applica la formattazione condizionale per evidenziare in verde le risposte corrette.

Il progetto è stato sviluppato in due versioni principali: una versione originale in Python e una versione client-side in HTML/JavaScript.

---

## Versione Originale in Python

La versione Python del generatore prevede i seguenti elementi:

- **Input Testuale**:  
  Il programma accetta un file o un input testuale in cui le domande sono indicate con un `#` all'inizio e le risposte sono elencate successivamente. La risposta corretta è identificata da un asterisco `*` davanti al testo.

- **Logica di Shuffling**:  
  Le domande e le rispettive risposte vengono mescolate in modo casuale. Questo permette di creare versioni differenti dello stesso quiz, garantendo varietà e prevenendo copie identiche.

- **Generazione del Documento Word**:  
  Utilizzando librerie come *python-docx* (o simili), il programma crea un file DOCX che include:  
  - Un'intestazione con il titolo del quiz in "dimensione 13" (in termini di formattazione tipografica, ad esempio, 13pt).  
  - La scritta "Compito n° ..." in formato normale, seguita dai campi per nome, cognome e data.  
  - Le domande vengono disposte in due colonne e, **solamente prima di ogni domanda**, viene inserita una riga vuota.  
  - Una griglia di correzione viene inserita subito dopo l'ultima domanda, con il numero della domanda e lo spazio per la risposta corretta.

- **Generazione del File Excel**:  
  Utilizzando librerie come *openpyxl*, il programma genera un file XLSX con due schede:  
  - Una scheda con le risposte corrette, che viene nascosta.  
  - Una scheda per le risposte degli studenti, in cui sono presenti celle aggiuntive di controllo (inserite in colonne successive) che, tramite formule, verificano se la risposta immessa corrisponde a quella corretta.  
  - La formattazione condizionale applica uno sfondo verde alle celle delle risposte degli studenti quando la risposta è corretta.

---

## Versione Client-Side in HTML/JavaScript

La versione sviluppata in HTML e JavaScript è completamente eseguibile nel browser e si basa su librerie JavaScript per generare i file di output:

- **Interfaccia Web**:  
  Un form permette all'utente di:  
  - Inserire il testo delle domande (con il formato `#domanda` per le domande e `*risposta` per indicare la risposta corretta).  
  - Impostare il numero di versioni e il numero di domande per ciascun quiz.  
  - Specificare i nomi dei file di output (ad esempio, `quiz.docx` e `correttore.xlsx`).

- **Librerie Utilizzate**:  
  - **docx.js**: Per generare il file DOCX direttamente nel browser.  
  - **FileSaver.js**: Per consentire il salvataggio e il download dei file generati.  
  - **exceljs**: Per generare il file XLSX e applicare la formattazione condizionale (utilizzando la proprietà `bgColor` per impostare lo sfondo verde nelle celle con risposte corrette).

- **Funzionalità Principali**:  
  - La logica di shuffling (mescolare domande e risposte) è stata trasposta in JavaScript.  
  - Nel documento DOCX, il titolo iniziale (proveniente dal campo "intestazione") viene visualizzato in "dimensione 13" (ad esempio, con size 26 in docx.js) mentre la scritta "Compito n° ..." viene visualizzata in formato normale.  
  - Viene inserita una riga vuota **solo prima di ogni domanda** (quelle che corrispondono al pattern "numero - ...") per una migliore leggibilità.  
  - Nel file Excel, il foglio "Risposte_Corrette" viene creato e forzato a essere nascosto, mentre il foglio "Risposte Studenti" contiene celle di controllo aggiuntive, le quali determinano tramite formule se la risposta inserita corrisponde a quella corretta. Quando questo accade, la cella viene evidenziata con uno sfondo verde.

- **Download**:  
  Al termine della generazione, i file DOCX e XLSX vengono scaricati direttamente sul computer dell'utente.



## Conclusioni

Questo generatore di quiz offre una soluzione completa per creare quiz personalizzati con output professionali in formato Word e Excel.  
La versione Python è ideale per integrazioni lato script o per ambienti in cui si preferisce l'elaborazione server-side tramite Python, mentre la versione HTML/JavaScript permette di eseguire tutto il processo nel browser, senza necessità di installazioni lato server.
