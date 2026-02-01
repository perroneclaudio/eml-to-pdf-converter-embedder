

# eml-to-pdf-converter-embedder 📧➔📄

Uno strumento Python per convertire file `.eml` in documenti PDF leggibili, progettato per preservare l'integrità dell'email originale incorporando tutti gli allegati e il file .eml sorgente direttamente all'interno del PDF.

##  Caratteristiche principali

- **Attachment Embedding**: Gli allegati (PDF, immagini, documenti) non vengono solo citati, ma inseriti fisicamente nel PDF. Sono consultabili dal pannello "Allegati" di qualsiasi lettore PDF (Adobe Acrobat, Foxit ecc.).
- **Archiviazione Sicura**: Il file `.eml` originale viene sempre incluso come allegato all'interno del PDF.
- **Batch Processing**: Converte centinaia di email in un colpo solo semplicemente indicando una cartella di origine.
- **Output Personalizzato**: Supporto per font TrueType (.ttf), gestione della dimensione del testo e dei margini.

##  Installazione

1. Clona il repository o scarica i file:
   ```bash
   git clone https://github.com/perroneclaudio/eml-to-pdf-converter-embedder.git
   cd eml-to-pdf-converter-embedder
   ```
2. Installa le dipendenze:

   ```bash
    pip install -r requirements.txt
    ```
##  Utilizzo

Conversione singola (pdf generato nella stessa cartella):
   ```bash
  python eml_to_pdf.py messaggio.eml
   ```
Conversione singola con specifica percorso di output:

  ```bash
  python eml_to_pdf.py messaggio.eml -o /folder/out.pdf
  ```
Conversione batch con specifica percorso di output:

  ```bash
  python eml_to_pdf.py ./cartella_eml --batch -o /folder_out
  ```
Conversione con font personalizzato:

  ```bash
  python eml_to_pdf.py messaggio.eml --font ./DejaVuSans.ttf
  ```
## Opzioni disponibili:

  | Opzione            | Descrizione                                                  |
  | ------------------ | ------------------------------------------------------------ |
  | `-o`, `--output`   | Percorso PDF di output (o cartella output in modalità batch) |
  | `--batch`          | Elabora tutti i file `.eml` nella cartella indicata          |
  | `--font`           | Percorso font `.ttf` da usare nel PDF                        |
  | `--font-size`      | Dimensione testo (default: 10)                               |
  | `--margins`        | Margini in mm (default: 20)                                  |


