# eml-to-pdf-converter-embedder 📧➔📄

A Python tool to convert `.eml` files into readable PDF documents, designed to preserve the integrity of the original email by embedding all attachments and the source `.eml` file directly inside the PDF.

## Key Features

- **Attachment Embedding**: Attachments (PDFs, images, documents) are not only listed, but physically embedded into the PDF. They can be accessed from the “Attachments” panel of most PDF readers (Adobe Acrobat, Foxit, etc.).
- **Archiving**: The original `.eml` file is always included as an attachment inside the PDF.
- **Batch Processing**: Converts hundreds of emails at once by simply providing a source folder.
- **Custom Output**: Supports TrueType fonts (`.ttf`), text size settings, and margin configuration.

## Installation

1. Clone the repository (or download the files):

   ```bash
   git clone https://github.com/perroneclaudio/eml-to-pdf-converter-embedder.git
   
   cd eml-to-pdf-converter-embedder
   ```
2. Install dependencies (Python 3.10+ required)
   
   ```bash
   pip install -r requirements.txt
   ```
## Usage
Single conversion (PDF generated in the same folder):
   ```bash
   python eml_to_pdf.py messaggio.eml
   ```
Single conversion (custom output path):
   ```bash
   python eml_to_pdf.py messaggio.eml -o /folder/out.pdf
   ```
Batch conversion with a custom output folder:
   ```bash
   python eml_to_pdf.py ./cartella_eml --batch -o /folder_out
   ```
Conversion using a custom font:
   ```bash
   python eml_to_pdf.py messaggio.eml --font ./DejaVuSans.ttf
   ```
## Available Options:

  | Option             | Description                                                  |
  | ------------------ | ------------------------------------------------------------ |
  | `-o`, `--output`   | Output PDF path (or output folder when using batch mode)     |
  | `--batch`          | Processes all .eml files in the given folder                 |
  | `--font`           | Path to a .ttf font to use in the generated PDF              |
  | `--font-size`      | Text size (default: 10)                                      |
  | `--margins`        | Margins in mm (default: 20)                                  |

## AI Disclosure

This project was developed with the support of AI tools (Gemini 3-Pro). The final code was reviewed and tested by the author (not NASA-level testing, though).

<img width="579" height="458" alt="image" src="https://github.com/user-attachments/assets/c6ef7a35-4688-4d9f-baca-fa9995c9299e" />

