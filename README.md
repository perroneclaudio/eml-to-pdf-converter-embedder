# eml-to-pdf-converter-embedder 📧➔(📄📎📎)

A Python tool to convert `.eml` files into readable PDF documents, designed to preserve the integrity of the original email by embedding all attachments and the source `.eml` file directly inside the PDF.

## Key Features

- **Attachment Embedding**: Attachments (PDFs, images, documents) are not only listed, but physically embedded into the PDF. They can be accessed from the “Attachments” panel of most PDF readers (Adobe Acrobat, Foxit, etc.).
- **Archiving**: The original .eml file is included as an attachment (can be excluded with ```--no-embed-eml```).
- **Batch Processing**: Converts hundreds of emails at once by simply providing a source folder.
- **Custom Output**: Supports TrueType fonts (`.ttf`), text size settings, and margin configuration.

## Installation

1. Clone the repository (or download the files):

   ```bash
   git clone https://github.com/perroneclaudio/eml-to-pdf-converter-embedder.git
   cd eml-to-pdf-converter-embedder
   ```
2. Create virtual environment (if needed) and activate (Python 3.10+ required):
   ```bash
   python3 -m venv venv
   source venv/bin/activate
   ```
   
3. Install dependencies:
   
   ```bash
   pip install -r requirements.txt
   ```

## Usage
Single conversion (PDF generated in the same folder):
   ```bash
   python3 eml_to_pdf.py messaggio.eml
   ```
Single conversion (custom output path):
   ```bash
   python3 eml_to_pdf.py messaggio.eml -o /folder/out.pdf
   # Windows: Use backslashes and quotes for paths with spaces
   ```
Batch conversion with a custom output folder:
   ```bash
   python3 eml_to_pdf.py ./cartella_eml --batch -o /folder_out
   # Windows: Use backslashes and quotes for paths with spaces

   ```
Conversion using a custom font:
   ```bash
   python3 eml_to_pdf.py messaggio.eml --font ./DejaVuSans.ttf
   # Windows: Use backslashes and quotes for paths with spaces

   ```
## Available Options:

  | Option             | Description                                                  |
  | ------------------ | ------------------------------------------------------------ |
  | `-o`, `--output`   | Output PDF path (or output folder when using batch mode)     |
  | `--batch`          | Processes all .eml files in the given folder                 |
  | `--font`           | Path to a .ttf font to use in the generated PDF              |
  | `--font-size`      | Text size (default: 10)                                      |
  | `--margins`        | Margins in mm (default: 20)                                  |
  | `--no-embed-eml`   | Skip original .eml embedding                                 |

## Roadmap

Next goal: Full PDF/A-3b compliance

## AI Disclosure

This project was developed with the support of AI tools (Gemini 3-Pro). The final code was reviewed and tested by the author (not NASA-level testing, though).

## Example in Foxit Reader

<img src="https://github.com/user-attachments/assets/c6ef7a35-4688-4d9f-baca-fa9995c9299e" alt="Example Image" style="max-width: 100%; height: auto;">
