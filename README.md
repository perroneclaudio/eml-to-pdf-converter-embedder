# EML(and MSG)-to-pdf-converter-embedder 📧➔(📄📎)

A Python tool to convert .EML and .MSG files into readable PDF documents, designed to preserve the integrity of the original email by embedding all attachments and the source `.eml` file directly inside the PDF.

## Key Features

- **Attachment Embedding**: Attachments (PDFs, images, documents) are not only listed, but physically embedded into the PDF. They can be accessed from the “Attachments” panel of most PDF readers (Adobe Acrobat, Foxit, etc.).
- **PDF/A-3b Compliance & Legal Archiving**: Generates ISO 19005-3 compliant documents suitable for long-term preservation. Enforces strict XMP metadata, timezone integrity, and ICC color profiles to ensure validation with tools like **veraPDF**, the **industry standard** for **PDF/A** validation.
- **Source Archiving**: The original .eml/.msg file is included as an attachment (can be excluded with `--no-embed-orig`).
- **Batch Processing**: Converts hundreds of emails at once by simply providing a source folder.
- **Custom Output**: Supports TrueType fonts (`.ttf`) embedding, text size settings, and margin configuration.

## Installation

1. Clone the repository (or download the files):

   ```bash
   git clone https://github.com/perroneclaudio/eml-to-pdf-converter-embedder.git
   cd eml-to-pdf-converter-embedder
   ```
2. Create virtual environment (if needed) and activate (Python 3.10+ required):
   ```bash
   python -m venv venv
   source venv/bin/activate
   ```
   
3. Install dependencies:
   
   ```bash
   pip install -r requirements.txt
   ```

## To generate a fully compliant PDF/A-3b file, you MUST provide a regular font, a bold font, and an ICC profile via the command-line options. Sample files can be found in the assets folder. You MUST use an ICC profile with Device Class Monitor (mntr) or Printer (prtr).

## Usage

Single conversion (PDF generated in the same folder):
   ```bash
   python eml_to_pdf.py messaggio.eml --icc "./assets/srgb.icc" --font "./assets/DejaVuSans.ttf" --font-bold "./assets/DejaVuSans-Bold.ttf"
   ```
Single conversion (custom output path):
   ```bash
   python eml_to_pdf.py messaggio.eml -o /folder/out.pdf --icc "./assets/srgb.icc" --font "./assets/DejaVuSans.ttf" --font-bold "./assets/DejaVuSans-Bold.ttf"
   # Windows: Use backslashes and quotes for paths with spaces
   ```
Batch conversion with a custom output folder:
   ```bash
   python eml_to_pdf.py ./cartella_eml --batch -o /folder_out --icc "./assets/srgb.icc" --font "./assets/DejaVuSans.ttf" --font-bold "./assets/DejaVuSans-Bold.ttf"
   # Windows: Use backslashes and quotes for paths with spaces

 
## Available Options:

  | Option             | Description                                                  |
  | ------------------ | ------------------------------------------------------------ |
  | `-o`, `--output`   | Output PDF path (or output folder when using batch mode)     |
  | `--batch`          | Processes all .eml files in the given folder                 |
  | `--font`           | Path to a .ttf REGULAR font to use in the generated PDF      |
  | `--font-bold`      | Path to a .ttf BOLD    font to use in the generated PDF      |
  | `--icc      `      | Path to an ICC color profile                                 |
  | `--font-size`      | Text size (default: 10)                                      |
  | `--margins`        | Margins in mm (default: 20)                                  |
  | `--no-embed-orig`  | Skip original .eml embedding                                 |
  | `--exclude-inline` | Exclude inline objects as attachments                        |

## AI Disclosure

This project was developed with the support of AI tools (Gemini 3-Pro). The final code was reviewed and tested by the author (not NASA-level testing, though).

## Example in Foxit Reader

<img src="https://github.com/user-attachments/assets/c6ef7a35-4688-4d9f-baca-fa9995c9299e" alt="Example Image" style="max-width: 100%; height: auto;">
