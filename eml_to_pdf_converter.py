#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
eml-to-pdf-converter-embedder (Strict PDF/A-3b Compliant)
Target: Legal Archiving & Long-term Preservation
Features: Robust Parsing, Secure Embedding, Strict Metadata, No Warnings
Copyright (c) 2026 Claudio Perrone
Licensed under the MIT License.
"""

import argparse
import os
import re
import shutil
import tempfile
import datetime
import mimetypes
import sys
import email.utils
from pathlib import Path
from email import policy
from email.parser import BytesParser
from xml.sax.saxutils import escape as xml_escape

# --- Gestione Dipendenze ---
try:
    import pikepdf
except ImportError:
    sys.exit("ERRORE CRITICO: Libreria 'pikepdf' non trovata. Installa con: pip install pikepdf")

try:
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("ERRORE CRITICO: Libreria 'beautifulsoup4' non trovata. Installa con: pip install beautifulsoup4")

# Gestione opzionale MSG
try:
    import extract_msg
except ImportError:
    extract_msg = None

# Componenti ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily

# --- Costanti per PDF/A ---
XMP_TEMPLATE = """<?xpacket begin="" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
 <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
  <rdf:Description rdf:about=""
    xmlns:dc="http://purl.org/dc/elements/1.1/"
    xmlns:xmp="http://ns.adobe.com/xap/1.0/"
    xmlns:pdf="http://ns.adobe.com/pdf/1.3/"
    xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">
   <dc:format>application/pdf</dc:format>
   <dc:title><rdf:Alt><rdf:li xml:lang="x-default">{title}</rdf:li></rdf:Alt></dc:title>
   <dc:creator><rdf:Seq><rdf:li>{author}</rdf:li></rdf:Seq></dc:creator>
   <xmp:CreateDate>{created_date}</xmp:CreateDate>
   <xmp:ModifyDate>{mod_date}</xmp:ModifyDate>
   <xmp:MetadataDate>{mod_date}</xmp:MetadataDate>
   <xmp:CreatorTool>EML to PDF Converter Embedder</xmp:CreatorTool>
   <pdf:Producer>Pikepdf &amp; ReportLab</pdf:Producer>
   <pdf:Keywords>{keywords}</pdf:Keywords>
   <pdfaid:part>3</pdfaid:part>
   <pdfaid:conformance>B</pdfaid:conformance>
  </rdf:Description>
 </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>"""

MONTHS_IT = {
    1: "gennaio", 2: "febbraio", 3: "marzo", 4: "aprile",
    5: "maggio", 6: "giugno", 7: "luglio", 8: "agosto",
    9: "settembre", 10: "ottobre", 11: "novembre", 12: "dicembre"
}

# ---------------------------------------------------------------------
# Utilities: Date & Stringhe (FIXED FOR STRICT VALIDATION)
# ---------------------------------------------------------------------

def sanitize_filename(name: str, fallback: str = "file") -> str:
    """Pulisce i nomi dei file per compatibilità filesystem e PDF name trees."""
    name = (name or "").strip()
    name = re.sub(r"[^\w\-. ()\[\]]+", "_", name, flags=re.UNICODE)
    name = name.strip(" ._")
    return name[:250] if name else fallback

def get_xmp_date() -> str:
    """Costruisce manualmente la data XMP (YYYY-MM-DDThh:mm:ss+HH:MM)."""
    now = datetime.datetime.now().astimezone()
    offset_seconds = now.utcoffset().total_seconds()
    sign = '+' if offset_seconds >= 0 else '-'
    offset_seconds = abs(int(offset_seconds))
    hours, remainder = divmod(offset_seconds, 3600)
    minutes, _ = divmod(remainder, 60)
    
    date_part = now.strftime('%Y-%m-%dT%H:%M:%S')
    tz_part = f"{sign}{hours:02d}:{minutes:02d}"
    return f"{date_part}{tz_part}"

def get_pdf_date() -> str:
    """Costruisce manualmente la data PDF String (D:YYYYMMDDhhmmss+HH'mm')."""
    now = datetime.datetime.now().astimezone()
    offset_seconds = now.utcoffset().total_seconds()
    sign = '+' if offset_seconds >= 0 else '-'
    offset_seconds = abs(int(offset_seconds))
    hours, remainder = divmod(offset_seconds, 3600)
    minutes, _ = divmod(remainder, 60)
    
    date_part = now.strftime('%Y%m%d%H%M%S')
    tz_part = f"{sign}{hours:02d}'{minutes:02d}'"
    return f"D:{date_part}{tz_part}"

def format_date_italian(date_str: str) -> str:
    """Converte l'header Date in formato leggibile italiano."""
    if not date_str:
        return ""
    try:
        parsed = email.utils.parsedate_to_datetime(date_str)
        day = parsed.day
        month = MONTHS_IT.get(parsed.month, parsed.strftime("%B"))
        year = parsed.year
        hour = parsed.hour
        minute = parsed.minute
        return f"{day:02d} {month} {year}, {hour:02d}:{minute:02d}"
    except Exception:
        return date_str

# ---------------------------------------------------------------------
# Parsing & Content Extraction
# ---------------------------------------------------------------------

def html_to_text(html: str) -> str:
    """Converte HTML in testo strutturato."""
    if not html: return ""
    parser = "lxml"
    try:
        import lxml
    except ImportError:
        parser = "html.parser"
    
    try:
        soup = BeautifulSoup(html, parser)
        for tag in soup(["script", "style", "noscript", "header", "footer", "meta", "link"]):
            tag.decompose()
        for a in soup.find_all('a', href=True):
            if a.string:
                a.string = f"{a.string} ({a['href']})"
        return soup.get_text("\n").strip()
    except Exception as e:
        print(f"Warning: Errore parsing HTML ({e}), ritorno testo grezzo.", file=sys.stderr)
        return html

def extract_text_from_eml(msg) -> str:
    """Estrae il corpo del messaggio dando priorità all'HTML convertito."""
    body_plain = ""
    body_html = ""

    def safe_decode(part):
        try:
            return part.get_content()
        except Exception:
            payload = part.get_payload(decode=True)
            if not payload: return ""
            for enc in ['utf-8', 'windows-1252', 'iso-8859-1']:
                try: return payload.decode(enc)
                except UnicodeDecodeError: continue
            return payload.decode('utf-8', errors='replace')

    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_disposition() == "attachment": continue
            ctype = part.get_content_type()
            if ctype == "text/plain":
                body_plain += safe_decode(part)
            elif ctype == "text/html":
                body_html += html_to_text(safe_decode(part))
    else:
        ctype = msg.get_content_type()
        content = safe_decode(msg)
        if ctype == "text/html":
            return html_to_text(content)
        return content.strip()

    return body_html.strip() if body_html.strip() else body_plain.strip()

# ---------------------------------------------------------------------
# Generazione PDF Base (ReportLab)
# ---------------------------------------------------------------------

def create_pdf_from_data(pdf_path: Path, headers: dict, body_text: str, attachments: list, font_path: Path = None, font_bold_path: Path = None, font_size=10, margins=20):
    """Genera il layout visivo del PDF."""
    font_name = "Helvetica"
    
    if font_path and font_path.exists():
        try:
            pdfmetrics.registerFont(TTFont('CustomFont', str(font_path)))
            font_name = 'CustomFont'
            if font_bold_path and font_bold_path.exists():
                pdfmetrics.registerFont(TTFont('CustomFont-Bold', str(font_bold_path)))
                registerFontFamily('CustomFont', normal='CustomFont', bold='CustomFont-Bold')
            else:
                registerFontFamily('CustomFont', normal='CustomFont', bold='CustomFont')
        except Exception as e:
            print(f"Warning: Errore font custom ({e}). Uso Helvetica.")
    
    m = margins * mm
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=m, leftMargin=m, topMargin=m, bottomMargin=m)
    styles = getSampleStyleSheet()
    
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName=font_name, fontSize=font_size, leading=font_size*1.2, spaceAfter=6)
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontName=font_name)
    header_lbl_style = ParagraphStyle('HeaderLbl', parent=styles['Normal'], fontName=font_name, fontSize=font_size, spaceAfter=2)
    
    story = [Paragraph("Archivio Email", title_style), Spacer(1, 10*mm)]
    
    header_order = ["From", "To", "Cc", "Date", "Subject"]
    for k in header_order:
        v = headers.get(k, "")
        if v:
            if k == "Date": v = format_date_italian(v)
            clean_v = xml_escape(str(v))
            txt = f"<b>{k}:</b> {clean_v}"
            story.append(Paragraph(txt, header_lbl_style))
            
    story.append(Spacer(1, 10*mm))
    story.append(HRFlowable(width="100%", thickness=0.5, color="grey"))
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("<b>Testo del messaggio:</b>", body_style))
    
    if body_text:
        body_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', body_text)
        for line in body_text.splitlines():
            if not line.strip():
                story.append(Spacer(1, 2*mm))
                continue
            clean_line = xml_escape(line)
            story.append(Paragraph(clean_line, body_style))
            
    if attachments:
        story.append(Spacer(1, 10*mm))
        story.append(HRFlowable(width="100%", thickness=1, color="black"))
        story.append(Spacer(1, 5*mm))
        story.append(Paragraph("<b>Elenco allegati incorporati nel file:</b>", body_style))
        for f in attachments:
            try: size_kb = os.path.getsize(f) // 1024
            except OSError: size_kb = 0
            clean_fname = xml_escape(f.name)
            story.append(Paragraph(f"• {clean_fname} ({size_kb} KB)", body_style))
            
    doc.build(story)

# ---------------------------------------------------------------------
# Pipeline PDF/A (Pikepdf)
# ---------------------------------------------------------------------

def generate_pdfa_metadata(pdf_path: Path) -> bytes:
    """Genera XML XMP valido con formattazione rigorosa della data."""
    xmp_date = get_xmp_date()
    safe_title = xml_escape(os.path.basename(pdf_path))
    safe_author = xml_escape("EML to PDF Converter")
    
    meta_xml = XMP_TEMPLATE.format(
        title=safe_title,
        author=safe_author,
        created_date=xmp_date,
        mod_date=xmp_date,
        keywords="email, archive, pdf-a, legal"
    )
    return meta_xml.encode('utf-8')

def finalize_pdf_with_attachments(pdf_in: Path, pdf_out: Path, files_to_attach: list, icc_profile_path: Path = None):
    """Step finale: Embedding file, iniezione ICC, XMP, PDF/A ID."""
    with pikepdf.open(pdf_in) as pdf:
        
        embedded_files_data = []
        pdf_date_str = get_pdf_date()
        
        for f in files_to_attach:
            if not f.exists(): continue
            try: file_data = f.read_bytes()
            except IOError: continue
            
            mime_type, _ = mimetypes.guess_type(f.name)
            if not mime_type: mime_type = "application/octet-stream"
            
            ef_stream = pdf.make_stream(file_data)
            ef_stream.Type = pikepdf.Name("/EmbeddedFile")
            ef_stream.Subtype = pikepdf.Name("/" + mime_type.replace("/", "#2f"))
            ef_stream.Params = pikepdf.Dictionary({
                "/Size": len(file_data),
                "/ModDate": pikepdf.String(pdf_date_str),
                "/CreationDate": pikepdf.String(pdf_date_str)
            })
            
            safe_fname = sanitize_filename(f.name)
            fs = pikepdf.Dictionary({
                "/Type": pikepdf.Name("/Filespec"),
                "/F": pikepdf.String(safe_fname),
                "/UF": pikepdf.String(safe_fname),
                "/EF": pikepdf.Dictionary({"/F": ef_stream}),
                "/AFRelationship": pikepdf.Name("/Source") if f.suffix.lower() in [".eml", ".msg"] else pikepdf.Name("/Data"),
                "/Desc": pikepdf.String(safe_fname)
            })
            embedded_files_data.append((safe_fname, pdf.make_indirect(fs)))
            
        if "/Names" not in pdf.Root: pdf.Root.Names = pikepdf.Dictionary()
        
        name_array = pikepdf.Array()
        af_array = pikepdf.Array()
        
        for fname, fs_ref in sorted(embedded_files_data, key=lambda x: x[0]):
            name_array.append(pikepdf.String(fname))
            name_array.append(fs_ref)
            af_array.append(fs_ref) 
            
        pdf.Root.Names.EmbeddedFiles = pikepdf.Dictionary({"/Names": name_array})
        pdf.Root["/AF"] = af_array

        metadata_stm = pdf.make_stream(generate_pdfa_metadata(pdf_in))
        metadata_stm.Type = pikepdf.Name("/Metadata")
        metadata_stm.Subtype = pikepdf.Name("/XML")
        pdf.Root.Metadata = metadata_stm

        if icc_profile_path and icc_profile_path.exists():
            icc_stream = pdf.make_stream(icc_profile_path.read_bytes())
            icc_stream.N = 3
            icc_stream.Alternate = pikepdf.Name("/DeviceRGB")
            
            output_intent = pikepdf.Dictionary({
                "/Type": pikepdf.Name("/OutputIntent"),
                "/S": pikepdf.Name("/GTS_PDFA1"),
                "/OutputConditionIdentifier": pikepdf.String("sRGB"),
                "/DestOutputProfile": icc_stream,
                "/Info": pikepdf.String("sRGB IEC61966-2.1")
            })
            pdf.Root.OutputIntents = pikepdf.Array([output_intent])

        # FIX: Disabilitiamo l'auto-update del Producer da parte di Pikepdf per evitare il warning
        with pdf.open_metadata(set_pikepdf_as_editor=False) as meta:
            meta["dc:title"] = os.path.basename(pdf_in)
            meta["pdf:Producer"] = "EML to PDF/A Converter"

        pdf.save(pdf_out, fix_metadata_version=True)

# ---------------------------------------------------------------------
# Logica Principale
# ---------------------------------------------------------------------

def process_file(input_path: Path, out_pdf: Path, font_p: Path, font_bold_p: Path, font_size: int, margins: int, include_inline: bool, embed_orig: bool, icc_path: Path):
    ext = input_path.suffix.lower()
    
    with tempfile.TemporaryDirectory(prefix="eml2pdf_proc_") as tmp_dir:
        tmp_path = Path(tmp_dir)
        att_dir = tmp_path / "attachments"
        att_dir.mkdir()
        
        attachments_paths = []
        headers = {}
        body = ""

        if ext == ".eml":
            try:
                msg = BytesParser(policy=policy.default).parsebytes(input_path.read_bytes())
            except Exception as e:
                raise ValueError(f"File EML illeggibile: {e}")

            headers = {k: str(msg.get(k, "")) for k in ["From", "To", "Cc", "Date", "Subject"]}
            
            for i, part in enumerate(msg.walk(), 1):
                filename = part.get_filename()
                disp = (part.get_content_disposition() or "").lower()
                is_inline = (disp == "inline") or (part.get("Content-ID") is not None)
                if is_inline and not include_inline: continue
                if not filename and disp != "attachment": continue
                
                data = part.get_payload(decode=True)
                if not data: continue
                
                final_name = filename if filename else f"attachment_{i}.bin"
                safe_name = sanitize_filename(final_name, f"att_{i}")
                dest = att_dir / safe_name
                counter = 1
                while dest.exists():
                    dest = att_dir / f"{counter}_{safe_name}"
                    counter += 1
                dest.write_bytes(data)
                attachments_paths.append(dest)
            
            body = extract_text_from_eml(msg)

        elif ext == ".msg":
            if not extract_msg: raise ImportError("Libreria 'extract-msg' mancante.")
            try:
                msg = extract_msg.Message(str(input_path))
                headers = {"From": getattr(msg, 'sender', ''), "To": getattr(msg, 'to', ''), "Cc": getattr(msg, 'cc', ''), "Date": getattr(msg, 'date', ''), "Subject": getattr(msg, 'subject', '')}
                raw_body = msg.htmlBody
                if raw_body:
                    if isinstance(raw_body, bytes): raw_body = raw_body.decode('utf-8', errors='ignore')
                    body = html_to_text(raw_body)
                else:
                    body = msg.body or ""

                for i, att in enumerate(msg.attachments):
                    if not include_inline and getattr(att, 'cid', None): continue
                    fname = getattr(att, 'longFilename', None) or getattr(att, 'shortFilename', None) or f"att_{i}"
                    safe_name = sanitize_filename(fname)
                    dest = att_dir / safe_name
                    counter = 1
                    while dest.exists():
                        dest = att_dir / f"{counter}_{safe_name}"
                        counter += 1
                    dest.write_bytes(att.data)
                    attachments_paths.append(dest)
                msg.close()
            except Exception as e: raise ValueError(f"Errore MSG: {e}")

        all_to_embed = []
        if embed_orig:
            orig_copy = tmp_path / sanitize_filename(input_path.name, f"source{ext}")
            shutil.copy2(input_path, orig_copy)
            all_to_embed.append(orig_copy)
        
        all_to_embed.extend(attachments_paths)
        intermediate_pdf = tmp_path / "layout_temp.pdf"
        create_pdf_from_data(intermediate_pdf, headers, body, all_to_embed, font_p, font_bold_p, font_size, margins)
        finalize_pdf_with_attachments(intermediate_pdf, out_pdf, all_to_embed, icc_path)

def main():
    ap = argparse.ArgumentParser(description="EML/MSG to PDF/A-3b Converter (Legal Archive Ready)")
    ap.add_argument("input_path", help="Percorso file .eml, .msg o cartella")
    ap.add_argument("-o", "--output", help="Percorso PDF o cartella di destinazione")
    ap.add_argument("--font", help="Percorso font .ttf (Regular)")
    ap.add_argument("--font-bold", help="Percorso font .ttf (Bold)")
    ap.add_argument("--icc", required=True, help="Percorso profilo colore .icc")
    ap.add_argument("--font-size", type=int, default=10)
    ap.add_argument("--margins", type=int, default=20)
    ap.add_argument("--batch", action="store_true", help="Elabora cartella intera")
    ap.add_argument("--no-embed-orig", action="store_true", help="Escludi file sorgente dall'embedding")
    ap.add_argument("--exclude-inline", action="store_true", help="Escludi oggetti inline come allegati")
    ap.add_argument("--validate", action="store_true", help="Ignorato (per compatibilità)")
    
    args = ap.parse_args()
    ip = Path(args.input_path).resolve()
    icc = Path(args.icc).resolve()
    fp = Path(args.font).resolve() if args.font else None
    fbp = Path(args.font_bold).resolve() if args.font_bold else None

    if not icc.exists(): sys.exit(f"ERRORE: Profilo ICC non trovato in {icc}")
    
    files = []
    if args.batch:
        if not ip.is_dir(): sys.exit("Errore: input batch deve essere cartella")
        files.extend(ip.glob("*.eml"))
        files.extend(ip.glob("*.msg"))
        if not files: sys.exit("Nessun file trovato")
    else:
        if not ip.is_file(): sys.exit("Errore: file input non trovato")
        files.append(ip)

    out_base = Path(args.output).resolve() if args.output else (ip if args.batch else ip.parent)
    if args.batch: out_base.mkdir(parents=True, exist_ok=True)
    elif out_base.suffix.lower() == ".pdf": out_base.parent.mkdir(parents=True, exist_ok=True)

    success, fail = 0, 0
    print(f"Elaborazione {len(files)} files...")
    for f in files:
        target = out_base / (f.stem + ".pdf") if args.batch else (out_base if str(out_base).lower().endswith(".pdf") else out_base / (f.stem + ".pdf"))
        print(f" -> {f.name}...", end=" ", flush=True)
        try:
            process_file(f, target, fp, fbp, args.font_size, args.margins, not args.exclude_inline, not args.no_embed_orig, icc)
            print("OK")
            success += 1
        except Exception as e:
            print(f"FAIL: {e}")
            fail += 1
    sys.exit(1 if fail > 0 else 0)

if __name__ == "__main__":
    main()