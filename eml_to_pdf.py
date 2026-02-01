#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
eml-to-pdf-converter-embedder
Copyright (c) 2026 Claudio Perrone
Licensed under the MIT License.
See LICENSE file in the project root for full license information.
"""

import argparse
import os
import re
import shutil
import tempfile
from pathlib import Path
from email import policy
from email.parser import BytesParser

import pikepdf
from bs4 import BeautifulSoup

# Componenti ReportLab
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def sanitize_filename(name: str, fallback: str = "file") -> str:
    """Pulisce i nomi dei file per evitare errori di sistema o nel PDF."""
    name = (name or "").strip()
    name = re.sub(r"[^\w\-. ()\[\]]+", "_", name, flags=re.UNICODE)
    name = name.strip(" ._")
    return name if name else fallback

def html_to_text(html: str) -> str:
    """Converte HTML in testo semplice usando il miglior parser disponibile."""
    parser = "lxml"
    try:
        import lxml
    except ImportError:
        parser = "html.parser"
    
    soup = BeautifulSoup(html or "", parser)
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()
    return soup.get_text("\n").strip()

def extract_text_from_message(msg) -> str:
    """
    Estrae il corpo del messaggio preferendo l'HTML per evitare duplicati 
    tipici dei messaggi multipart/alternative.
    """
    body_plain = ""
    body_html = ""

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = (part.get_content_disposition() or "").lower()
            if disp == "attachment":
                continue
                
            if ctype == "text/plain":
                try:
                    content = part.get_content()
                except:
                    content = part.get_payload(decode=True).decode(errors='replace')
                body_plain += content
            elif ctype == "text/html":
                try:
                    content = part.get_content()
                except:
                    content = part.get_payload(decode=True).decode(errors='replace')
                body_html += html_to_text(content)
    else:
        ctype = msg.get_content_type()
        try:
            content = msg.get_content()
        except:
            content = msg.get_payload(decode=True).decode(errors='replace')
            
        if ctype == "text/html":
            return html_to_text(content)
        return content.strip()

    return (body_html if body_html.strip() else body_plain).strip()

def extract_attachments(eml_msg, out_dir: Path, include_inline: bool = False):
    """Estrae allegati reali filtrando (opzionalmente) firme e loghi."""
    out_dir.mkdir(parents=True, exist_ok=True)
    saved = []
    for i, part in enumerate(eml_msg.walk(), 1):
        filename = part.get_filename()
        disp = (part.get_content_disposition() or "").lower()
        
        if not filename or (disp != "attachment" and not include_inline):
            continue
            
        data = part.get_payload(decode=True)
        if not data: continue
        
        base_name = sanitize_filename(filename, fallback=f"attachment_{i}")
        safe_name = base_name
        if (out_dir / safe_name).exists():
            safe_name = f"{i}_{base_name}"
            
        dest = out_dir / safe_name
        dest.write_bytes(data)
        saved.append(dest)
    return saved

def create_pdf_from_email(pdf_path: Path, headers: dict, body_text: str, attachments: list, font_path: Path = None, font_size=10, margins=20):
    """Genera la visualizzazione PDF del testo dell'email."""
    font_name = "Helvetica"
    if font_path and font_path.exists():
        try:
            pdfmetrics.registerFont(TTFont('CustomFont', str(font_path)))
            font_name = 'CustomFont'
        except: pass
        
    m = margins * mm
    doc = SimpleDocTemplate(str(pdf_path), pagesize=A4, rightMargin=m, leftMargin=m, topMargin=m, bottomMargin=m)
    styles = getSampleStyleSheet()
    body_style = ParagraphStyle('BodyStyle', parent=styles['Normal'], fontName=font_name, fontSize=font_size, leading=font_size*1.2, spaceAfter=6)
    title_style = ParagraphStyle('TitleStyle', parent=styles['Title'], fontName=font_name)
    
    story = [Paragraph("Archivio Email", title_style), Spacer(1, 10*mm)]
    for k in ["From", "To", "Cc", "Date", "Subject"]:
        v = headers.get(k, "")
        if v:
            clean_v = str(v).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(f"<b>{k}:</b> {clean_v}", body_style))
            
    story.append(Spacer(1, 10*mm))
    story.append(Paragraph("<b>Testo del messaggio:</b>", body_style))
    
    if body_text:
        for line in body_text.splitlines():
            clean_line = line.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
            story.append(Paragraph(clean_line or " ", body_style))
            
    if attachments:
        story.append(Spacer(1, 10*mm))
        story.append(HRFlowable(width="100%", thickness=1, color="black"))
        story.append(Spacer(1, 5*mm))
        story.append(Paragraph("<b>Elenco allegati incorporati nel file:</b>", body_style))
        for f in attachments:
            size_kb = os.path.getsize(f) // 1024
            story.append(Paragraph(f"• {f.name} ({size_kb} KB)", body_style))
    doc.build(story)

def finalize_pdf_with_attachments(pdf_in: Path, pdf_out: Path, files_to_attach: list[Path]):
    """Incorpora i file nel PDF come EmbeddedFiles per massima compatibilità."""
    with pikepdf.open(pdf_in) as pdf:
        embedded_files = []
        for f in files_to_attach:
            if not f.exists(): continue
            file_data = f.read_bytes()
            
            ef_stream = pdf.make_stream(file_data)
            ef_stream.Type = pikepdf.Name("/EmbeddedFile")
            ef_stream.Params = pikepdf.Dictionary({"/Size": len(file_data)})
            
            fs = pikepdf.Dictionary({
                "/Type": pikepdf.Name("/Filespec"),
                "/F": pikepdf.String(f.name),
                "/UF": pikepdf.String(f.name),
                "/EF": pikepdf.Dictionary({"/F": ef_stream}),
                "/AFRelationship": pikepdf.Name("/Source") if f.suffix.lower() == ".eml" else pikepdf.Name("/Data")
            })
            embedded_files.append(fs)
            
        if "/Names" not in pdf.Root: pdf.Root.Names = pikepdf.Dictionary()
        name_array = pikepdf.Array()
        for fs in sorted(embedded_files, key=lambda x: str(x["/F"])):
            name_array.append(pikepdf.String(str(fs["/F"]))); name_array.append(fs)
            
        pdf.Root.Names.EmbeddedFiles = pikepdf.Dictionary({"/Names": name_array})
        pdf.Root["/AF"] = pikepdf.Array(embedded_files)
        pdf.save(pdf_out)

def process_single_eml(eml_path: Path, out_pdf: Path, font_p: Path, font_size: int, margins: int, include_inline: bool):
    """Esegue l'intero processo di conversione per un singolo file .eml."""
    with tempfile.TemporaryDirectory(prefix="eml_conv_") as tmp_dir_name:
        tmpdir = Path(tmp_dir_name)
        msg = BytesParser(policy=policy.default).parsebytes(eml_path.read_bytes())
        headers = {k: str(msg.get(k, "")) for k in ["From", "To", "Cc", "Date", "Subject"]}
        body = extract_text_from_message(msg)
        
        attachments = extract_attachments(msg, tmpdir / "files", include_inline)
        
        # Aggiungiamo sempre l'EML originale come allegato "ancora"
        eml_copy = tmpdir / sanitize_filename(eml_path.name, "original.eml")
        shutil.copy2(eml_path, eml_copy)
        all_to_embed = [eml_copy] + attachments
        
        intermediate_pdf = tmpdir / "temp.pdf"
        create_pdf_from_email(intermediate_pdf, headers, body, all_to_embed, font_p, font_size, margins)
        finalize_pdf_with_attachments(intermediate_pdf, out_pdf, all_to_embed)

def main():
    ap = argparse.ArgumentParser(description="Convertitore EML to PDF professionale.")
    ap.add_argument("input_path", help="Percorso file .eml o cartella")
    ap.add_argument("-o", "--output", help="Percorso PDF o cartella di destinazione")
    ap.add_argument("--font", help="Percorso font .ttf")
    ap.add_argument("--font-size", type=int, default=10)
    ap.add_argument("--margins", type=int, default=20)
    ap.add_argument("--batch", action="store_true", help="Elabora tutti i file nella cartella")
    ap.add_argument("--include-inline", action="store_true", help="Includi anche immagini delle firme")
    args = ap.parse_args()

    ip = Path(args.input_path).resolve()
    fp = Path(args.font).resolve() if args.font else None

    if args.batch:
        if not ip.is_dir():
            print(f"Errore: {ip} non è una cartella.")
            return
        emls = list(ip.glob("*.eml"))
        if not emls:
            print("Nessun file .eml trovato.")
            return
        
        od = Path(args.output).resolve() if args.output else ip
        od.mkdir(parents=True, exist_ok=True)
        
        for f in emls:
            target = od / (f.stem + ".pdf")
            print(f"In corso: {f.name}...", end=" ", flush=True)
            try:
                process_single_eml(f, target, fp, args.font_size, args.margins, args.include_inline)
                print("OK")
            except Exception as e:
                print(f"ERRORE: {e}")
    else:
        if not ip.is_file():
            print(f"Errore: {ip} non trovato.")
            return
        op = Path(args.output).resolve() if args.output else ip.with_suffix(".pdf")
        print(f"Conversione {ip.name}...", end=" ", flush=True)
        process_single_eml(ip, op, fp, args.font_size, args.margins, args.include_inline)
        print("OK")

if __name__ == "__main__":
    main()
