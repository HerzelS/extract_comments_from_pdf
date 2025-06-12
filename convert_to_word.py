import os
import time
import re
import pdfplumber
from docx import Document

def convert_numbered_paragraph_pdfs_to_word(pdf_folder, output_folder):
    """
    Converts numbered-paragraph PDF files in a folder into Word documents,
    fixing line breaks and reconstructing clean paragraphs.
    """
    os.makedirs(output_folder, exist_ok=True)
    n_reports = 0
    start_time = time.time()

    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            doc_name = os.path.splitext(filename)[0] + ".docx"
            output_path = os.path.join(output_folder, doc_name)

            with pdfplumber.open(pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"

            # Fix broken paragraph numbers like "1\n1\n9." → "119."
            # Merge isolated digits that are broken across lines into full numbers (like 1\n2\n0. → 120.)
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\n(\d)\.', r'\1\2\3.', text)
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\.', r'\1\2.', text)



            # Merge all non-paragraph-break line breaks into spaces
            text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)

            # Optional: normalize spaces
            text = re.sub(r'[ \t]+', ' ', text)

            # Now split into paragraphs by paragraph numbers, e.g. "67."
            paragraphs = re.split(r'(?=\n?\d{1,3}\.\s)', text)

            # Create Word doc and write paragraphs
            doc = Document()
            for para in paragraphs:
                clean_para = para.strip()
                if clean_para:
                    doc.add_paragraph(clean_para)

            doc.save(output_path)
            n_reports += 1
            print(f"{n_reports} PDF(s) converted to Word: {filename}")

    elapsed_time = (time.time() - start_time) / 60
    print(f"\nConversion completed.")
    print(f"Total PDFs processed: {n_reports}")
    print(f"Process completed in {elapsed_time:.2f} minutes.")

    
convert_numbered_paragraph_pdfs_to_word(pdf_folder="test_pdf", output_folder="word_outputs")