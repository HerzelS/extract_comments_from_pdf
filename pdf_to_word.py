import os
import time
import re
import pdfplumber
from docx import Document

def fix_split_number_lines(text):
    """
    Fix lines that contain only digits and are followed by a numbered paragraph (like '6\\n7.').
    Merges them into one line (e.g. '67.') before cleaning and paragraph splitting.
    """
    lines = text.splitlines()
    cleaned_lines = []
    i = 0

    while i < len(lines):
        line = lines[i].strip()
        if line.isdigit() and i + 1 < len(lines):
            next_line = lines[i + 1].strip()
            match = re.match(r'^(\d+)\.(.*)', next_line)
            if match:
                # Merge: "3" and "6. Some text" → "36. Some text"
                merged_line = line + match.group(1) + "." + match.group(2)
                cleaned_lines.append(merged_line.strip())
                i += 2
                continue
        cleaned_lines.append(line)
        i += 1

    return '\n'.join(cleaned_lines)


def clean_text(raw_text):
    """
    Remove control characters and normalize whitespace.
    """
    raw_text = re.sub(r'[\x00-\x1F\x7F]', ' ', raw_text)
    return re.sub(r'\s+', ' ', raw_text).strip()


def split_paragraphs(text):
    """
    Split text into numbered paragraphs like '21. ...', '22. ...'
    """
    return re.split(r'(?=\b\d{1,3}\.\s)', text)


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

            # Step 1: Fix digit-line splits (e.g., "3\n6." → "36.")
            text = fix_split_number_lines(text)

            # Step 2: Fix additional split number patterns like "1\n2." or "1\n1\n9."
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\n(\d)\.', r'\1\2\3.', text)
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\.', r'\1\2.', text)

            # Step 3: Merge line breaks into spaces
            text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)

            # Step 4: Normalize whitespace and control characters
            cleaned_text = clean_text(text)

            # Step 5: Split clean text into paragraphs
            paragraphs = split_paragraphs(cleaned_text)

            # Step 6: Write to Word
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


# Example usage:
convert_numbered_paragraph_pdfs_to_word(pdf_folder="test_pdf", output_folder="word_outputs")