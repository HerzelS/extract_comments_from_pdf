import os
import time
import re
import pdfplumber
from docx import Document


def fix_split_number_lines(text):
    """
    Fix lines that contain only digits and are followed by numbered paragraphs.
    Converts sequences like:
    8\n9. Some text...
    Into:
    89. Some text...
    """
    lines = text.splitlines()
    cleaned_lines = []
    skip_next = False

    for i in range(len(lines)):
        if skip_next:
            skip_next = False
            continue

        current_line = lines[i].strip()

        # Check if current line is only a digit
        if current_line.isdigit() and i + 1 < len(lines):
            next_line = lines[i + 1].strip()

            # If the next line starts with a number and a period (e.g. 9.), join them
            if re.match(r'^\d+\.', next_line):
                combined_line = current_line + next_line
                cleaned_lines.append(combined_line)
                skip_next = True
            else:
                cleaned_lines.append(current_line)
        else:
            cleaned_lines.append(current_line)

    return '\n'.join(cleaned_lines)


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

            # First fix isolated digit line breaks (e.g. "8\n9." → "89.")
            text = fix_split_number_lines(text)

            # Fix broken paragraph numbers spanning 3 digits split over lines (e.g., "1\n1\n9." → "119.")
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\n(\d)\.', r'\1\2\3.', text)
            # Fix broken paragraph numbers spanning 2 digits split over lines (e.g., "1\n2." → "12.")
            text = re.sub(r'(?:\n)?(?<!\d)(\d)\n(\d)\.', r'\1\2.', text)

            # Merge all line breaks that are NOT paragraph breaks (paragraph breaks expected before a number + period)
            text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)

            # Normalize spaces and tabs
            text = re.sub(r'[ \t]+', ' ', text)

            # Split into paragraphs by paragraph numbers like "67."
            paragraphs = re.split(r'(?=\n?\d{1,3}\.\s)', text)

            # Create Word document and write paragraphs
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