# PDF Comment Extractor to Word

This Python script extracts comments and highlighted text (as "comment scope") from a PDF file and exports them into a formatted Word document table.

## Features

- Extracts annotations (comments) from PDF files
- Captures the highlighted text as "Comment Scope"
- Exports to a `.docx` Word document with a structured table
- Includes author, page number, comment text, and placeholder for team responses
- Outputs in **landscape orientation** for better readability

## Example Output Table

| Author | Page | Paragraph | Comment Scope | Comment Text | Team Response |
|--------|------|-----------|----------------|---------------|----------------|

## Requirements

- Python 3.x
- [PyMuPDF (`fitz`)](https://pymupdf.readthedocs.io/)
- [python-docx](https://python-docx.readthedocs.io/)

Install dependencies:

```bash
pip install pymupdf python-docx
````

## Usage

```python
from extract_comments import extract_comments_to_word

extract_comments_to_word("your_file.pdf", "output_comments.docx")
```

## Function Description
```python
def extract_comments_to_word(pdf_path, word_output_path):
    """
    Extracts comments and highlighted text from a PDF and writes them to a Word table.

    Args:
        pdf_path (str): Path to the input PDF file.
        word_output_path (str): Path to the output Word (.docx) file.
    """
```
## Notes

* Only highlight annotations are currently supported for extracting the comment scope.
* Paragraph information is marked as `N/A` since it's not natively tracked in PDFs.
* The table includes a `Team Response` column for later collaboration or review.

## License

### MIT License


