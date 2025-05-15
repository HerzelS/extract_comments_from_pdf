from docx import Document

# Load the Word document
doc = Document("comments_output_1.docx")

# Access the first table
table = doc.tables[0]

# Iterate over the rows
for i, row in enumerate(table.rows):
    # Ensure the row has 3 cells (you must add an extra empty column in Word manually if not)
    if len(row.cells) < 3:
        continue  # skip rows without enough columns

    col1_text = row.cells[1].text.strip()
    col2_text = row.cells[3].text.strip()

    # Combine and store in the 3rd column
    if i == 0:
        row.cells[2].text = "Combined"  # Header row
    else:
        row.cells[2].text = f"{col1_text}, {col2_text}"

# Save the modified document
doc.save("my_table_updated.docx")

