from docx2pdf import convert
from PyPDF2 import PdfMerger
import os

# Function to convert Word documents to PDF
def convert_docs_to_pdfs(input_folder, output_folder):
    pdf_files = []
    for doc_file in os.listdir(input_folder):
        if doc_file.endswith('.docx'):
            input_path = os.path.join(input_folder, doc_file)
            output_path = os.path.join(output_folder, doc_file.replace('.docx', '.pdf'))
            convert(input_path, output_path, keep_active=True)
            pdf_files.append(output_path)
    return pdf_files

# Function to merge PDF files
def merge_pdfs(pdf_files, output_file):
    merger = PdfMerger()
    for pdf_file in pdf_files:
        merger.append(pdf_file)
    merger.write(output_file)
    merger.close()

# Define input and output folders
input_folder = 'input'
output_folder = 'output'

# Ensure the output folder exists
os.makedirs(output_folder, exist_ok=True)

# Convert Word documents to PDFs
pdf_files = convert_docs_to_pdfs(input_folder, output_folder)

# Merge the resulting PDF files into a single PDF
output_file = os.path.join(output_folder, 'merged_document.pdf')
merge_pdfs(pdf_files, output_file)

print(f"The Word documents from {input_folder} have been converted to PDFs and merged into {output_file}.")