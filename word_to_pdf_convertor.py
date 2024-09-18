from docx2pdf import convert
import os

def convert_all_docx_to_pdf(folder_path):
    # Get all .docx files in the folder
    docx_files = [f for f in os.listdir(folder_path) if f.endswith('.docx')]

    # Convert each .docx file to PDF
    for docx_file in docx_files:
        docx_file_path = os.path.join(folder_path, docx_file)
        print(f"Converting {docx_file_path} to PDF...")
        convert(docx_file_path)  # This converts the .docx file to PDF in the same folder
        print(f"Converted: {docx_file_path} to PDF.")

if __name__ == '__main__':
    folder_path = './files'  # Replace with the folder path where your .docx files are stored
    convert_all_docx_to_pdf(folder_path)
