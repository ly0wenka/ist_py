import os
import subprocess

def convert_all_docx_to_pdf_with_libreoffice(folder_path):
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"Directory not found: {folder_path}")
    
    # Walk through all subfolders and files
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            # Skip temporary files
            if filename.startswith('~$'):
                continue

            if filename.endswith('.doc') or filename.endswith('.docx'):
                docx_path = os.path.join(root, filename)
                pdf_path = os.path.splitext(docx_path)[0] + '.pdf'

                print(f"Converting {docx_path} to {pdf_path} using LibreOffice...")

                try:
                    # Run unoconv to convert the document to PDF
                    subprocess.run(['unoconv', '-f', 'pdf', docx_path], check=True)
                
                except subprocess.CalledProcessError as e:
                    print(f"Failed to process {docx_path}: {e}")
    
    print("Conversion completed.")

# Example usage
folder_path = r"S:\Users\L\Downloads\OneDrive_1_10-6-2024 (1)"
convert_all_docx_to_pdf_with_libreoffice(folder_path)
