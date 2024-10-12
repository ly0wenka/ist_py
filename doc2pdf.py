import os
import win32com.client as win32

def convert_all_docx_to_pdf(folder_path):
    # Check if the folder exists
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"Directory not found: {folder_path}")
    
    # Initialize the Word application
    word = win32.Dispatch('Word.Application')
    
    # Walk through all subfolders and files
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename.endswith('.doc') or filename.endswith('.docx'):
                docx_path = os.path.join(root, filename)
                pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
                
                print(f"Converting {docx_path} to {pdf_path}...")
                
                # Open the .docx file
                doc = word.Documents.Open(docx_path)
                
                # Save as PDF (file format 17 is for PDF)
                doc.SaveAs(pdf_path, FileFormat=17)
                
                # Close the document
                doc.Close()
    
    # Quit Word application
    word.Quit()

    print("Conversion completed.")

# Example usage
folder_path = r"S:\Users\L\Downloads\OneDrive_4_10-12-2024"
convert_all_docx_to_pdf(folder_path)

