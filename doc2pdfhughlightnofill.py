import os
import win32com.client as win32
WD_COLOR_AUTOMATIC = -16777216
def remove_highlight_from_doc(doc):
    # Go through the entire document range and clear font shading
    for paragraph in doc.Paragraphs:
        # Apply the 'no highlight' (automatic) shading directly to the paragraph's range
        paragraph.Range.Shading.BackgroundPatternColor = WD_COLOR_AUTOMATIC

def convert_all_docx_to_pdf(folder_path):
    # Check if the folder exists
    if not os.path.isdir(folder_path):
        raise NotADirectoryError(f"Directory not found: {folder_path}")
    
    # Initialize the Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # Keep Word application hidden during the process
    
    # Walk through all subfolders and files
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            # Skip temporary files created by Word (those starting with '~$')
            if filename.startswith('~$'):
                continue
            
            if filename.endswith('.doc') or filename.endswith('.docx'):
                docx_path = os.path.join(root, filename)
                docx_path_no_h = os.path.join(root, filename+"no_h")

                pdf_path = os.path.splitext(docx_path)[0] + '.pdf'
                
                print(f"Converting {docx_path} to {pdf_path}...")

                try:
                    # Open the .docx file
                    doc_no_h = word.Documents.Open(docx_path)

                    # Remove any shading (no fill)
                    remove_highlight_from_doc(doc_no_h)
                    doc_no_h.SaveAs(docx_path_no_h)
                    doc_no_h.Close()
                    doc = word.Documents.Open(docx_path)
                    # Save as PDF (file format 17 is for PDF)
                    doc.SaveAs(pdf_path, FileFormat=17)
                    
                    # Close the document
                    doc.Close()
                
                except Exception as e:
                    print(f"Failed to process {docx_path}: {e}")
    
    # Quit Word application
    word.Quit()

    print("Conversion completed.")

# Example usage
folder_path = r"S:\Users\L\Downloads\OneDrive_1_10-6-2024 (1)"
convert_all_docx_to_pdf(folder_path)
