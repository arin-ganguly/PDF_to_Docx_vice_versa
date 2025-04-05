from pdf2docx import Converter
from docx2pdf import convert as convert_to_pdf
import os

# Function to convert PDF to Word (.docx)
def pdf_to_word(pdf_file, word_file, start_page=0, end_page=None):
    # Check if the input file is a PDF
    if not pdf_file.endswith(".pdf"):
        print("Error: The input file must be a .pdf file.")
        return

    # Initialize converter and convert with optional page range
    try:
        cv = Converter(pdf_file)
        cv.convert(word_file, start=start_page, end=end_page)
        cv.close()
        print(f"Successfully converted {pdf_file} to {word_file}")
    except Exception as e:
        print(f"Error occurred during PDF to Word conversion: {e}")

# Function to convert Word (.docx) to PDF
def word_to_pdf(word_file, output_dir=None):
    # Ensure the input file is a .docx file
    if not word_file.endswith(".docx"):
        print("Error: The input file must be a .docx file.")
        return

    # Check if the output directory is valid
    if output_dir and not os.path.isdir(output_dir):
        print("Error: The output directory is invalid.")
        return

    # Convert Word to PDF
    try:
        if output_dir:
            convert_to_pdf(word_file, output_dir)  # Output path is directory, not file
        else:
            convert_to_pdf(word_file)  # This will save the PDF in the same directory as the Word file
        print(f"Successfully converted {word_file} to PDF.")
    except Exception as e:
        print(f"Error occurred during Word to PDF conversion: {e}")

def main():
    choice = input("1. PDF to Word\n2. Word to PDF\nChoose an option: ")

    if choice == '1':
        pdf_file = input("Enter path of the PDF file (with .pdf): ")
        word_file = input("Enter path to save Word file (with .docx): ")
        
        # Optional: ask user for page range
        start_page = input("Enter start page (default is 0): ")
        end_page = input("Enter end page (leave blank for all pages): ")
        
        # Handle optional page numbers
        start_page = int(start_page) if start_page.isdigit() else 0
        end_page = int(end_page) if end_page.isdigit() else None
        
        # Call conversion function
        pdf_to_word(pdf_file, word_file, start_page=start_page, end_page=end_page)
    
    elif choice == '2':
        word_file = input("Enter path of the Word file (with .docx): ")
        output_dir = input("Enter directory to save the PDF (leave blank to save in the same folder): ")
        
        if output_dir.strip() == "":
            output_dir = None

        # Call conversion function
        word_to_pdf(word_file, output_dir)
    
    else:
        print("Invalid choice!")

# Start the program
if __name__ == "__main__":
    main()
