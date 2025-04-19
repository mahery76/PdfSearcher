import os
import subprocess
import pdfplumber  # Changed from PyPDF2 to pdfplumber
import re
import contextlib
import io

def search_and_open_pdfs(folder_path, search_term):
    """
    Searches through all PDFs in a folder and its subfolders for a keyword
    and opens each PDF at the pages where the keyword appears.
    
    Args:
        folder_path (str): Path to folder containing PDF files
        search_term (str): Keyword to search for
    """
    # Convert to absolute path
    folder_path = os.path.abspath(folder_path)
    
    # Check if folder exists
    if not os.path.isdir(folder_path):
        print(f"Folder not found: {folder_path}")
        return
    
    # Dictionary to store results: {pdf_path: [page_numbers]}
    results = {}
    
    # Count files processed and matches found
    total_files = 0
    files_with_matches = 0
    
    print(f"Searching for '{search_term}' in {folder_path} and its subfolders...")
    
    # Walk through directory and all subdirectories
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if filename.lower().endswith('.pdf'):
                total_files += 1
                pdf_path = os.path.join(root, filename)
                matches = search_pdf(pdf_path, search_term)
                
                if matches:
                    files_with_matches += 1
                    results[pdf_path] = matches
                    relative_path = os.path.relpath(pdf_path, folder_path)
                    print(f"Found matches in {relative_path} on pages: {matches}")
    
    print(f"\nSearch complete. Found matches in {files_with_matches} of {total_files} PDF files.")
    
    # Ask user which PDFs to open
    open_pdfs(results)

def search_pdf(pdf_path, search_term):
    """
    Searches a PDF file for a keyword and returns page numbers where the keyword appears.
    
    Args:
        pdf_path (str): Path to PDF file
        search_term (str): Keyword to search for
        
    Returns:
        list: List of page numbers (1-based) where the keyword appears
    """
    matches = []
    
    try:
        # Using pdfplumber instead of PyPDF2
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages):
                try:
                    text = page.extract_text()
                    if text and re.search(search_term, text, re.IGNORECASE):
                        matches.append(i + 1)  # 1-based page numbering
                except Exception as e:
                    print(f"Error processing page {i+1} in {pdf_path}: {e}")
                    continue
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
    
    return matches

def open_pdf_with_foxit(pdf_path, page_number):
    """
    Opens a PDF file at a specific page using Foxit Reader
    
    Args:
        pdf_path (str): Full path to the PDF file
        page_number (int): Page number to open
    """
    # Default Foxit Reader installation path
    foxit_path = r"C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe"
    
    # Check if Foxit Reader exists at the default location
    if not os.path.exists(foxit_path):
        # Try alternate installation location
        foxit_path = r"C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe"
        if not os.path.exists(foxit_path):
            print("Foxit Reader not found. Please update the path in the script.")
            return False
    
    # Construct the command
    command = [foxit_path, "/A", f"page={page_number}", pdf_path]
    
    try:
        # Run the command
        subprocess.Popen(command)
        print(f"Opening {os.path.basename(pdf_path)} at page {page_number}")
        return True
    except Exception as e:
        print(f"Error opening PDF: {e}")
        return False

def open_pdfs(results):
    """
    Asks user which PDFs to open and opens them at the specified pages
    
    Args:
        results (dict): Dictionary of {pdf_path: [page_numbers]}
    """
    if not results:
        print("No matches found.")
        return
    
    print("\nFound matches in the following PDFs:")
    pdf_list = list(results.keys())
    
    for i, pdf_path in enumerate(pdf_list):
        filename = os.path.basename(pdf_path)
        folder = os.path.dirname(pdf_path)
        pages = results[pdf_path]
        print(f"{i+1}. {filename} ({len(pages)} matches on pages {pages})")
        print(f"   Location: {folder}")
    
    while True:
        choice = input("\nEnter the number of the PDF to open (or 'all' to open all, 'q' to quit): ")
        
        if choice.lower() == 'q':
            break
            
        if choice.lower() == 'all':
            for pdf_path, pages in results.items():
                # Open each PDF at its first match page
                if pages:
                    open_pdf_with_foxit(pdf_path, pages[0])
            break
        
        try:
            idx = int(choice) - 1
            if 0 <= idx < len(pdf_list):
                pdf_path = pdf_list[idx]
                pages = results[pdf_path]
                
                if len(pages) == 1:
                    # Only one match, open directly
                    open_pdf_with_foxit(pdf_path, pages[0])
                else:
                    # Multiple matches, ask which page to open
                    page_choice = input(f"Multiple matches found. Enter page number to open {pages} or 'all' to open all: ")
                    
                    if page_choice.lower() == 'all':
                        for page in pages:
                            open_pdf_with_foxit(pdf_path, page)
                    else:
                        try:
                            page = int(page_choice)
                            if page in pages:
                                open_pdf_with_foxit(pdf_path, page)
                            else:
                                print(f"Invalid page number. Must be one of {pages}")
                        except ValueError:
                            print("Invalid input. Please enter a valid page number.")
            else:
                print("Invalid choice. Please enter a number from the list.")
        except ValueError:
            print("Invalid input. Please enter a number or 'all' or 'q'.")

if __name__ == "__main__":
    # Get user input
    # folder_path = input("Enter the folder path containing PDF files: ")
    # search_term = input("Enter the keyword to search for: ")
    
    folder_path = r"E:\Devs\Python\PdfSearcher\test"
    search_term = "ouverture de connexion TCP"
    
    search_and_open_pdfs(folder_path, search_term)