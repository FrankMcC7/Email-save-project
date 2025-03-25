import os
import argparse
import tkinter as tk
from tkinter import filedialog
from googletrans import Translator
from pathlib import Path
import docx
import PyPDF2
from tqdm import tqdm

def detect_and_translate(text):
    """
    Detect language and translate text to English
    
    Args:
        text (str): Text to translate
        
    Returns:
        tuple: (translated_text, source_language)
    """
    translator = Translator()
    
    # First detect the language
    detection = translator.detect(text[:5000])  # Use first 5000 chars for detection
    source_lang = detection.lang
    
    # Translate the text to English if it's not already in English
    if source_lang != 'en':
        # Break into chunks if text is very long (Google Translate API has character limits)
        chunks = [text[i:i+5000] for i in range(0, len(text), 5000)]
        translated_chunks = []
        
        # Show progress bar for long translations
        for chunk in tqdm(chunks, desc=f"Translating from {source_lang} to English"):
            translation = translator.translate(chunk, dest='en', src=source_lang)
            translated_chunks.append(translation.text)
            
        translated_text = ' '.join(translated_chunks)
        return translated_text, source_lang
    else:
        return text, 'en'

def extract_text_from_txt(file_path):
    """Extract text from a .txt file"""
    with open(file_path, 'r', encoding='utf-8', errors='replace') as file:
        return file.read()

def extract_text_from_docx(file_path):
    """Extract text from a .docx file"""
    doc = docx.Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def extract_text_from_pdf(file_path):
    """Extract text from a .pdf file"""
    with open(file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page_num in range(len(reader.pages)):
            text += reader.pages[page_num].extract_text() + '\n'
        return text

def save_translated_text(translated_text, output_path):
    """Save translated text to output file"""
    with open(output_path, 'w', encoding='utf-8') as file:
        file.write(translated_text)

def select_input_file():
    """
    Open a file dialog to select an input file
    
    Returns:
        str: Path to the selected file
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.askopenfilename(
        title="Select document to translate",
        filetypes=(
            ("Text files", "*.txt"),
            ("Word documents", "*.docx"),
            ("PDF files", "*.pdf"),
            ("All files", "*.*")
        )
    )
    return file_path

def select_output_file(default_filename):
    """
    Open a file dialog to select where to save the output file
    
    Args:
        default_filename (str): Default filename to suggest
        
    Returns:
        str: Path where the output file should be saved
    """
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    file_path = filedialog.asksaveasfilename(
        title="Save translated document as",
        defaultextension=Path(default_filename).suffix,
        initialfile=Path(default_filename).name,
        filetypes=(
            ("Text files", "*.txt"),
            ("Word documents", "*.docx"),
            ("PDF files", "*.pdf"),
            ("All files", "*.*")
        )
    )
    return file_path

def translate_document(file_path=None):
    """
    Translate a document from any language to English
    
    Args:
        file_path (str, optional): Path to the document. If None, will prompt user to select.
    
    Returns:
        tuple: (output_path, source_language)
    """
    # If no file path provided, ask user to select one
    if not file_path:
        file_path = select_input_file()
        if not file_path:  # User canceled selection
            print("No file selected. Exiting.")
            return None, None
    
    file_path = Path(file_path)
    
    # Create default output filename
    default_output_path = file_path.parent / f"{file_path.stem}_translated_to_en{file_path.suffix}"
    
    # Ask user where to save the translated file
    output_path = select_output_file(str(default_output_path))
    if not output_path:  # User canceled selection
        print("No output location selected. Exiting.")
        return None, None
    
    output_path = Path(output_path)
    
    # Extract text based on file extension
    file_extension = file_path.suffix.lower()
    
    print(f"Processing {file_path.name}...")
    
    try:
        if file_extension == '.txt':
            text = extract_text_from_txt(file_path)
        elif file_extension == '.docx':
            text = extract_text_from_docx(file_path)
        elif file_extension == '.pdf':
            text = extract_text_from_pdf(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_extension}")
        
        # Detect language and translate
        translated_text, source_lang = detect_and_translate(text)
        
        # Save translated text
        save_translated_text(translated_text, output_path)
        
        print(f"Translation complete: {file_path.name} ({source_lang}) → {output_path.name} (en)")
        return str(output_path), source_lang
        
    except Exception as e:
        print(f"Error processing {file_path.name}: {str(e)}")
        return None, None

def main():
    parser = argparse.ArgumentParser(description='Translate documents from any language to English')
    parser.add_argument('--file', '-f', help='Path to document (optional, if not provided will open file dialog)')
    
    args = parser.parse_args()
    
    # If file argument provided, use that, otherwise prompt user
    translate_document(args.file)
    
    print("Translation completed.")

if __name__ == "__main__":
    main()