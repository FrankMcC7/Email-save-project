import os
import argparse
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

def translate_document(file_path, output_dir=None):
    """
    Translate a document from any language to English
    
    Args:
        file_path (str): Path to the document
        output_dir (str, optional): Directory to save translated file. Defaults to same directory.
    
    Returns:
        tuple: (output_path, source_language)
    """
    file_path = Path(file_path)
    
    # Determine the output directory
    if output_dir:
        output_dir = Path(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    else:
        output_dir = file_path.parent
    
    # Create output filename
    output_path = output_dir / f"{file_path.stem}_translated_to_en{file_path.suffix}"
    
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

def translate_directory(directory_path, output_dir=None, extensions=None):
    """
    Translate all documents in a directory
    
    Args:
        directory_path (str): Path to directory containing documents
        output_dir (str, optional): Directory to save translated files
        extensions (list, optional): List of file extensions to process
    """
    if extensions is None:
        extensions = ['.txt', '.docx', '.pdf']
    
    directory_path = Path(directory_path)
    
    # Process all files with the specified extensions
    for ext in extensions:
        for file_path in directory_path.glob(f"*{ext}"):
            translate_document(file_path, output_dir)

def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='Translate documents from any language to English')
    parser.add_argument('path', help='Path to document or directory')
    parser.add_argument('--output', '-o', help='Output directory for translated documents')
    parser.add_argument('--recursive', '-r', action='store_true', help='Process directories recursively')
    parser.add_argument('--extensions', '-e', nargs='+', default=['.txt', '.docx', '.pdf'], 
                        help='File extensions to process (default: .txt .docx .pdf)')
    
    args = parser.parse_args()
    
    path = Path(args.path)
    
    if path.is_file():
        # Translate single file
        translate_document(path, args.output)
    elif path.is_dir():
        # Translate all files in directory
        if args.recursive:
            # Process recursively
            for root, _, _ in os.walk(path):
                translate_directory(root, args.output, args.extensions)
        else:
            # Process just the top directory
            translate_directory(path, args.output, args.extensions)
    else:
        print(f"Error: {path} does not exist.")

if __name__ == "__main__":
    main()