import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
from googletrans import Translator

# Function to extract a .docx file
def extract_docx(docx_path, extract_dir):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

# Function to recreate a .docx file from the extracted content
def recreate_docx(original_extract_dir, new_docx_path):
    with zipfile.ZipFile(new_docx_path, 'w') as docx:
        for foldername, subfolders, filenames in os.walk(original_extract_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, original_extract_dir)
                docx.write(file_path, arcname)

# Function to translate text
def translate_text(text, target_language='en'):
    
    return text+'++'


# Function to translate paragraphs while preserving the XML structure
def translate_paragraphs(root, target_language='en'):
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    for paragraph in root.findall('.//w:p', namespaces):
        text_elements = paragraph.findall('.//w:t', namespaces)
        if text_elements:
            full_text = ''.join([elem.text for elem in text_elements if elem.text])
            if full_text.strip():
                translated_text = translate_text(full_text, target_language)
                translated_words = translated_text.split()
                word_index = 0
                for elem in text_elements:
                    original_words = elem.text.split()
                    elem.text = ' '.join(translated_words[word_index:word_index + len(original_words)])
                    word_index += len(original_words)

# Function to handle headers and footers
def translate_headers_footers(extract_dir, target_language='en'):
    word_dir = os.path.join(extract_dir, 'word')
    header_footer_files = [f for f in os.listdir(word_dir) if f.startswith('header') or f.startswith('footer')]

    for file_name in header_footer_files:
        file_path = os.path.join(word_dir, file_name)
        if os.path.exists(file_path):
            tree = ET.parse(file_path)
            root = tree.getroot()

            # Translate only the text elements while preserving the structure
            translate_paragraphs(root, target_language)

            # Write back the modified XML while preserving structure
            tree.write(file_path, xml_declaration=True, encoding='utf-8')

# Main function to handle translation of a .docx file
def analyze_and_translate_docx(input_docx, output_docx, target_language='en'):
    extract_dir = 'extracted_docx'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    extract_docx(input_docx, extract_dir)

    document_xml_path = os.path.join(extract_dir, 'word/document.xml')
    if os.path.exists(document_xml_path):
        tree = ET.parse(document_xml_path)
        root = tree.getroot()
        
        # Translate the main document content by paragraph
        translate_paragraphs(root, target_language)

        # Write back the modified XML
        tree.write(document_xml_path, xml_declaration=True, encoding='utf-8')

    # Translate headers and footers while preserving structure and placement
    translate_headers_footers(extract_dir, target_language)

    recreate_docx(extract_dir, output_docx)
    shutil.rmtree(extract_dir)

# Example usage:
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
target_language = 'en'  # Target language code (e.g., 'en' for English, 'fr' for French)
analyze_and_translate_docx(input_docx, output_docx, target_language)
