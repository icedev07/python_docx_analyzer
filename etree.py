import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
import re
from googletrans import Translator  # You can replace this with another translation API you use

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

# Mock translation function (Replace this with actual translation logic)
def translate_text(text, target_language='en'):
    
    return text.upper()

# Function to preserve XML structure while translating text nodes
def translate_and_preserve_structure(xml_content, target_language='en'):
    # Parse the XML content into an ElementTree
    root = ET.fromstring(xml_content)

    for elem in root.iter():
        if elem.tag.endswith('t'):  # Only target text nodes ('w:t')
            original_text = elem.text if elem.text else ''
            translated_text = translate_text(original_text, target_language)
            
            # Replace the original text with the translated text
            elem.text = translated_text
    
    # Return the modified XML structure as a string
    return ET.tostring(root, encoding='unicode')

# Main function to handle translation of a .docx file
def analyze_and_translate_docx(input_docx, output_docx, target_language='en'):
    # Step 1: Extract the original docx file
    extract_dir = 'extracted_docx'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    extract_docx(input_docx, extract_dir)

    # Step 2: Replace text in document.xml with translated text
    document_xml_path = os.path.join(extract_dir, 'word/document.xml')
    if os.path.exists(document_xml_path):
        with open(document_xml_path, 'r', encoding='utf-8') as file:
            document_xml_content = file.read()
        
        # Perform the translation and structure preservation
        translated_xml_content = translate_and_preserve_structure(document_xml_content, target_language)
        
        # Write the translated content back to the document.xml
        with open(document_xml_path, 'w', encoding='utf-8') as file:
            file.write(translated_xml_content)
    else:
        print("document.xml not found in the extracted .docx file.")

    # Step 3: Recreate the docx file from the modified content
    recreate_docx(extract_dir, output_docx)

    # Clean up the extracted directory
    shutil.rmtree(extract_dir)

# Example usage:
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
target_language = 'en'  # Target language code (e.g., 'en' for English, 'fr' for French)
analyze_and_translate_docx(input_docx, output_docx, target_language)
