import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
from googletrans import Translator  # Replace with your translation API if needed

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
    # print("Original paragraph:\n", text)
    translated = text.upper() + '++'
    print(translated)
    print("=======================")
    return translated

# Function to translate each paragraph as a whole
def translate_paragraphs(xml_content, target_language='en'):
    root = ET.fromstring(xml_content)
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Find all paragraphs in the document
    for paragraph in root.findall('.//w:p', namespaces):
        # Gather all text within this paragraph
        paragraph_text = ''.join(node.text for node in paragraph.findall('.//w:t', namespaces) if node.text)
        if paragraph_text.strip():  # Skip empty paragraphs
            translated_text = translate_text(paragraph_text, target_language)

            # Replace original text nodes with translated text
            text_nodes = paragraph.findall('.//w:t', namespaces)
            for node in text_nodes:
                node.clear()  # Clear the text node

            # Reassign the translated text back into the first text node
            if text_nodes:
                text_nodes[0].text = translated_text

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
        
        # Perform the translation by paragraph
        translated_xml_content = translate_paragraphs(document_xml_content, target_language)
        
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
