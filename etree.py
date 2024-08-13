import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
import re
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

# Translation function
def translate_text(text, target_language='en'):
    return text+'+++'

# Function to translate text within paragraphs while preserving numbering and structure
def translate_paragraphs_preserve_structure(xml_content, target_language='en'):
    root = ET.fromstring(xml_content)
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Pattern to identify numbering/citation markers (e.g., "1.", "2.")
    marker_pattern = r'(\[\d+\]|\d+[\.\)])'

    for paragraph in root.findall('.//w:p', namespaces):
        paragraph_text = ""
        text_nodes = []

        # Extract all text in the paragraph
        for node in paragraph.findall('.//w:t', namespaces):
            if node.text:
                paragraph_text += node.text + " "
                text_nodes.append(node)

        # Translate the entire paragraph text
        if paragraph_text.strip():
            # Preserve the markers
            markers = re.findall(marker_pattern, paragraph_text)
            paragraph_text_no_markers = re.sub(marker_pattern, "<<<MARKER>>>", paragraph_text)

            # Translate the text without markers
            translated_text = translate_text(paragraph_text_no_markers.strip(), target_language)
            
            # Reinsert the markers into the translated text
            for marker in markers:
                translated_text = translated_text.replace("<<<MARKER>>>", marker, 1)

            # Reinsert translated text into the original nodes
            idx = 0
            words = translated_text.split()
            for node in text_nodes:
                original_text = node.text
                node.text = ""
                word_count = len(original_text.split())
                node.text = " ".join(words[idx:idx + word_count])
                idx += word_count

    return ET.tostring(root, encoding='unicode')

# Function to translate headers and footers if they exist
def translate_headers_footers(extract_dir, target_language='en'):
    word_dir = os.path.join(extract_dir, 'word')
    header_footer_files = [f for f in os.listdir(word_dir) if f.startswith('header') or f.startswith('footer')]

    for file_name in header_footer_files:
        file_path = os.path.join(word_dir, file_name)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            translated_content = translate_paragraphs_preserve_structure(content, target_language)

            with open(file_path, 'w', encoding='utf-8') as file:
                file.write(translated_content)

# Main function to handle translation of a .docx file
def analyze_and_translate_docx(input_docx, output_docx, target_language='en'):
    extract_dir = 'extracted_docx'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    extract_docx(input_docx, extract_dir)

    document_xml_path = os.path.join(extract_dir, 'word/document.xml')
    if os.path.exists(document_xml_path):
        with open(document_xml_path, 'r', encoding='utf-8') as file:
            document_xml_content = file.read()
        
        translated_xml_content = translate_paragraphs_preserve_structure(document_xml_content, target_language)

        with open(document_xml_path, 'w', encoding='utf-8') as file:
            file.write(translated_xml_content)

    translate_headers_footers(extract_dir, target_language)

    recreate_docx(extract_dir, output_docx)
    shutil.rmtree(extract_dir)

# Example usage:
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
target_language = 'en'  # Target language code (e.g., 'en' for English, 'fr' for French)
analyze_and_translate_docx(input_docx, output_docx, target_language)
