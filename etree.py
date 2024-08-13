import os
import zipfile
import xml.etree.ElementTree as ET
import shutil

# Mock Translation Function (For demonstration purposes)
def translate_text(text, target_language='en'):
    # This mock function just adds '++' after the text. Replace it with your actual translation logic.
    return text + '++'

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

# Function to translate text within paragraphs while preserving structure and excluding hyperlinks
def translate_paragraphs_exclude_hyperlinks(xml_content, target_language='en'):
    root = ET.fromstring(xml_content)
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    for paragraph in root.findall('.//w:p', namespaces):
        for node in paragraph.findall('.//w:t', namespaces):
            # Check if the text is part of a hyperlink or URL
            parent = node.find("..")
            if parent is not None and (parent.tag.endswith('hyperlink') or 'http' in node.text or 'www' in node.text):
                continue  # Skip translation for hyperlinks and URLs

            if node.text:
                original_text = node.text
                translated_text = translate_text(original_text.strip(), target_language)
                node.text = translated_text

    return ET.tostring(root, encoding='unicode')

# Function to translate headers, footers, and footnotes
def translate_headers_footers_footnotes(extract_dir, target_language='en'):
    word_dir = os.path.join(extract_dir, 'word')
    special_files = [f for f in os.listdir(word_dir) if f.startswith('header') or f.startswith('footer') or f == 'footnotes.xml']

    for file_name in special_files:
        file_path = os.path.join(word_dir, file_name)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            translated_content = translate_paragraphs_exclude_hyperlinks(content, target_language)

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
        
        translated_xml_content = translate_paragraphs_exclude_hyperlinks(document_xml_content, target_language)

        with open(document_xml_path, 'w', encoding='utf-8') as file:
            file.write(translated_xml_content)

    translate_headers_footers_footnotes(extract_dir, target_language)

    recreate_docx(extract_dir, output_docx)
    shutil.rmtree(extract_dir)

# Example usage:
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
target_language = 'en'  # Target language code (e.g., 'en' for English, 'fr' for French)
analyze_and_translate_docx(input_docx, output_docx, target_language)
