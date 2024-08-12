import os
import zipfile
import xml.etree.ElementTree as ET
import shutil

def extract_docx(docx_path, extract_dir):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def recreate_docx(original_extract_dir, new_docx_path):
    with zipfile.ZipFile(new_docx_path, 'w') as docx:
        for foldername, subfolders, filenames in os.walk(original_extract_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, original_extract_dir)
                docx.write(file_path, arcname)

def translate_text(text):
    # Mock translation function (for example, converting text to uppercase)
    # Replace this function with actual translation logic (e.g., using an API)
    return text.upper()

def replace_paragraph_text_with_translation(xml_file_path):
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()

        # Namespace dictionary for looking up XML elements with namespaces
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }

        # Find all paragraph elements
        paragraphs = root.findall('.//w:p', namespaces)

        for p in paragraphs:
            # Find all runs within the paragraph
            runs = p.findall('.//w:r', namespaces)
            for r in runs:
                text_element = r.find('.//w:t', namespaces)
                if text_element is not None and text_element.text:
                    # Translate the text and replace it
                    translated_text = translate_text(text_element.text)
                    text_element.text = translated_text

        tree.write(xml_file_path, encoding='utf-8', xml_declaration=True)
    
    except ET.ParseError as e:
        print(f"Error parsing {xml_file_path}: {e}")

def analyze_and_translate_docx(input_docx, output_docx):
    # Step 1: Extract the original docx file
    extract_dir = 'extracted_docx'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    extract_docx(input_docx, extract_dir)

    # Step 2: Replace text in document.xml with translated text
    document_xml_path = os.path.join(extract_dir, 'word/document.xml')
    if os.path.exists(document_xml_path):
        replace_paragraph_text_with_translation(document_xml_path)
    else:
        print("document.xml not found in the extracted .docx file.")

    # Step 3: Recreate the docx file from the modified content
    recreate_docx(extract_dir, output_docx)

    # Clean up the extracted directory
    shutil.rmtree(extract_dir)

# Usage
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
analyze_and_translate_docx(input_docx, output_docx)
