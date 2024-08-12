import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
from googletrans import Translator  # You can replace this with another translation API if needed

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

def translate_text(text, target_language='en'):
    # translator = Translator()
    # translated = translator.translate(text, dest=target_language)
    return text.upper()

def replace_paragraph_text_with_translation(xml_file_path, target_language='en'):
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
            # Collect all text within the paragraph into a single string
            paragraph_text = ""
            for r in p.findall('.//w:r', namespaces):
                for t in r.findall('.//w:t', namespaces):
                    if t.text:
                        paragraph_text += t.text

            if paragraph_text.strip():  # Only translate if there is non-empty text
                # Translate the entire paragraph text as a whole
                translated_text = translate_text(paragraph_text, target_language)

                # Clear existing runs but keep one to preserve the structure
                for r in p.findall('.//w:r', namespaces):
                    for t in r.findall('.//w:t', namespaces):
                        t.clear()

                # Insert the translated text back into the paragraph
                first_run = p.find('.//w:r', namespaces)
                if first_run is not None:
                    first_text_element = first_run.find('.//w:t', namespaces)
                    if first_text_element is not None:
                        first_text_element.text = translated_text
                    else:
                        new_text_element = ET.SubElement(first_run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                        new_text_element.text = translated_text

        tree.write(xml_file_path, encoding='utf-8', xml_declaration=True)

    except ET.ParseError as e:
        print(f"Error parsing {xml_file_path}: {e}")

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
        replace_paragraph_text_with_translation(document_xml_path, target_language)
    else:
        print("document.xml not found in the extracted .docx file.")

    # Step 3: Recreate the docx file from the modified content
    recreate_docx(extract_dir, output_docx)

    # Clean up the extracted directory
    shutil.rmtree(extract_dir)

# Usage
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
target_language = 'en'  # Target language code (e.g., 'en' for English, 'fr' for French)
analyze_and_translate_docx(input_docx, output_docx, target_language)
