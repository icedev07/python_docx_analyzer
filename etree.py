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

def print_xml_text_values(xml_file_path):
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        for elem in root.iter():
            if elem.text:
                print(elem.text.strip())  # Print the text, stripping leading/trailing whitespace
    except ET.ParseError as e:
        print(f"Error parsing {xml_file_path}: {e}")

def analyze_and_duplicate_docx(input_docx, output_docx):
    # Step 1: Extract the original docx file
    extract_dir = 'extracted_docx'
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)
    
    extract_docx(input_docx, extract_dir)

    # Step 2: Print text content from XML files
    document_xml_path = os.path.join(extract_dir, 'word/document.xml')
    if os.path.exists(document_xml_path):
        print("Text content from document.xml:")
        print_xml_text_values(document_xml_path)
    else:
        print("document.xml not found in the extracted .docx file.")

    # Optionally, print text from other XML files, such as headers, footers, etc.
    for root, dirs, files in os.walk(os.path.join(extract_dir, 'word')):
        for file in files:
            if file.endswith('.xml') and file != 'document.xml':
                xml_file_path = os.path.join(root, file)
                print(f"\nText content from {file}:")
                print_xml_text_values(xml_file_path)

    # Step 3: Recreate the docx file from the extracted and analyzed content
    recreate_docx(extract_dir, output_docx)

    # Clean up the extracted directory
    shutil.rmtree(extract_dir)

# Usage
input_docx = 'input.docx'  # Path to your input .docx file
output_docx = 'output.docx'  # Path to the output .docx file to be generated
analyze_and_duplicate_docx(input_docx, output_docx)
