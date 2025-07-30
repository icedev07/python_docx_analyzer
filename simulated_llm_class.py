#
# This is the final version of the file that containing translate document
# extract translated title, description, meta title, meta descriptio
# and extract keywords and syopis##
import os
import zipfile
import xml.etree.ElementTree as ET
import shutil
import regex
import string
import openai
import re
import json
import docx
import time
# from json_extractor import JsonExtractor
import unicodedata
import html
import tempfile
import uuid
import threading

from threading import Lock
import multiprocessing
import threading
import time

# from llama_cpp import Llama
from spire.doc import *
from spire.doc.common import *

from urllib.parse import urlparse
import subprocess

_libreoffice_lock = multiprocessing.Lock()

base_dir = os.path.dirname(os.path.abspath(__file__))
_results_dir_global = os.path.join(base_dir, 'results')

os.makedirs(_results_dir_global, exist_ok=True)

_work_dir = os.path.join(_results_dir_global, f"llmtrans_tests")
_converted_dir = os.path.join(_work_dir, "converted")
os.makedirs(_converted_dir, exist_ok=True)

_file_name = "691247.docx"

def extract_docx(docx_path, extract_dir):
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        zip_ref.extractall(extract_dir)

def translate_paragraphs_preserve_structure(xml_content, xml_contents=None):

    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    ET.register_namespace('w', namespaces['w'])  # Register namespace to avoid auto-generated prefixes

    root = ET.fromstring(xml_content)

    # Parse the main XML content and other language XML contents
    # root = ET.fromstring(xml_content)
    # namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

    # Parse XML contents for different languages
    root_en = ET.fromstring(xml_contents["en"])
    root_de = ET.fromstring(xml_contents["de"])
    root_it = ET.fromstring(xml_contents["it"])
    root_es = ET.fromstring(xml_contents["es"])
    root_pt = ET.fromstring(xml_contents["pt"])

    # Dictionary to map language codes to XML roots
    language_roots = {
        "en": root_en,
        "de": root_de,
        "it": root_it,
        "es": root_es,
        "pt": root_pt
    }

    # Get all paragraphs from the original XML content
    paragraphs = root.findall('.//w:p', namespaces)
    
    # Get the corresponding paragraphs from each language's XML content
    lang_paragraphs = {
        lang_code: lang_root.findall('.//w:p', namespaces) for lang_code, lang_root in language_roots.items()
    }

    # Iterate over paragraphs in the main XML content
    for paragraph_idx, paragraph in enumerate(paragraphs):
        segment_data = []  # Store text and footnote references as segments
        text_nodes = []    # Store original text nodes for later use

        # Loop over nodes to capture text and footnote markers
        for node in paragraph.iter():
            if node.tag.endswith('}t') and node.text:
                # Accumulate text for translation
                text_nodes.append(node)
                # segment_data.append((node.text, 'text'))
                segment_data.append((node.text or "", 'text'))
            elif node.tag.endswith('}footnoteRef') or node.tag.endswith('}endnoteRef'):
                # Preserve footnote references in the segment data without translation
                segment_data.append((node, 'ref'))

        # Prepare translations, ensuring footnotes are preserved in their positions
        translations = {lang_code: [] for lang_code in language_roots}
        for text, typ in segment_data:
            if typ == 'text':
                if is_url(text):  # Skip URLs from translation
                    for lang_code in translations:
                        translations[lang_code].append([text])
                else:
        
                    translated_text = translate_text(text)

                    for lang_code in translations:
                        translations[lang_code].append([translated_text[lang_code]])
            else:
                # Append footnote reference elements directly without modification
                for lang_code in translations:
                    translations[lang_code].append([text])

        # Reinsert translated text while preserving structure and footnote positions
        for lang_code, lang_paragraph_list in lang_paragraphs.items():
            lang_paragraph = lang_paragraph_list[paragraph_idx]

            # lang_text_nodes = lang_paragraph.findall('.//w:t', namespaces)
            lang_text_nodes = []
            for node in lang_paragraph.iter():
                if node.tag.endswith('}t') and node.text:  # Include all <w:t> nodes, even empty ones
                    lang_text_nodes.append(node)
                elif node.tag.endswith('}footnoteRef') or node.tag.endswith('}endnoteRef'):
                    # Footnotes or endnotes might need to be handled in their own context
                    lang_text_nodes.append(node)  # Append to maintain structural consistency


            idx = 0  # Tracks the position in the translation list

            # Flattened translation list for each language
            lang_translation = translations[lang_code]

            # Reinsert the translated text, distributing across nodes as needed
            for lang_node in lang_text_nodes:
                
                if isinstance(lang_translation[idx][0], str):
                        # Determine how many words to insert into this node
                    # Insert text segments
                    original_text = lang_node.text or ""
                    word_count = len(original_text.split())
                    
                    # Collect exactly 'word_count' words, or as many as remain in the translation
                    translated_segment = lang_translation[idx]
                    lang_node.text = " ".join(translated_segment)

                    if original_text.endswith(" "):
                        lang_node.text += " "

                elif isinstance(lang_translation[idx][0], ET.Element):
                    lang_node.text = " "
            
                # Update idx by the number of words we actually used
                idx += 1
                

    # Return updated XML for all languages as a dictionary
    return {
        "en": ET.tostring(root_en, encoding='unicode'),
        "de": ET.tostring(root_de, encoding='unicode'),
        "it": ET.tostring(root_it, encoding='unicode'),
        "es": ET.tostring(root_es, encoding='unicode'),
        "pt": ET.tostring(root_pt, encoding='unicode'),
    }

# Function to translate headers, footers, and footnotes if they exist
def translate_headers_footers_footnotes(extract_dir, extract_dirs={}):
    
    word_dir = os.path.join(extract_dir, 'word')

    word_dirs = {}
    for key, value in extract_dirs.items():
        word_dirs[key] = os.path.join(extract_dirs[key], 'word')

    header_footer_files = [f for f in os.listdir(word_dir) if f.startswith('header') or f.startswith('footer')]

    # Translate headers and footers
    for file_name in header_footer_files:
        file_path = os.path.join(word_dir, file_name)
        if os.path.exists(file_path):
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            contents = {}
            for key, value in word_dirs.items():
                with open(os.path.join(value, file_name), 'r', encoding='utf-8') as file:
                    contents[key] = file.read()

            translated_contents = translate_paragraphs_preserve_structure(content, contents)
            for key, value in word_dirs.items():
                with open(os.path.join(value, file_name), 'w', encoding='utf-8') as file:
                    file.write(translated_contents[key])

    # Translate footnotes
    footnotes_file = os.path.join(word_dir, 'footnotes.xml')
    if os.path.exists(footnotes_file):
        with open(footnotes_file, 'r', encoding='utf-8') as file:
            footnotes_content = file.read()

        footnotes_contents = {}
        for key, value in word_dirs.items():
            with open(os.path.join(value, 'footnotes.xml'), 'r', encoding='utf-8') as file:
                footnotes_contents[key] = file.read()

        translated_footnotes = translate_paragraphs_preserve_structure(footnotes_content, footnotes_contents)
        for key, value in word_dirs.items():
            with open(os.path.join(value, 'footnotes.xml'), 'w', encoding='utf-8') as file:
                file.write(translated_footnotes[key])

# Main function to handle translation of a .docx file
def analyze_and_translate_docx(input_docx, output_docxs):

    input_file_name_without_extension = os.path.splitext(os.path.basename(input_docx))[0]

    # extract_dir = 'extracted_docx'
    extract_dir = os.path.join(_work_dir, f"extracted_{input_file_name_without_extension}")

    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)
    os.makedirs(extract_dir)

    extract_dirs = {
        "en" : os.path.join(extract_dir, '_en'),
        "de" : os.path.join(extract_dir, '_de'),
        "it" : os.path.join(extract_dir, '_it'),
        "es" : os.path.join(extract_dir, '_es'),
        "pt" : os.path.join(extract_dir, '_pt'),
    }

    for key, value in extract_dirs.items():
        if os.path.exists(value):
            shutil.rmtree(value)
        os.makedirs(value)
        extract_docx(input_docx, value)
    
    extract_docx(input_docx, extract_dir)

    document_xml_path = os.path.join(extract_dir, 'word/document.xml')

    document_xml_paths = {}
    for key, value in extract_dirs.items():
        document_xml_paths[key] = os.path.join(extract_dirs[key], 'word/document.xml')


    if os.path.exists(document_xml_path):
        with open(document_xml_path, 'r', encoding='utf-8') as file:
            document_xml_content = file.read()

        document_xml_contents = {}

        for key, value in document_xml_paths.items():
            with open(value, 'r', encoding='utf-8') as file:
                document_xml_contents[key] = file.read()

        translated_xml_contents = translate_paragraphs_preserve_structure(
            document_xml_content,
            document_xml_contents,
        )
        for key, value in document_xml_paths.items():
            with open(value, 'w', encoding='utf-8') as file:
                file.write(translated_xml_contents[key])
        

    translate_headers_footers_footnotes(extract_dir, extract_dirs)

    for key, value in extract_dirs.items():
        recreate_docx(value, output_docxs[key])
        shutil.rmtree(value)

# Translation function
def translate_text(text):

    # print("---------------translate_text-----------------------")
    # print(text+"*")

    # if not text.strip() =='':
    if contains_any_language_alpha(text):
    # if is_non_alphanumeric(text):
        special_chars = ''
        for char in text:
            if char.isalpha() or char.isdigit(): 
                break
            special_chars += char
    
        text = text[len(special_chars):]  # Get the rest of the string after special characters

        prompt = build_trans_prompt()
        _prompt = prompt

        ai_response = extract_json_from_llm_result(trans_with_sambanova(prompt, text))

        json_data = ai_response

        print("--------json result---------")
        print(json_data)

        traslated_text = {
            "en" : special_chars + json_data.get("EN_RESULT", ""),
            "de" : special_chars + json_data.get("DE_RESULT", ""),
            "it" : special_chars + json_data.get("IT_RESULT", ""),
            "es" : special_chars + json_data.get("ES_RESULT", ""),
            "pt" : special_chars + json_data.get("PT_RESULT", "")
        }

        # traslated_text = {
        #     "en" : special_chars + text + "++",
        #     "de" : special_chars + text + "++",
        #     "it" : special_chars + text + "++",
        #     "es" : special_chars + text + "++",
        #     "pt" : special_chars + text + "++"
        # }

        return traslated_text
    else:
        traslated_text = {
                "en" : text,
                "de" : text,
                "it" : text,
                "es" : text,
                "pt" : text
        }
        return traslated_text


def trans_with_sambanova(prompt, text):

    return json.dumps({
                "EN_RESULT": text + "++",
                "DE_RESULT": text + "++",
                "IT_RESULT": text + "++",
                "ES_RESULT": text + "++",
                "PT_RESULT": text + "++"
            })      

    # return text + "++"

def translate_document():
    source_file_path = os.path.join('uploads', _file_name)
    source_filename_without_extension = os.path.splitext(_file_name)[0]
    converted_file_name = source_filename_without_extension + '.docx'
    converted_file_path = os.path.join(_converted_dir, converted_file_name)
    conversion_command = f'libreoffice --headless --convert-to docx:"MS Word 2007 XML" {os.path.abspath(source_file_path)} --outdir {_converted_dir}'
    print("Executing conversion command:", conversion_command)
    try:
        with _libreoffice_lock:
            subprocess.run(conversion_command, check=True, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
    except subprocess.CalledProcessError as e:
        print("An error occurred during file conversion:", e)
        print("Conversion_command : ", conversion_command)
        raise RuntimeError(f"File conversion failed: {e.stderr}") from e
    output_paths = {
        "en": os.path.join(_results_dir_global, source_filename_without_extension + '_en.docx'),
        "de": os.path.join(_results_dir_global, source_filename_without_extension + '_de.docx'),
        "it": os.path.join(_results_dir_global, source_filename_without_extension + '_it.docx'),
        "es": os.path.join(_results_dir_global, source_filename_without_extension + '_es.docx'),
        "pt": os.path.join(_results_dir_global, source_filename_without_extension + '_pt.docx'),
    }
    analyze_and_translate_docx(converted_file_path, output_paths)   
    output_filenames = {k: os.path.basename(v) for k, v in output_paths.items()}
    return output_filenames

def contains_any_language_alpha(text):
    # \p{L} matches any kind of letter from any language
    return bool(regex.search(r"\p{L}", text))

def is_special_or_space(char):
    if not char:  # Check if the string is empty
        return False
    # Check if the first character is a special character (including space)
    if char in string.punctuation or char.isspace():
        return True

def is_non_alphanumeric(text):
    non_representable_chars = []
    
    for char in text:
        # Get Unicode category of the character
        char_category = unicodedata.category(char)
        
        # Check if it's NOT alphabetic or numeric
        if not (
            char_category.startswith("L")  # Letter
            or char_category.startswith("N")  # Number
        ):
            non_representable_chars.append(char)
    
    # return non_representable_chars
    if non_representable_chars:
        return True
    else:
        return False

def is_url(text):
    try:
        result = urlparse(text)
        # Check if scheme and netloc are present and if the netloc includes at least one period
        return all([result.scheme, result.netloc]) and '.' in result.netloc
    except ValueError:
        return False

# Function to append logs to the end of the document
def append_log_to_document(document_xml_path, log_text, document_xml_paths={}):

    for key, value in document_xml_paths.items():
        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

        # Read and parse the existing document.xml content
        tree = ET.parse(value)
        root = tree.getroot()

        # Create a new paragraph element to hold the log
        new_paragraph = ET.Element(f"{{{namespaces['w']}}}p")
        run = ET.SubElement(new_paragraph, f"{{{namespaces['w']}}}r")
        text_element = ET.SubElement(run, f"{{{namespaces['w']}}}t")
        text_element.text = log_text

        # Append the new paragraph to the document body
        body = root.find(f".//w:body", namespaces)
        body.append(new_paragraph)

        # Write the updated content back to document.xml
        tree.write(value, xml_declaration=True, encoding='utf-8')

def build_trans_prompt():

    prompt = """
        Translate the provided text into English, German, Italian, Spanish, and Portuguese. Follow these specific guidelines:
        1. Translate all elements of the text, including titles, headers, labels, and any formatted components, ensuring their exact structure and formatting are maintained in each target language.
        2. Preserve specific terms, names, numbers, symbols, or unique identifiers exactly as they appear in the source text. Any content that should remain unchanged must not be translated.
        3. Ensure sentence-by-sentence correspondence between the source text and the target translations. Each sentence in the source text must have a direct equivalent in each target language.
        4. Maintain all punctuation, capitalization, and specific phrase structures to accurately retain the original meaning.
        5. Consistently translate technical or domain-specific terminology across all target languages to ensure coherence and accuracy.
        6. Encode all output in Unicode and ensure proper character escaping is applied where necessary.
        7. Provide only the translations, without adding explanations, comments, or additional content.
        8. Return the translated text in this JSON format:
        {
        "EN_RESULT": "TRANSLATED TEXT",
        "DE_RESULT": "TRANSLATED TEXT",
        "IT_RESULT": "TRANSLATED TEXT",
        "ES_RESULT": "TRANSLATED TEXT",
        "PT_RESULT": "TRANSLATED TEXT"
        }
    """

    return prompt


# Function to recreate a .docx file from the extracted content
def recreate_docx(original_extract_dir, new_docx_path):
    
    # Generate temporary DOCX file name
    temp_file_name = os.path.basename(new_docx_path)
    
    # Create the pre-generated .docx file
    with zipfile.ZipFile(temp_file_name, 'w') as docx:
        for foldername, subfolders, filenames in os.walk(original_extract_dir):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                arcname = os.path.relpath(file_path, original_extract_dir)
                docx.write(file_path, arcname)
    
    # Ensure the /result directory exists
    # result_dir = os.path.abspath('results')
    result_dir = _results_dir_global

    print("--temp_file_name path : ", temp_file_name)
    print("--result path : ", result_dir)
    os.makedirs(result_dir, exist_ok=True)
    
    # Convert the file using LibreOffice
    conversion_command = (
        f'libreoffice --headless --convert-to docx:"MS Word 2007 XML" '
        f'{os.path.abspath(temp_file_name)} --outdir {result_dir}'
    )
    
    try:
        with _libreoffice_lock:
            subprocess.run(conversion_command, check=True, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        # print(f"Recreate_docx successfully. Saved to {result_dir}")
    except subprocess.CalledProcessError as e:
        print("An error occurred during recreate_docx", e.stderr)
        raise RuntimeError(f"Recreate_docx failed: {e.stderr}") from e
    
    # Remove the pre-generated .docx file
    if os.path.exists(temp_file_name):
        os.remove(temp_file_name)
        # print(f"Temporary file {temp_file_name} removed.")


def decode_unicode_escapes(input_string):
    # Handle double-escaped unicode sequences first
    while '\\\\u' in input_string:
        pos = input_string.find('\\\\u')
        # Ensure the escape is followed by exactly four hexadecimal digits
        if pos != -1 and len(input_string) >= pos + 7 and all(c in '0123456789abcdefABCDEF' for c in input_string[pos+3:pos+7]):
            char = chr(int(input_string[pos+3:pos+7], 16))
            input_string = input_string[:pos] + char + input_string[pos+7:]
        else:
            break  # No valid double-escaped sequence found, break the loop

    # Then handle normal escaped unicode sequences
    while '\\u' in input_string:
        pos = input_string.find('\\u')
        # Ensure the escape is followed by exactly four hexadecimal digits
        if pos != -1 and len(input_string) >= pos + 6 and all(c in '0123456789abcdefABCDEF' for c in input_string[pos+2:pos+6]):
            char = chr(int(input_string[pos+2:pos+6], 16))
            input_string = input_string[:pos] + char + input_string[pos+6:]
        else:
            break  # No valid escape sequence found, break the loop

    return input_string



def extract_json_from_llm_result(input_string):

    input_string = re.sub(r'<text\|begin>', '', input_string)
    input_string = re.sub(r'<text\|end>', '', input_string)

    keys = ["EN_RESULT", "DE_RESULT", "IT_RESULT", "ES_RESULT", "PT_RESULT"]
    # Try to extract the JSON block containing all keys
    pattern = r'\{[^\{\}]*"EN_RESULT"[\s\S]*?"PT_RESULT"[\s\S]*?\}'
    match = re.search(pattern, input_string)
    if match:
        json_str = match.group(0)
        try:
            data = json.loads(json_str)
            # Optionally decode unicode escapes if needed
            for k, v in data.items():
                data[k] = decode_unicode_escapes(html.unescape(v))
            return data
        except Exception:
            pass  # fallback below

    # extract only the translation block (EN_RESULT to PT_RESULT, stop at first empty line or end)
    try:
        # Find the block from EN_RESULT: ... to PT_RESULT: ... (stop at first empty line or end after PT_RESULT)
        block_pattern = r'(EN_RESULT:.*?PT_RESULT:.*?)(?:\n\s*\n|$)'
        block_match = re.search(block_pattern, input_string, re.DOTALL)
        if block_match:
            translation_block = block_match.group(1)
        else:
            translation_block = input_string  # fallback to whole string if not found
        key_pattern = r'(EN_RESULT|DE_RESULT|IT_RESULT|ES_RESULT|PT_RESULT):'
        matches = list(re.finditer(key_pattern, translation_block))
        if matches and len(matches) >= 3:  # at least 3 keys found, likely a match
            extracted_values = {}
            for i, match in enumerate(matches):
                key = match.group(1)
                value_start = match.end()
                value_end = matches[i + 1].start() if i + 1 < len(matches) else len(translation_block)
                value = translation_block[value_start:value_end].strip(' ,{}"\n')
                value = decode_unicode_escapes(html.unescape(value))
                extracted_values[key] = value
            # Fill missing keys with empty string
            for key in keys:
                if key not in extracted_values:
                    extracted_values[key] = ''
            return extracted_values
    except Exception:
        pass  # fallback to next solution

    # Fallback: try to parse as JSON from the start (may work if input is just JSON)
    try:
        data = json.loads(input_string)
        for k, v in data.items():
            data[k] = decode_unicode_escapes(html.unescape(v))
        return data
    except Exception:
        pass
    
    # Fallback: old method (fragile, but last resort)
    positions = {key: input_string.find(f'"{key}":') for key in keys}
    extracted_values = {}
    for i, key in enumerate(keys):
        start = positions[key] + len(f'"{key}": ')
        end = positions[keys[i + 1]] if i + 1 < len(keys) else len(input_string)
        text = input_string[start:end].strip(' ,{}"\n')
        text = decode_unicode_escapes(html.unescape(text))
        extracted_values[key] = text
    return extracted_values

translate_document()