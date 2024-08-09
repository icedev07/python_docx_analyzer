from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
import os

# Load the original document
doc = Document('input.docx')

# Create a new document
new_doc = Document()

# Directory to save images
image_dir = "extracted_images"
os.makedirs(image_dir, exist_ok=True)

# Function to copy run properties, including footnote references
def copy_run_properties(source_run, target_run):
    target_run.bold = source_run.bold
    target_run.italic = source_run.italic
    target_run.underline = source_run.underline
    target_run.font.name = source_run.font.name
    target_run.font.size = source_run.font.size
    if source_run.font.color and source_run.font.color.rgb:
        target_run.font.color.rgb = source_run.font.color.rgb
    target_run.font.highlight_color = source_run.font.highlight_color
    target_run.font.superscript = source_run.font.superscript
    target_run.font.subscript = source_run.font.subscript
    target_run.font.strike = source_run.font.strike
    target_run.font.all_caps = source_run.font.all_caps
    
    # Copy footnote references using XML manipulation
    for element in source_run._element:
        if element.tag.endswith('footnoteReference'):
            footnote_id = element.get(qn('w:id'))
            create_footnote_reference(target_run, footnote_id)

def create_footnote_reference(run, footnote_id):
    footnote_ref = OxmlElement('w:footnoteReference')
    footnote_ref.set(qn('w:id'), footnote_id)
    run._r.append(footnote_ref)

# Function to copy paragraph properties, including list formatting
def copy_paragraph_properties(source_paragraph, target_paragraph):
    target_paragraph.style = source_paragraph.style
    target_paragraph.alignment = source_paragraph.alignment
    target_paragraph.paragraph_format.left_indent = source_paragraph.paragraph_format.left_indent
    target_paragraph.paragraph_format.right_indent = source_paragraph.paragraph_format.right_indent
    target_paragraph.paragraph_format.space_before = source_paragraph.paragraph_format.space_before
    target_paragraph.paragraph_format.space_after = source_paragraph.paragraph_format.space_after
    target_paragraph.paragraph_format.line_spacing = source_paragraph.paragraph_format.line_spacing
    target_paragraph.paragraph_format.keep_together = source_paragraph.paragraph_format.keep_together
    target_paragraph.paragraph_format.keep_with_next = source_paragraph.paragraph_format.keep_with_next
    target_paragraph.paragraph_format.page_break_before = source_paragraph.paragraph_format.page_break_before
    target_paragraph.paragraph_format.widow_control = source_paragraph.paragraph_format.widow_control
    target_paragraph.paragraph_format.first_line_indent = source_paragraph.paragraph_format.first_line_indent

    # Preserve bullet and numbering format
    if source_paragraph.style.name in ['List Bullet', 'List Number']:
        target_paragraph.style = source_paragraph.style

# Function to copy images within a paragraph
def copy_images(source_run, target_run):
    nsmap = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}
    
    inline_shapes = source_run._element.findall('.//a:blip', namespaces=nsmap)
    
    if inline_shapes:
        for inline_shape in inline_shapes:
            rId = inline_shape.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if rId:
                try:
                    image_part = doc.part.related_parts[rId]
                    image_name = os.path.basename(image_part.partname)
                    image_path = os.path.join(image_dir, image_name)
                    with open(image_path, 'wb') as img_file:
                        img_file.write(image_part.blob)

                    # Try to get the original size from the shape
                    ext_elements = source_run._element.xpath('.//a:ext')
                    if ext_elements:
                        ext = ext_elements[0]
                        cx = ext.get('cx')
                        cy = ext.get('cy')
                        if cx and cy:
                            cx = int(cx) / 914400  # EMU to inches
                            cy = int(cy) / 914400  # EMU to inches
                            target_run.add_picture(image_path, width=Pt(cx * 72), height=Pt(cy * 72))
                        else:
                            target_run.add_picture(image_path)
                    else:
                        target_run.add_picture(image_path)
                except Exception as e:
                    print(f"Error copying image: {e}")
    else:
        pass

# Function to copy paragraphs, including runs, footnotes, and images
def copy_paragraph(paragraph, new_paragraph):
    copy_paragraph_properties(paragraph, new_paragraph)

    for run in paragraph.runs:
        new_run = new_paragraph.add_run(run.text)
        copy_run_properties(run, new_run)
        copy_images(run, new_run)

# Function to copy headers and footers, including images
def copy_headers_footers(old_section, new_section):
    if old_section.header:
        header = old_section.header
        new_header = new_section.header
        new_header.is_linked_to_previous = False
        for paragraph in header.paragraphs:
            new_paragraph = new_header.add_paragraph()
            copy_paragraph(paragraph, new_paragraph)

    if old_section.footer:
        footer = old_section.footer
        new_footer = new_section.footer
        new_footer.is_linked_to_previous = False
        for paragraph in footer.paragraphs:
            new_paragraph = new_footer.add_paragraph()
            copy_paragraph(paragraph, new_paragraph)

# Function to copy footnotes, including content and references
def copy_footnotes(old_doc, new_doc):
    try:
        old_footnotes = old_doc.part.footnotes
        new_footnotes = new_doc.part.footnotes

        for footnote in old_footnotes.element.findall('.//w:footnote', namespaces=old_footnotes.element.nsmap):
            new_footnote = new_footnotes._element.addfootnote()
            new_footnote.set(qn('w:id'), footnote.get(qn('w:id')))  # Ensure the IDs match
            for paragraph in footnote.findall('.//w:p', namespaces=old_footnotes.element.nsmap):
                new_paragraph = new_footnote.add_paragraph()
                copy_paragraph(paragraph, new_paragraph)
    except AttributeError:
        print("No footnotes found in the original document.")
        pass

# Function to copy page size, orientation, and margins
def copy_page_setup(old_section, new_section):
    new_section.page_width = old_section.page_width
    new_section.page_height = old_section.page_height
    new_section.orientation = old_section.orientation
    new_section.left_margin = old_section.left_margin
    new_section.right_margin = old_section.right_margin
    new_section.top_margin = old_section.top_margin
    new_section.bottom_margin = old_section.bottom_margin
    new_section.header_distance = old_section.header_distance
    new_section.footer_distance = old_section.footer_distance
    new_section.gutter = old_section.gutter

# Copy paragraphs, maintaining their properties
for paragraph in doc.paragraphs:
    new_paragraph = new_doc.add_paragraph()
    copy_paragraph(paragraph, new_paragraph)

# Copy headers, footers, and page setup
for section_index, section in enumerate(doc.sections):
    if section_index == 0:
        new_section = new_doc.sections[0]
    else:
        new_section = new_doc.add_section()

    copy_page_setup(section, new_section)
    copy_headers_footers(section, new_section)

# Copy footnotes if they exist
copy_footnotes(doc, new_doc)

# Save the new document
new_doc.save('result.docx')
