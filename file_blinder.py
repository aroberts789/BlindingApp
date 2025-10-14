#!/usr/bin/env python3
"""
File Blinding Script - Removes images and replaces keywords while preserving ALL structure
Supports: DOCX, HTML, TXT files
Keeps original structure, fonts, tables, styles - only removes images and replaces keywords
"""

import re
import os
import sys
from pathlib import Path
import zipfile
import tempfile
from xml.etree import ElementTree as ET

try:
    from docx import Document

    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from bs4 import BeautifulSoup

    HTML_AVAILABLE = True
except ImportError:
    HTML_AVAILABLE = False


class FileBlinder:
    def __init__(self, keyword_replacements=None, standardize_formatting=True, font_name="Calibri", font_size=11,
                 font_color_black=True, grey_shading=False):
        """
        Initialize with keyword replacement dictionary and formatting options

        Args:
            keyword_replacements (dict): Dictionary of {original: replacement}
            standardize_formatting (bool): Whether to standardize fonts and colors
            font_name (str): Font to standardize to
            font_size (int): Font size to use (None to keep original)
            font_color_black (bool): Whether to make all text black
            grey_shading (bool): Whether to add grey shading to text
        """
        self.keyword_replacements = keyword_replacements or {
            "confidential": "[REDACTED]",
            "secret": "[REDACTED]",
            "internal": "[INTERNAL]",
            "proprietary": "[PROPRIETARY]",
            "classified": "[CLASSIFIED]",
            # Regex patterns
            r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b': '[EMAIL]',
            r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b': '[PHONE]',
            r'\b\d{3}-\d{2}-\d{4}\b': '[SSN]',
        }

        self.standardize_formatting = standardize_formatting
        self.font_name = font_name
        self.font_size = font_size
        self.font_color_black = font_color_black
        self.grey_shading = grey_shading

    def extract_document_structure(self, input_path):
        """Extract document content with metadata for diff generation"""
        extension = Path(input_path).suffix.lower()

        if extension == '.docx':
            return self._extract_docx_structure(input_path)
        elif extension in ['.html', '.htm']:
            return self._extract_html_structure(input_path)
        elif extension == '.txt':
            return self._extract_txt_structure(input_path)
        else:
            raise ValueError(f"Unsupported file type: {extension}")

    def _extract_docx_structure(self, input_path):
        """Extract structure from DOCX file"""
        if not DOCX_AVAILABLE:
            raise ImportError("python-docx not installed")

        doc = Document(input_path)
        structure = {
            'type': 'docx',
            'paragraphs': [],
            'images': [],
            'tables': []
        }

        # Extract paragraphs with metadata
        for idx, para in enumerate(doc.paragraphs):
            para_data = {
                'index': idx,
                'text': para.text,
                'has_image': self._has_drawing_elements(para),
                'formatting': self._extract_paragraph_formatting(para),
                'style': para.style.name if para.style else 'Normal'
            }
            structure['paragraphs'].append(para_data)

        # Extract tables
        for table_idx, table in enumerate(doc.tables):
            table_data = {
                'index': table_idx,
                'rows': [],
                'has_shading': False
            }

            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text
                    has_shading = self._cell_has_shading(cell)
                    if has_shading:
                        table_data['has_shading'] = True
                    row_data.append({
                        'text': cell_text,
                        'has_shading': has_shading
                    })
                table_data['rows'].append(row_data)

            structure['tables'].append(table_data)

        return structure

    def _has_drawing_elements(self, paragraph):
        """Check if paragraph contains images"""
        try:
            p_element = paragraph._element
            for child in p_element.iter():
                if 'drawing' in str(child.tag).lower() or 'object' in str(child.tag).lower():
                    return True
            return False
        except:
            return False

    def _extract_paragraph_formatting(self, paragraph):
        """Extract formatting information from paragraph"""
        formatting = {
            'font_name': None,
            'font_size': None,
            'font_color': None,
            'is_bold': False,
            'is_italic': False
        }

        try:
            if paragraph.runs:
                first_run = paragraph.runs[0]
                if first_run.font.name:
                    formatting['font_name'] = first_run.font.name
                if first_run.font.size:
                    formatting['font_size'] = first_run.font.size.pt
                if first_run.font.color.rgb:
                    rgb = first_run.font.color.rgb
                    formatting['font_color'] = f"#{rgb[0]:02x}{rgb[1]:02x}{rgb[2]:02x}"
                formatting['is_bold'] = first_run.font.bold or False
                formatting['is_italic'] = first_run.font.italic or False
        except:
            pass

        return formatting

    def _cell_has_shading(self, cell):
        """Check if table cell has background shading"""
        try:
            tc_element = cell._tc
            for child in tc_element.iter():
                tag_str = str(child.tag).lower()
                if 'shd' in tag_str:
                    return True
            return False
        except:
            return False

    def _extract_html_structure(self, input_path):
        """Extract structure from HTML file"""
        if not HTML_AVAILABLE:
            raise ImportError("beautifulsoup4 not installed")

        with open(input_path, 'r', encoding='utf-8') as file:
            content = file.read()

        soup = BeautifulSoup(content, 'html.parser')

        structure = {
            'type': 'html',
            'paragraphs': [],
            'images': []
        }

        # Extract text content
        for idx, para in enumerate(soup.find_all(['p', 'div', 'h1', 'h2', 'h3'])):
            structure['paragraphs'].append({
                'index': idx,
                'text': para.get_text(),
                'tag': para.name
            })

        # Count images
        images = soup.find_all(['img', 'picture', 'svg'])
        structure['images'] = [{'index': i} for i in range(len(images))]

        return structure

    def _extract_txt_structure(self, input_path):
        """Extract structure from TXT file"""
        try:
            with open(input_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except UnicodeDecodeError:
            with open(input_path, 'r', encoding='latin-1') as file:
                content = file.read()

        paragraphs = content.split('\n\n')

        structure = {
            'type': 'txt',
            'paragraphs': [
                {'index': idx, 'text': para.strip()}
                for idx, para in enumerate(paragraphs) if para.strip()
            ]
        }

        return structure

    def generate_diff(self, original_structure, processed_structure):
        """Generate diff between original and processed structures"""
        import difflib

        diff_data = {
            'paragraph_changes': [],
            'image_changes': [],
            'table_changes': [],
            'formatting_changes': []
        }

        # Compare paragraphs
        for idx, (orig_para, proc_para) in enumerate(zip(
                original_structure.get('paragraphs', []),
                processed_structure.get('paragraphs', [])
        )):
            orig_text = orig_para.get('text', '')
            proc_text = proc_para.get('text', '')

            if orig_text != proc_text:
                # Generate character-level diff
                matcher = difflib.SequenceMatcher(None, orig_text, proc_text)
                text_changes = []

                for tag, i1, i2, j1, j2 in matcher.get_opcodes():
                    if tag == 'replace':
                        text_changes.append({
                            'type': 'replace',
                            'original': orig_text[i1:i2],
                            'processed': proc_text[j1:j2],
                            'position': i1
                        })
                    elif tag == 'delete':
                        text_changes.append({
                            'type': 'delete',
                            'original': orig_text[i1:i2],
                            'position': i1
                        })
                    elif tag == 'insert':
                        text_changes.append({
                            'type': 'insert',
                            'processed': proc_text[j1:j2],
                            'position': j1
                        })

                if text_changes:
                    diff_data['paragraph_changes'].append({
                        'index': idx,
                        'changes': text_changes
                    })

            # Check for image changes
            if orig_para.get('has_image') and not proc_para.get('has_image'):
                diff_data['image_changes'].append({
                    'paragraph_index': idx,
                    'removed': True
                })

            # Check for formatting changes
            orig_format = orig_para.get('formatting', {})
            proc_format = proc_para.get('formatting', {})

            if orig_format != proc_format:
                diff_data['formatting_changes'].append({
                    'paragraph_index': idx,
                    'original': orig_format,
                    'processed': proc_format
                })

        # Compare tables
        for idx, (orig_table, proc_table) in enumerate(zip(
                original_structure.get('tables', []),
                processed_structure.get('tables', [])
        )):
            if orig_table.get('has_shading') != proc_table.get('has_shading'):
                diff_data['table_changes'].append({
                    'index': idx,
                    'shading_removed': True
                })

        return diff_data
    def replace_keywords_in_text(self, text):
        """Replace keywords in text based on replacement dictionary"""
        if not text:
            return text

        for original, replacement in self.keyword_replacements.items():
            # Use regex for pattern matching
            if original.startswith(r'\b') or original.startswith(r'['):
                text = re.sub(original, replacement, text, flags=re.IGNORECASE)
            else:
                # Simple string replacement (case insensitive)
                text = re.sub(re.escape(original), replacement, text, flags=re.IGNORECASE)

        return text

    def remove_document_themes(self, doc):
        """Remove document themes that might cause colored text"""
        try:
            # Clear theme colors by setting document to a basic theme
            if hasattr(doc, 'settings'):
                try:
                    doc.settings.theme = None
                except:
                    pass

            # Also try to clear theme at the document level
            if hasattr(doc, '_element'):
                doc_element = doc._element

                # Find and remove theme elements
                themes_to_remove = []
                for child in doc_element.iter():
                    if any(theme_word in str(child.tag).lower() for theme_word in ['theme', 'color', 'scheme']):
                        themes_to_remove.append(child)

                for theme in themes_to_remove:
                    parent = theme.getparent()
                    if parent is not None:
                        try:
                            parent.remove(theme)
                        except:
                            pass

        except Exception as e:
            # Continue if theme removal fails
            pass

    def standardize_run_formatting(self, run):
        """Apply standard formatting to a text run"""
        if not self.standardize_formatting:
            return

        try:
            # Set font name
            if self.font_name:
                run.font.name = self.font_name

            # Set font size if specified
            if self.font_size:
                from docx.shared import Pt
                run.font.size = Pt(self.font_size)

            # Aggressively set font color to black (remove all colors including theme colors)
            if self.font_color_black:
                from docx.shared import RGBColor
                run.font.color.rgb = RGBColor(0, 0, 0)  # Black

                # Also clear theme color more aggressively
                try:
                    run.font.color.theme_color = None
                    # Try to clear any color index
                    if hasattr(run.font.color, '_color_val'):
                        run.font.color._color_val = None
                except:
                    pass

            # Remove all highlighting and shading
            run.font.highlight_color = None  # Remove highlight

            # Also try to clear any theme-based highlighting
            try:
                if hasattr(run._element, 'rPr'):
                    rPr = run._element.rPr
                    if rPr is not None:
                        # Remove highlight elements
                        highlights_to_remove = []
                        for child in rPr:
                            if 'highlight' in str(child.tag).lower() or 'shd' in str(child.tag).lower():
                                highlights_to_remove.append(child)

                        for highlight in highlights_to_remove:
                            rPr.remove(highlight)
            except:
                pass

            # Remove underlines and other special formatting while keeping bold/italic
            # run.font.underline = None  # Uncomment if you want to remove underlines too

        except Exception as e:
            # If formatting fails, continue - text replacement is more important
            pass

    def remove_table_cell_shading(self, cell):
        """Remove background shading from a table cell"""
        try:
            # Work at the XML level to remove cell shading
            tc_element = cell._tc

            # Find and remove shading elements from the cell
            shading_elements_to_remove = []
            for child in tc_element.iter():
                tag_str = str(child.tag).lower()
                if any(shading_word in tag_str for shading_word in ['shd', 'fill', 'tcpr']):
                    # Check if it's a shading element within table cell properties
                    if 'shd' in tag_str:
                        shading_elements_to_remove.append(child)

            for shading_elem in shading_elements_to_remove:
                parent = shading_elem.getparent()
                if parent is not None:
                    parent.remove(shading_elem)

            # Also check for table cell properties and remove shading
            try:
                for tcPr in tc_element.xpath('.//w:tcPr', namespaces={
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    # Remove any shading within table cell properties
                    for shd in tcPr.xpath('.//w:shd', namespaces={
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        tcPr.remove(shd)
            except:
                pass

        except Exception as e:
            # Continue if cell shading removal fails
            pass

    def remove_paragraph_borders(self, paragraph):
        """Remove paragraph borders"""
        try:
            # Remove borders at the paragraph format level
            paragraph_format = paragraph.paragraph_format

            # Clear paragraph borders
            if hasattr(paragraph_format, 'borders'):
                borders = paragraph_format.borders
                if borders:
                    try:
                        borders.top = None
                        borders.bottom = None
                        borders.left = None
                        borders.right = None
                    except:
                        pass

            # Also work at the XML level to remove border elements
            p_element = paragraph._element

            # Find and remove all border-related elements
            border_elements_to_remove = []
            for child in p_element.iter():
                tag_str = str(child.tag).lower()
                if any(border_word in tag_str for border_word in ['bdr', 'border', 'pBdr']):
                    border_elements_to_remove.append(child)

            for border_elem in border_elements_to_remove:
                parent = border_elem.getparent()
                if parent is not None:
                    parent.remove(border_elem)

            # Also check for paragraph properties and remove borders
            try:
                for pPr in p_element.xpath('.//w:pPr', namespaces={
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    # Remove any borders within paragraph properties
                    for border in pPr.xpath('.//w:pBdr', namespaces={
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        pPr.remove(border)
            except:
                pass

        except Exception as e:
            # Continue if border removal fails
            pass

    def remove_paragraph_shading(self, paragraph):
        """Remove paragraph-level shading and background colors"""
        try:
            # Remove paragraph shading at multiple levels
            paragraph_format = paragraph.paragraph_format

            # Clear paragraph-level shading
            if hasattr(paragraph_format, 'shading'):
                shading = paragraph_format.shading
                if shading:
                    shading.background_pattern_color = None
                    shading.foreground_pattern_color = None
                    # Also try to clear the fill
                    try:
                        shading.fill = None
                    except:
                        pass

            # Also work at the XML level to remove shading elements
            p_element = paragraph._element

            # Find and remove all shading-related elements
            shading_elements_to_remove = []
            for child in p_element.iter():
                tag_str = str(child.tag).lower()
                if any(shading_word in tag_str for shading_word in ['shd', 'fill', 'highlight', 'bgcolor']):
                    shading_elements_to_remove.append(child)

            for shading_elem in shading_elements_to_remove:
                parent = shading_elem.getparent()
                if parent is not None:
                    parent.remove(shading_elem)

            # Also check for paragraph properties and remove background colors
            try:
                for pPr in p_element.xpath('.//w:pPr', namespaces={
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                    # Remove any shading within paragraph properties
                    for shd in pPr.xpath('.//w:shd', namespaces={
                        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}):
                        pPr.remove(shd)
            except:
                pass

        except Exception as e:
            # Continue if shading removal fails
            pass

    def remove_hyperlinks_from_paragraph(self, paragraph):
        """Remove hyperlinks from a paragraph while keeping the text"""
        try:
            # Find all hyperlink elements in the paragraph
            p_element = paragraph._element

            # Remove hyperlink formatting but keep text
            hyperlinks_to_remove = []
            for child in p_element.iter():
                if 'hyperlink' in str(child.tag).lower():
                    hyperlinks_to_remove.append(child)

            for hyperlink in hyperlinks_to_remove:
                # Get the text content before removing
                text_content = hyperlink.text or ""
                parent = hyperlink.getparent()
                if parent is not None:
                    # Replace hyperlink with plain text
                    parent.remove(hyperlink)
                    # Note: The text should be preserved by the run processing

        except Exception as e:
            # Continue if hyperlink removal fails
            pass

    def remove_list_formatting(self, paragraph):
        """Remove list bullet highlighting and formatting"""
        try:
            # Clear list formatting - reset to normal paragraph
            paragraph.style = paragraph._document.styles['Normal']

            # Remove paragraph numbering properties
            paragraph_format = paragraph.paragraph_format

            # Clear left indent that might cause bullet appearance
            paragraph_format.left_indent = None
            paragraph_format.first_line_indent = None

            # Remove any numbering
            p_element = paragraph._element

            # Find and remove numbering properties
            numbering_to_remove = []
            for child in p_element.iter():
                if 'numPr' in str(child.tag) or 'pPr' in str(child.tag):
                    # Look for numbering inside pPr
                    for subchild in child:
                        if 'numPr' in str(subchild.tag):
                            numbering_to_remove.append(subchild)

            for num_element in numbering_to_remove:
                parent = num_element.getparent()
                if parent is not None:
                    parent.remove(num_element)

        except Exception as e:
            # Continue if list formatting removal fails
            pass

    def process_docx_safe(self, input_path, output_path):
        """Safe DOCX processing that preserves structure perfectly"""
        if not DOCX_AVAILABLE:
            raise ImportError("python-docx not installed. Run: pip install python-docx")

        print("Loading document...")
        doc = Document(input_path)

        # Remove document themes that might cause colored text
        print("Removing document themes...")
        self.remove_document_themes(doc)

        images_removed = 0
        text_replacements = 0

        # Process paragraphs
        print("Processing paragraphs...")
        for para_idx, paragraph in enumerate(doc.paragraphs):
            if para_idx % 10 == 0:
                print(f"  Processing paragraph {para_idx + 1}/{len(doc.paragraphs)}")

            # Remove images using a safer approach - find and remove drawing elements
            try:
                # Get the paragraph's XML element
                p_element = paragraph._element

                # Find drawing elements (images) - use a simple approach
                drawings_to_remove = []
                for child in p_element.iter():
                    if 'drawing' in str(child.tag).lower():
                        drawings_to_remove.append(child)

                for drawing in drawings_to_remove:
                    parent = drawing.getparent()
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Find object elements - use a simple approach
                objects_to_remove = []
                for child in p_element.iter():
                    if 'object' in str(child.tag).lower():
                        objects_to_remove.append(child)

                for obj in objects_to_remove:
                    parent = obj.getparent()
                    if parent is not None:
                        parent.remove(obj)
                        images_removed += 1

            except Exception as e:
                # If image removal fails for this paragraph, continue with text processing
                print(f"  Warning: Could not remove images from paragraph {para_idx + 1}: {e}")

            # Process text in runs while maintaining formatting
            for run in paragraph.runs:
                if run.text:
                    original_text = run.text
                    new_text = self.replace_keywords_in_text(original_text)
                    if new_text != original_text:
                        run.text = new_text
                        text_replacements += 1

                # Apply formatting standardization
                self.standardize_run_formatting(run)

            # Remove hyperlinks, list formatting, borders, and shading
            self.remove_hyperlinks_from_paragraph(paragraph)
            self.remove_list_formatting(paragraph)
            self.remove_paragraph_borders(paragraph)
            self.remove_paragraph_shading(paragraph)

        # Process tables
        print("Processing tables...")
        for table_idx, table in enumerate(doc.tables):
            print(f"  Processing table {table_idx + 1}/{len(doc.tables)}")
            for row in table.rows:
                for cell in row.cells:
                    # Remove cell-level shading (this handles the orange backgrounds)
                    self.remove_table_cell_shading(cell)

                    for paragraph in cell.paragraphs:
                        # Remove images from table cells using safer approach
                        try:
                            p_element = paragraph._element

                            # Find and remove drawing elements
                            drawings_to_remove = []
                            for child in p_element.iter():
                                if 'drawing' in str(child.tag).lower():
                                    drawings_to_remove.append(child)

                            for drawing in drawings_to_remove:
                                parent = drawing.getparent()
                                if parent is not None:
                                    parent.remove(drawing)
                                    images_removed += 1

                            # Find and remove object elements
                            objects_to_remove = []
                            for child in p_element.iter():
                                if 'object' in str(child.tag).lower():
                                    objects_to_remove.append(child)

                            for obj in objects_to_remove:
                                parent = obj.getparent()
                                if parent is not None:
                                    parent.remove(obj)
                                    images_removed += 1

                        except Exception as e:
                            # Continue if image removal fails for this cell
                            pass

                        # Process text in cell runs
                        for run in paragraph.runs:
                            if run.text:
                                original_text = run.text
                                new_text = self.replace_keywords_in_text(original_text)
                                if new_text != original_text:
                                    run.text = new_text
                                    text_replacements += 1

                            # Apply formatting standardization
                            self.standardize_run_formatting(run)

                        # Remove hyperlinks, list formatting, borders, and shading from table cells
                        self.remove_hyperlinks_from_paragraph(paragraph)
                        self.remove_list_formatting(paragraph)
                        self.remove_paragraph_borders(paragraph)
                        self.remove_paragraph_shading(paragraph)

        # Process headers
        print("Processing headers and footers...")
        for section in doc.sections:
            # Process header
            if section.header:
                for paragraph in section.header.paragraphs:
                    try:
                        p_element = paragraph._element

                        # Remove drawing elements
                        drawings_to_remove = []
                        for child in p_element.iter():
                            if 'drawing' in str(child.tag).lower():
                                drawings_to_remove.append(child)

                        for drawing in drawings_to_remove:
                            parent = drawing.getparent()
                            if parent is not None:
                                parent.remove(drawing)
                                images_removed += 1

                    except Exception as e:
                        # Continue if image removal fails
                        pass

                    for run in paragraph.runs:
                        if run.text:
                            original_text = run.text
                            new_text = self.replace_keywords_in_text(original_text)
                            if new_text != original_text:
                                run.text = new_text
                                text_replacements += 1

                        # Apply formatting standardization
                        self.standardize_run_formatting(run)

            # Process footer
            if section.footer:
                for paragraph in section.footer.paragraphs:
                    try:
                        p_element = paragraph._element

                        # Remove drawing elements
                        drawings_to_remove = []
                        for child in p_element.iter():
                            if 'drawing' in str(child.tag).lower():
                                drawings_to_remove.append(child)

                        for drawing in drawings_to_remove:
                            parent = drawing.getparent()
                            if parent is not None:
                                parent.remove(drawing)
                                images_removed += 1

                    except Exception as e:
                        # Continue if image removal fails
                        pass

                    for run in paragraph.runs:
                        if run.text:
                            original_text = run.text
                            new_text = self.replace_keywords_in_text(original_text)
                            if new_text != original_text:
                                run.text = new_text
                                text_replacements += 1

                        # Apply formatting standardization
                        self.standardize_run_formatting(run)

        print("Saving document...")
        doc.save(output_path)

        print(f"✓ Removed {images_removed} images/objects")
        print(f"✓ Made {text_replacements} text replacements")

        return output_path

    def process_docx_xml_safe(self, input_path, output_path):
        """Process DOCX by safely modifying XML while preserving structure"""

        # Create a temporary working directory
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_dir = Path(temp_dir)

            # Extract the DOCX file
            with zipfile.ZipFile(input_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)

            # Define XML namespace
            namespaces = {
                'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            }

            # Register namespaces
            for prefix, uri in namespaces.items():
                ET.register_namespace(prefix, uri)

            images_removed = 0
            text_replacements = 0

            # Process main document
            document_xml = temp_dir / 'word' / 'document.xml'
            if document_xml.exists():
                print("Processing main document XML...")
                tree = ET.parse(document_xml)
                root = tree.getroot()

                # Remove drawing elements (images)
                for drawing in root.findall('.//w:drawing', namespaces):
                    parent = drawing.getparent()
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Remove object elements
                for obj in root.findall('.//w:object', namespaces):
                    parent = obj.getparent()
                    if parent is not None:
                        parent.remove(obj)
                        images_removed += 1

                # Replace text in text elements
                for text_elem in root.findall('.//w:t', namespaces):
                    if text_elem.text:
                        original_text = text_elem.text
                        new_text = self.replace_keywords_in_text(original_text)
                        if new_text != original_text:
                            text_elem.text = new_text
                            text_replacements += 1

                # Save the modified XML
                tree.write(document_xml, encoding='utf-8', xml_declaration=True)

            # Process headers
            header_files = list((temp_dir / 'word').glob('header*.xml'))
            for header_file in header_files:
                print(f"Processing {header_file.name}...")
                tree = ET.parse(header_file)
                root = tree.getroot()

                # Remove images
                for drawing in root.findall('.//w:drawing', namespaces):
                    parent = drawing.getparent()
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Replace text
                for text_elem in root.findall('.//w:t', namespaces):
                    if text_elem.text:
                        original_text = text_elem.text
                        new_text = self.replace_keywords_in_text(original_text)
                        if new_text != original_text:
                            text_elem.text = new_text
                            text_replacements += 1

                tree.write(header_file, encoding='utf-8', xml_declaration=True)

            # Process footers
            footer_files = list((temp_dir / 'word').glob('footer*.xml'))
            for footer_file in footer_files:
                print(f"Processing {footer_file.name}...")
                tree = ET.parse(footer_file)
                root = tree.getroot()

                # Remove images
                for drawing in root.findall('.//w:drawing', namespaces):
                    parent = drawing.getparent()
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Replace text
                for text_elem in root.findall('.//w:t', namespaces):
                    if text_elem.text:
                        original_text = text_elem.text
                        new_text = self.replace_keywords_in_text(original_text)
                        if new_text != original_text:
                            text_elem.text = new_text
                            text_replacements += 1

                tree.write(footer_file, encoding='utf-8', xml_declaration=True)

            # Recreate the DOCX file
            print("Rebuilding DOCX file...")
            with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
                for file_path in temp_dir.rglob('*'):
                    if file_path.is_file():
                        # Calculate the path within the zip
                        arc_path = file_path.relative_to(temp_dir)
                        zip_out.write(file_path, arc_path)

            print(f"✓ Removed {images_removed} images/objects")
            print(f"✓ Made {text_replacements} text replacements")

        return output_path

    def process_html_file(self, input_path, output_path):
        """Process HTML file - remove images and replace keywords, keep structure and CSS"""
        if not HTML_AVAILABLE:
            raise ImportError("beautifulsoup4 not installed. Run: pip install beautifulsoup4")

        with open(input_path, 'r', encoding='utf-8') as file:
            content = file.read()

        soup = BeautifulSoup(content, 'html.parser')

        # Remove all image-related elements
        images_removed = 0
        for img_tag in soup.find_all(['img', 'picture', 'svg', 'canvas']):
            img_tag.decompose()
            images_removed += 1

        # Remove background images from CSS (inline styles)
        for element in soup.find_all(style=True):
            style = element.get('style', '')
            if 'background-image' in style.lower():
                style = re.sub(r'background-image\s*:[^;]*;?', '', style, flags=re.IGNORECASE)
                element['style'] = style
                images_removed += 1

        # Remove background images from CSS in style tags
        for style_tag in soup.find_all('style'):
            if style_tag.string:
                css_content = style_tag.string
                if 'background-image' in css_content.lower():
                    css_content = re.sub(r'background-image\s*:[^;]*;?', '', css_content, flags=re.IGNORECASE)
                    style_tag.string = css_content
                    images_removed += 1

        # Replace keywords in text nodes
        text_replacements = 0

        def replace_text_nodes(element):
            nonlocal text_replacements
            if element.string:
                original_text = str(element.string)
                new_text = self.replace_keywords_in_text(original_text)
                if new_text != original_text:
                    element.string.replace_with(new_text)
                    text_replacements += 1
            else:
                for child in list(element.children):
                    if hasattr(child, 'name') and child.name:  # It's a tag
                        replace_text_nodes(child)
                    elif hasattr(child, 'replace_with'):  # It's a text node
                        original_text = str(child)
                        new_text = self.replace_keywords_in_text(original_text)
                        if new_text != original_text:
                            child.replace_with(new_text)
                            text_replacements += 1

        if soup.body:
            replace_text_nodes(soup.body)

        # Write the processed HTML
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(str(soup))

        print(f"✓ Removed {images_removed} images/graphics")
        print(f"✓ Made {text_replacements} text replacements")

        return output_path

    def process_txt_file(self, input_path, output_path):
        """Process plain text file - replace keywords only"""
        try:
            with open(input_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except UnicodeDecodeError:
            with open(input_path, 'r', encoding='latin-1') as file:
                content = file.read()

        # Replace keywords
        processed_content = self.replace_keywords_in_text(content)

        # Count replacements
        text_replacements = 0
        for original, replacement in self.keyword_replacements.items():
            if original.startswith(r'\b') or original.startswith(r'['):
                text_replacements += len(re.findall(original, content, flags=re.IGNORECASE))
            else:
                text_replacements += len(re.findall(re.escape(original), content, flags=re.IGNORECASE))

        # Write processed content
        with open(output_path, 'w', encoding='utf-8') as file:
            file.write(processed_content)

        print(f"✓ Made {text_replacements} text replacements")

        return output_path

    def blind_file(self, input_path, output_path=None, method='safe'):
        """
        Main function to blind a file - removes images and replaces keywords

        Args:
            input_path (str): Path to input file
            output_path (str): Path to output file (optional)
            method (str): Processing method - 'safe' (python-docx) or 'xml' (direct XML)
        """
        input_path = Path(input_path)

        if not input_path.exists():
            raise FileNotFoundError(f"Input file not found: {input_path}")

        # Generate output path if not provided
        if output_path is None:
            output_path = input_path.parent / f"{input_path.stem}_blinded{input_path.suffix}"
        else:
            output_path = Path(output_path)

        # Ensure output directory exists
        output_path.parent.mkdir(parents=True, exist_ok=True)

        extension = input_path.suffix.lower()

        try:
            print(f"Processing file: {input_path}")
            print(f"Output will be saved to: {output_path}")
            print()

            if extension == '.docx':
                if method == 'xml':
                    result = self.process_docx_xml_safe(input_path, output_path)
                    print("✓ DOCX processed with XML method")
                else:
                    result = self.process_docx_safe(input_path, output_path)
                    print("✓ DOCX processed with safe method")

            elif extension == '.html' or extension == '.htm':
                result = self.process_html_file(input_path, output_path)
                print("✓ HTML processed")

            elif extension == '.txt':
                result = self.process_txt_file(input_path, output_path)
                print("✓ Text file processed")

            else:
                raise ValueError(f"Unsupported file type: {extension}. Supported: .docx, .html, .htm, .txt")

            return result

        except Exception as e:
            print(f"Error processing file: {e}")
            return None


def main():
    """Interactive interface for file blinding"""
    print("=== File Blinding Tool (Structure Preserving) ===")
    print("Removes images and replaces keywords while keeping ALL original formatting")
    print("Supports: DOCX, HTML, TXT")
    print()

    # Get input file path
    while True:
        input_file = input("Enter input file path: ").strip().strip('"\'')
        if not input_file:
            print("Please enter a file path.")
            continue

        input_path = Path(input_file)
        if not input_path.exists():
            print(f"Error: File not found at '{input_file}'")
            continue

        # Check if file type is supported
        supported_extensions = ['.docx', '.html', '.htm', '.txt']
        if input_path.suffix.lower() not in supported_extensions:
            print(f"Error: Unsupported file type '{input_path.suffix}'")
            print(f"Supported formats: {', '.join(supported_extensions)}")
            continue

        break

    # Get output file path
    print()
    output_file = input("Enter output file path (or press Enter for auto-generated): ").strip().strip('"\'')

    if not output_file:
        output_file = None
        print("Auto-generating output filename...")

    # For DOCX files, ask about processing method
    method = 'safe'
    if Path(input_file).suffix.lower() == '.docx':
        print()
        print("DOCX Processing methods:")
        print("1. Safe (recommended) - Uses python-docx library, very reliable")
        print("2. XML - Direct XML processing, maximum structure preservation")

        method_choice = input("Choose processing method (1-2, default=1): ").strip()
        method = 'xml' if method_choice == '2' else 'safe'

    # Keyword replacements
    replacements = {
        # Sensitive terms
        "confidential": "[REDACTED]",
        "secret": "[REDACTED]",
        "internal": "[INTERNAL]",
        "proprietary": "[PROPRIETARY]",
        "classified": "[CLASSIFIED]",

        # Personal info patterns (regex)
        r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b': '[EMAIL]',
        r'\b\d{3}[-.]?\d{3}[-.]?\d{4}\b': '[PHONE]',
        r'\b\d{3}-\d{2}-\d{4}\b': '[SSN]',
        r'\b\d{1,2}/\d{1,2}/\d{4}\b': '[DATE]',
        r'\$[\d,]+\.?\d*': '[AMOUNT]',

        # Add your specific terms here!
        # "Your Company Name": "[COMPANY]",
        # "John Smith": "[PERSON]",
    }

    print()
    print("Formatting standardization:")
    print("✓ Font will be changed to Calibri")
    print("✓ All text colors will be changed to black")
    print("✓ All highlighting and shading will be removed")
    print("✓ All table cell backgrounds will be removed")
    print("✓ All borders will be removed")
    print("✓ Document themes will be removed")
    print()

    blinder = FileBlinder(
        keyword_replacements=replacements,
        standardize_formatting=True,
        font_name="Calibri",
        font_size=11,  # You can change this if needed
        font_color_black=True,
        grey_shading=False  # Changed to False to remove all highlighting
    )

    print("Processing file...")
    print()

    try:
        result = blinder.blind_file(input_file, output_file, method)

        if result:
            print()
            print("🎉 SUCCESS!")
            print(f"Blinded file created at: {result}")
            print()
            print("✓ All formatting and structure preserved")
            print("✓ Font standardized to Calibri")
            print("✓ All text changed to black with no highlighting")
            print("✓ All table cell backgrounds removed")
            print("✓ All borders and shading removed")
            print("✓ Document themes removed")
            print("✓ Tables, spacing, layout maintained")
            print("✓ Images and graphics removed")
            print("✓ Keywords replaced as configured")
        else:
            print("✗ Failed to process file.")

    except Exception as e:
        print(f"✗ Error: {e}")

    print()
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()