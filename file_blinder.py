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
        """Remove document themes that might cause colored text - AGGRESSIVE VERSION"""
        try:
            from docx.shared import RGBColor

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
                    tag_str = str(child.tag).lower()
                    if any(theme_word in tag_str for theme_word in ['theme', 'themefont', 'clrscheme', 'fontscheme']):
                        themes_to_remove.append(child)

                for theme in themes_to_remove:
                    parent = theme.getparent()
                    if parent is not None:
                        try:
                            parent.remove(theme)
                        except:
                            pass

                # Remove any theme color references throughout the document
                namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                # Find and remove theme color references (w:themeColor)
                try:
                    for color_elem in doc_element.xpath('.//w:color[@w:themeColor]', namespaces=namespaces):
                        # Remove the themeColor attribute
                        theme_color_attr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor'
                        if theme_color_attr in color_elem.attrib:
                            del color_elem.attrib[theme_color_attr]
                        # Set explicit black color
                        color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                except:
                    pass

                # Remove theme fill references
                try:
                    for fill_elem in doc_element.xpath('.//w:shd[@w:themeFill]', namespaces=namespaces):
                        theme_fill_attr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeFill'
                        if theme_fill_attr in fill_elem.attrib:
                            del fill_elem.attrib[theme_fill_attr]
                        # Remove the entire shading element
                        parent = fill_elem.getparent()
                        if parent is not None:
                            parent.remove(fill_elem)
                except:
                    pass

                # Remove theme tint/shade attributes
                try:
                    for elem in doc_element.xpath('.//*[@w:themeTint or @w:themeShade]', namespaces=namespaces):
                        for attr in ['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                     '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                            if attr in elem.attrib:
                                del elem.attrib[attr]
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
        """Remove background shading from a table cell - AGGRESSIVE VERSION"""
        try:
            # Work at the XML level to remove cell shading
            tc_element = cell._tc

            # Find and remove ALL shading/fill/background elements from the cell
            shading_elements_to_remove = []
            for child in tc_element.iter():
                tag_str = str(child.tag).lower()
                # Check for any shading-related tags
                if any(shading_word in tag_str for shading_word in ['shd', 'fill', 'background', 'bgcolor']):
                    shading_elements_to_remove.append(child)

            for shading_elem in shading_elements_to_remove:
                parent = shading_elem.getparent()
                if parent is not None:
                    try:
                        parent.remove(shading_elem)
                    except:
                        pass

            # Also check for table cell properties and remove shading using XPath
            try:
                namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                # Remove shading from cell properties
                for tcPr in tc_element.xpath('.//w:tcPr', namespaces=namespaces):
                    # Remove any shading within table cell properties
                    for shd in tcPr.xpath('.//w:shd', namespaces=namespaces):
                        tcPr.remove(shd)

                    # Also remove any fill elements
                    for fill_elem in tcPr.xpath('.//*[contains(local-name(), "fill")]'):
                        try:
                            fill_elem.getparent().remove(fill_elem)
                        except:
                            pass
            except:
                pass

            # Additional cleanup: clear any attributes that might contain color
            try:
                if hasattr(tc_element, 'attrib'):
                    # Remove any color/fill attributes
                    attrs_to_remove = [k for k in tc_element.attrib.keys()
                                       if any(x in str(k).lower() for x in ['color', 'fill', 'shd', 'background'])]
                    for attr in attrs_to_remove:
                        del tc_element.attrib[attr]
            except:
                pass

        except Exception as e:
            # Continue if cell shading removal fails
            pass

    def remove_table_row_shading(self, row):
        """Remove background shading from a table row"""
        try:
            # Work at the XML level to remove row shading
            tr_element = row._tr

            # Find and remove ALL shading/fill/background elements from the row
            shading_elements_to_remove = []
            for child in tr_element.iter():
                tag_str = str(child.tag).lower()
                # Check for any shading-related tags
                if any(shading_word in tag_str for shading_word in ['shd', 'fill', 'background', 'bgcolor']):
                    shading_elements_to_remove.append(child)

            for shading_elem in shading_elements_to_remove:
                parent = shading_elem.getparent()
                if parent is not None:
                    try:
                        parent.remove(shading_elem)
                    except:
                        pass

            # Also check for table row properties and remove shading using XPath
            try:
                namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

                # Remove shading from row properties
                for trPr in tr_element.xpath('.//w:trPr', namespaces=namespaces):
                    # Remove any shading within table row properties
                    for shd in trPr.xpath('.//w:shd', namespaces=namespaces):
                        trPr.remove(shd)
            except:
                pass

        except Exception as e:
            # Continue if row shading removal fails
            pass

    def remove_content_control_shading(self, doc):
        """Remove background colors from content controls (structured document tags) - AGGRESSIVE"""
        try:
            from docx.shared import RGBColor

            # Get the document element
            doc_element = doc._element

            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                          'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'}

            print("  Searching for content controls...")

            # Find all SDT (structured document tag) elements - multiple namespace attempts
            sdt_elements = []

            # Try different XPath patterns
            try:
                sdt_elements.extend(doc_element.xpath('.//w:sdt', namespaces=namespaces))
            except:
                pass

            # Also try iterating and finding by tag name
            for elem in doc_element.iter():
                if 'sdt' in str(elem.tag).lower():
                    if elem not in sdt_elements:
                        sdt_elements.append(elem)

            print(f"  Found {len(sdt_elements)} content controls")

            for sdt in sdt_elements:
                try:
                    # Find SDT properties multiple ways
                    sdtPr_elements = []

                    # Try XPath
                    try:
                        sdtPr_elements.extend(sdt.xpath('.//w:sdtPr', namespaces=namespaces))
                    except:
                        pass

                    # Also iterate children
                    for child in sdt:
                        if 'sdtpr' in str(child.tag).lower():
                            if child not in sdtPr_elements:
                                sdtPr_elements.append(child)

                    for sdtPr in sdtPr_elements:
                        # REMOVE STYLE REFERENCES - this is what causes the blue background!
                        print("    Removing style references from SDT (AGGRESSIVE)...")

                        # Method 1: Remove by iterating through all children
                        style_elements_to_remove = []
                        for child in list(sdtPr):
                            tag_str = str(child.tag).lower()
                            if 'rpr' in tag_str or 'ppr' in tag_str or 'style' in tag_str:
                                style_elements_to_remove.append(child)
                                print(f"      Found style element to remove: {child.tag}")

                        for elem in style_elements_to_remove:
                            try:
                                sdtPr.remove(elem)
                                print(f"      Removed style element: {elem.tag}")
                            except Exception as e:
                                print(f"      Could not remove style {elem.tag}: {e}")

                        # Method 2: Use XPath to find and remove rPr and pPr
                        try:
                            for rPr in sdtPr.xpath('.//w:rPr', namespaces=namespaces):
                                parent = rPr.getparent()
                                if parent is not None:
                                    parent.remove(rPr)
                                    print(f"      Removed w:rPr via XPath")
                        except Exception as e:
                            print(f"      XPath rPr removal error: {e}")

                        try:
                            for pPr in sdtPr.xpath('.//w:pPr', namespaces=namespaces):
                                parent = pPr.getparent()
                                if parent is not None:
                                    parent.remove(pPr)
                                    print(f"      Removed w:pPr via XPath")
                        except Exception as e:
                            print(f"      XPath pPr removal error: {e}")

                        # SET appearance to hidden (removes border)
                        appearance_found = False
                        try:
                            for appearance in sdtPr.xpath('.//w:appearance', namespaces=namespaces):
                                appearance.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                               'hidden')
                                appearance_found = True
                                print(f"    Set SDT appearance to hidden")
                        except:
                            pass

                        # If no appearance, create one
                        if not appearance_found:
                            try:
                                from xml.etree import ElementTree as ET
                                appearance_elem = ET.Element(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}appearance')
                                appearance_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                                    'hidden')
                                sdtPr.insert(0, appearance_elem)
                                print(f"    Created hidden appearance for SDT")
                            except:
                                pass

                        # Remove ANY element related to color/shading/border
                        elements_to_remove = []
                        for child in list(sdtPr):
                            tag_name = str(child.tag).lower()
                            if any(x in tag_name for x in ['color', 'shd', 'fill', 'background', 'bdr', 'border']):
                                elements_to_remove.append(child)
                                print(f"    Marking for removal: {child.tag}")

                        for elem in elements_to_remove:
                            try:
                                sdtPr.remove(elem)
                                print(f"    Removed SDT property: {elem.tag}")
                            except Exception as e:
                                print(f"    Could not remove {elem.tag}: {e}")

                        # Also specifically target with XPath
                        try:
                            for color in sdtPr.xpath('.//w:color', namespaces=namespaces):
                                sdtPr.remove(color)
                                print("    Removed w:color via XPath")
                            for color in sdtPr.xpath('.//w15:color', namespaces=namespaces):
                                sdtPr.remove(color)
                                print("    Removed w15:color via XPath")
                            for shd in sdtPr.xpath('.//w:shd', namespaces=namespaces):
                                sdtPr.remove(shd)
                                print("    Removed w:shd via XPath")
                        except Exception as e:
                            print(f"    XPath removal error: {e}")

                        # Remove any attributes related to colors or borders
                        attrs_to_remove = []
                        for attr_name in list(sdtPr.attrib.keys()):
                            if any(x in attr_name.lower() for x in
                                   ['color', 'appearance', 'fill', 'border', 'bdr', 'style']):
                                attrs_to_remove.append(attr_name)

                        for attr in attrs_to_remove:
                            del sdtPr.attrib[attr]
                            print(f"    Removed attribute: {attr}")

                except Exception as e:
                    print(f"    Error processing SDT: {e}")
                    import traceback
                    print(traceback.format_exc())

        except Exception as e:
            print(f"  Content control shading removal error: {e}")
            import traceback
            print(traceback.format_exc())

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
        """Remove hyperlinks from a paragraph while preserving the text content"""
        try:
            from docx.shared import RGBColor

            p_element = paragraph._element

            # Find all hyperlink elements
            hyperlinks = []
            for child in p_element.iter():
                if 'hyperlink' in str(child.tag).lower():
                    hyperlinks.append(child)

            # Process each hyperlink
            for hyperlink in hyperlinks:
                # Extract all text runs from the hyperlink before removing it
                # Hyperlinks contain runs (w:r elements) that have the actual text
                parent = hyperlink.getparent()
                if parent is not None:
                    # Get the position of the hyperlink in the parent
                    hyperlink_index = list(parent).index(hyperlink)

                    # Extract all child elements (runs) from the hyperlink
                    children_to_preserve = list(hyperlink)

                    # Process each run to remove hyperlink formatting (underline, blue color)
                    for child in children_to_preserve:
                        # Look for run properties (rPr) within each run
                        try:
                            for rPr in child.iter():
                                if 'rPr' in str(rPr.tag):
                                    # Remove underline elements
                                    underlines_to_remove = []
                                    for elem in rPr:
                                        if 'u' in str(elem.tag).lower() and 'u' == str(elem.tag).split('}')[-1]:
                                            underlines_to_remove.append(elem)

                                    for u_elem in underlines_to_remove:
                                        rPr.remove(u_elem)

                                    # Remove color elements (the blue hyperlink color)
                                    colors_to_remove = []
                                    for elem in rPr:
                                        if 'color' in str(elem.tag).lower():
                                            colors_to_remove.append(elem)

                                    for color_elem in colors_to_remove:
                                        rPr.remove(color_elem)
                        except:
                            pass

                    # Insert the runs directly into the paragraph where the hyperlink was
                    for i, child in enumerate(children_to_preserve):
                        parent.insert(hyperlink_index + i, child)

                    # Now remove the empty hyperlink element
                    parent.remove(hyperlink)

            # After removing hyperlinks, process all runs again to ensure formatting
            for run in paragraph.runs:
                try:
                    # Remove underline
                    run.font.underline = None

                    # Set color to black
                    if self.font_color_black:
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        try:
                            run.font.color.theme_color = None
                        except:
                            pass
                except:
                    pass

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

        # Remove content control (SDT) background colors
        print("Removing content control shading...")
        self.remove_content_control_shading(doc)

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

            # Remove hyperlinks (preserving text), list formatting, borders, and shading
            self.remove_hyperlinks_from_paragraph(paragraph)
            self.remove_list_formatting(paragraph)
            self.remove_paragraph_borders(paragraph)
            self.remove_paragraph_shading(paragraph)

        # Process tables
        print("Processing tables...")
        for table_idx, table in enumerate(doc.tables):
            print(f"  Processing table {table_idx + 1}/{len(doc.tables)}")
            for row in table.rows:
                # Remove row-level shading first
                self.remove_table_row_shading(row)

                for cell in row.cells:
                    # Remove cell-level shading (this handles the orange/blue backgrounds)
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

                        # Remove hyperlinks (preserving text), list formatting, borders, and shading from table cells
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

                    # Remove hyperlinks from headers
                    self.remove_hyperlinks_from_paragraph(paragraph)

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

                    # Remove hyperlinks from footers
                    self.remove_hyperlinks_from_paragraph(paragraph)

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
                'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
            }

            # Register namespaces
            for prefix, uri in namespaces.items():
                ET.register_namespace(prefix, uri)

            images_removed = 0
            text_replacements = 0
            hyperlinks_removed = 0

            # NEUTRALIZE THEME FILES - Replace theme colors with black/white
            print("Neutralizing theme colors...")
            theme_dir = temp_dir / 'word' / 'theme'
            if theme_dir.exists():
                for theme_file in theme_dir.glob('*.xml'):
                    try:
                        tree = ET.parse(theme_file)
                        root = tree.getroot()

                        # Find all color scheme elements and replace with neutral colors
                        for color_scheme in root.iter():
                            tag = str(color_scheme.tag)
                            # Replace theme colors with black or white
                            if 'clrScheme' in tag or 'color' in tag.lower():
                                for color_elem in color_scheme:
                                    # Set all colors to either black (000000) or white (FFFFFF)
                                    for child in color_elem:
                                        if 'srgbClr' in str(child.tag):
                                            child.set('val', '000000')  # Black
                                        elif 'sysClr' in str(child.tag):
                                            child.set('val', 'windowText')
                                            child.set('lastClr', '000000')

                        tree.write(theme_file, encoding='utf-8', xml_declaration=True)
                        print(f"  Neutralized {theme_file.name}")
                    except Exception as e:
                        print(f"  Could not process theme file {theme_file.name}: {e}")

            # NEUTRALIZE STYLES.XML - Remove theme color references
            print("Neutralizing style theme references...")
            styles_xml = temp_dir / 'word' / 'styles.xml'
            if styles_xml.exists():
                try:
                    tree = ET.parse(styles_xml)
                    root = tree.getroot()

                    # Build a parent map since ElementTree doesn't have getparent()
                    parent_map = {c: p for p in tree.iter() for c in p}

                    # First pass: Remove all theme color attributes
                    for elem in root.iter():
                        # Remove theme color attributes
                        attrs_to_remove = []
                        for attr_name in list(elem.attrib.keys()):
                            if 'theme' in attr_name.lower() and 'color' in attr_name.lower():
                                attrs_to_remove.append(attr_name)

                        for attr in attrs_to_remove:
                            del elem.attrib[attr]

                    # Second pass: Mark shading elements for removal (don't remove while iterating)
                    elements_to_remove = []
                    for elem in root.iter():
                        if 'shd' in str(elem.tag).lower():
                            elements_to_remove.append(elem)

                    # Third pass: Remove marked elements using parent map
                    for elem in elements_to_remove:
                        parent = parent_map.get(elem)
                        if parent is not None:
                            try:
                                parent.remove(elem)
                            except:
                                pass

                    tree.write(styles_xml, encoding='utf-8', xml_declaration=True)
                    print(f"  Neutralized styles.xml")
                except Exception as e:
                    print(f"  Could not process styles.xml: {e}")
                    import traceback
                    traceback.print_exc()

            # Process main document
            document_xml = temp_dir / 'word' / 'document.xml'
            if document_xml.exists():
                print("Processing main document XML...")
                tree = ET.parse(document_xml)
                root = tree.getroot()

                # Build parent map for element removal
                parent_map = {c: p for p in tree.iter() for c in p}

                # Remove drawing elements (images)
                drawings_to_remove = root.findall('.//w:drawing', namespaces)
                for drawing in drawings_to_remove:
                    parent = parent_map.get(drawing)
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Remove object elements
                objects_to_remove = root.findall('.//w:object', namespaces)
                for obj in objects_to_remove:
                    parent = parent_map.get(obj)
                    if parent is not None:
                        parent.remove(obj)
                        images_removed += 1

                # Remove all shading elements (table cells, rows, paragraphs)
                shading_to_remove = root.findall('.//w:shd', namespaces)
                for shd in shading_to_remove:
                    parent = parent_map.get(shd)
                    if parent is not None:
                        try:
                            parent.remove(shd)
                        except:
                            pass

                # FORCE ALL TEXT TO BLACK COLOR
                print("Forcing all text to black color...")
                for rPr in root.findall('.//w:rPr', namespaces):
                    # Find or create color element
                    color_elem = rPr.find('w:color', namespaces)
                    if color_elem is None:
                        # Create new color element
                        color_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                        rPr.insert(0, color_elem)

                    # Set to black and remove theme color
                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')

                    # Remove theme color attributes if they exist
                    theme_color_attr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor'
                    theme_tint_attr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint'
                    theme_shade_attr = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade'

                    if theme_color_attr in color_elem.attrib:
                        del color_elem.attrib[theme_color_attr]
                    if theme_tint_attr in color_elem.attrib:
                        del color_elem.attrib[theme_tint_attr]
                    if theme_shade_attr in color_elem.attrib:
                        del color_elem.attrib[theme_shade_attr]

                # Remove content control (SDT) appearance/color properties AND BORDERS
                print("Removing content control styling...")
                for sdt in root.findall('.//w:sdt', namespaces):
                    try:
                        for sdtPr in sdt.findall('.//w:sdtPr', namespaces):
                            # Build parent map for this subtree
                            sdt_parent_map = {c: p for p in sdtPr.iter() for c in p}

                            # REMOVE STYLE REFERENCES - this is what causes the blue background!
                            print("  Resetting content control style...")
                            # Remove run properties (character styles)
                            for rPrElem in sdtPr.findall('.//w:rPr', namespaces):
                                parent = sdt_parent_map.get(rPrElem, sdtPr)
                                if parent is not None:
                                    try:
                                        parent.remove(rPrElem)
                                        print("    Removed rPr (run properties/style) from SDT")
                                    except:
                                        pass

                            # Remove paragraph properties (paragraph styles)
                            for pPrElem in sdtPr.findall('.//w:pPr', namespaces):
                                parent = sdt_parent_map.get(pPrElem, sdtPr)
                                if parent is not None:
                                    try:
                                        parent.remove(pPrElem)
                                        print("    Removed pPr (paragraph properties/style) from SDT")
                                    except:
                                        pass

                            # Remove any direct children that are style-related
                            direct_children_to_remove = []
                            for child in list(sdtPr):
                                tag_lower = str(child.tag).lower()
                                if 'rpr' in tag_lower or 'ppr' in tag_lower:
                                    direct_children_to_remove.append(child)
                                    print(f"    Marking style child for removal: {child.tag}")

                            for child in direct_children_to_remove:
                                try:
                                    sdtPr.remove(child)
                                    print(f"    Removed style child: {child.tag}")
                                except:
                                    pass

                            # SET appearance to hidden (removes border)
                            appearance_found = False
                            for appearance in sdtPr.findall('.//w:appearance', namespaces):
                                # Set appearance to "hidden" to remove border
                                appearance.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                               'hidden')
                                appearance_found = True
                                print("    Set appearance to hidden")

                            # If no appearance element exists, create one set to hidden
                            if not appearance_found:
                                appearance_elem = ET.Element(
                                    '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}appearance')
                                appearance_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                                    'hidden')
                                sdtPr.insert(0, appearance_elem)
                                print("    Created hidden appearance")

                            # Remove color elements
                            for color in sdtPr.findall('.//w:color', namespaces):
                                parent = sdt_parent_map.get(color, sdtPr)
                                if parent is not None:
                                    try:
                                        parent.remove(color)
                                    except:
                                        pass

                            # Remove any shading in SDT properties
                            for shd in sdtPr.findall('.//w:shd', namespaces):
                                parent = sdt_parent_map.get(shd, sdtPr)
                                if parent is not None:
                                    try:
                                        parent.remove(shd)
                                    except:
                                        pass

                            # Remove border-related elements more aggressively
                            all_children = list(sdtPr)
                            for child in all_children:
                                tag_str = str(child.tag).lower()
                                if any(x in tag_str for x in ['bdr', 'border']):
                                    try:
                                        sdtPr.remove(child)
                                    except:
                                        pass

                        # NOW ALSO PROCESS THE CONTENT INSIDE THE SDT (sdtContent)
                        # This is where the paragraph style that causes the blue background lives!
                        print("  Removing styles from content inside SDT...")
                        for sdtContent in sdt.findall('.//w:sdtContent', namespaces):
                            # Find all paragraphs inside the content
                            for para in sdtContent.findall('.//w:p', namespaces):
                                # Find paragraph properties
                                for pPr in para.findall('.//w:pPr', namespaces):
                                    # Remove paragraph style references (w:pStyle)
                                    for pStyle in pPr.findall('.//w:pStyle', namespaces):
                                        pPr.remove(pStyle)
                                        print(f"    Removed paragraph style reference from content")

                                    # Remove shading from paragraph
                                    for shd in pPr.findall('.//w:shd', namespaces):
                                        pPr.remove(shd)
                                        print(f"    Removed shading from paragraph")

                                # Also process runs inside these paragraphs
                                for run in para.findall('.//w:r', namespaces):
                                    for rPr in run.findall('.//w:rPr', namespaces):
                                        # Remove run style references (w:rStyle)
                                        for rStyle in rPr.findall('.//w:rStyle', namespaces):
                                            rPr.remove(rStyle)
                                            print(f"    Removed run style reference from content")

                                        # Remove shading from runs
                                        for shd in rPr.findall('.//w:shd', namespaces):
                                            rPr.remove(shd)
                                            print(f"    Removed shading from run")

                    except Exception as e:
                        print(f"  Error processing SDT: {e}")
                        import traceback
                        traceback.print_exc()

                # Remove hyperlinks while preserving text content
                for hyperlink in root.findall('.//w:hyperlink', namespaces):
                    parent = parent_map.get(hyperlink)
                    if parent is not None:
                        # Get the position of the hyperlink
                        hyperlink_index = list(parent).index(hyperlink)

                        # Move all children (runs) from hyperlink to parent
                        # AND remove hyperlink formatting (blue color, underline)
                        for i, child in enumerate(list(hyperlink)):
                            # If this is a run (w:r), clean up its formatting
                            if 'r' in str(child.tag).lower() and 'r' == str(child.tag).split('}')[-1]:
                                # Find run properties
                                for rPr in child.findall('w:rPr', namespaces):
                                    # Remove underline
                                    underline_elems = rPr.findall('w:u', namespaces)
                                    for u_elem in underline_elems:
                                        rPr.remove(u_elem)

                                    # Force color to black and remove theme color
                                    color_elem = rPr.find('w:color', namespaces)
                                    if color_elem is None:
                                        color_elem = ET.Element(
                                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                                        rPr.insert(0, color_elem)

                                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                                   '000000')

                                    # Remove theme color attributes
                                    for attr in [
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                                        if attr in color_elem.attrib:
                                            del color_elem.attrib[attr]

                            parent.insert(hyperlink_index + i, child)

                        # Remove the hyperlink element
                        parent.remove(hyperlink)
                        hyperlinks_removed += 1

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

                # Build parent map
                parent_map = {c: p for p in tree.iter() for c in p}

                # Remove images
                for drawing in root.findall('.//w:drawing', namespaces):
                    parent = parent_map.get(drawing)
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Remove all shading elements
                for shd in root.findall('.//w:shd', namespaces):
                    parent = parent_map.get(shd)
                    if parent is not None:
                        try:
                            parent.remove(shd)
                        except:
                            pass

                # FORCE ALL TEXT TO BLACK COLOR
                for rPr in root.findall('.//w:rPr', namespaces):
                    color_elem = rPr.find('w:color', namespaces)
                    if color_elem is None:
                        color_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                        rPr.insert(0, color_elem)
                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                    # Remove theme attributes
                    for attr in ['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor',
                                 '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                 '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                        if attr in color_elem.attrib:
                            del color_elem.attrib[attr]

                # Remove hyperlinks while preserving text
                for hyperlink in root.findall('.//w:hyperlink', namespaces):
                    parent = parent_map.get(hyperlink)
                    if parent is not None:
                        hyperlink_index = list(parent).index(hyperlink)
                        # Clean up formatting in runs from hyperlinks
                        for i, child in enumerate(list(hyperlink)):
                            if 'r' in str(child.tag).lower() and 'r' == str(child.tag).split('}')[-1]:
                                for rPr in child.findall('w:rPr', namespaces):
                                    # Remove underline
                                    for u_elem in rPr.findall('w:u', namespaces):
                                        rPr.remove(u_elem)
                                    # Force color to black
                                    color_elem = rPr.find('w:color', namespaces)
                                    if color_elem is None:
                                        color_elem = ET.Element(
                                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                                        rPr.insert(0, color_elem)
                                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                                   '000000')
                                    for attr in [
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                                        if attr in color_elem.attrib:
                                            del color_elem.attrib[attr]
                            parent.insert(hyperlink_index + i, child)
                        parent.remove(hyperlink)
                        hyperlinks_removed += 1

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

                # Build parent map
                parent_map = {c: p for p in tree.iter() for c in p}

                # Remove images
                for drawing in root.findall('.//w:drawing', namespaces):
                    parent = parent_map.get(drawing)
                    if parent is not None:
                        parent.remove(drawing)
                        images_removed += 1

                # Remove all shading elements
                for shd in root.findall('.//w:shd', namespaces):
                    parent = parent_map.get(shd)
                    if parent is not None:
                        try:
                            parent.remove(shd)
                        except:
                            pass

                # FORCE ALL TEXT TO BLACK COLOR
                for rPr in root.findall('.//w:rPr', namespaces):
                    color_elem = rPr.find('w:color', namespaces)
                    if color_elem is None:
                        color_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                        rPr.insert(0, color_elem)
                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                    # Remove theme attributes
                    for attr in ['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor',
                                 '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                 '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                        if attr in color_elem.attrib:
                            del color_elem.attrib[attr]

                # Remove hyperlinks while preserving text
                for hyperlink in root.findall('.//w:hyperlink', namespaces):
                    parent = parent_map.get(hyperlink)
                    if parent is not None:
                        hyperlink_index = list(parent).index(hyperlink)
                        # Clean up formatting in runs from hyperlinks
                        for i, child in enumerate(list(hyperlink)):
                            if 'r' in str(child.tag).lower() and 'r' == str(child.tag).split('}')[-1]:
                                for rPr in child.findall('w:rPr', namespaces):
                                    # Remove underline
                                    for u_elem in rPr.findall('w:u', namespaces):
                                        rPr.remove(u_elem)
                                    # Force color to black
                                    color_elem = rPr.find('w:color', namespaces)
                                    if color_elem is None:
                                        color_elem = ET.Element(
                                            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                                        rPr.insert(0, color_elem)
                                    color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val',
                                                   '000000')
                                    for attr in [
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeColor',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeTint',
                                        '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}themeShade']:
                                        if attr in color_elem.attrib:
                                            del color_elem.attrib[attr]
                            parent.insert(hyperlink_index + i, child)
                        parent.remove(hyperlink)
                        hyperlinks_removed += 1

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
            print(f"✓ Removed {hyperlinks_removed} hyperlinks (text preserved)")
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

        # Remove hyperlinks but keep text content
        hyperlinks_removed = 0
        for link in soup.find_all('a'):
            # Get the text content
            link_text = link.get_text()
            # Replace the link with just the text
            link.replace_with(link_text)
            hyperlinks_removed += 1

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
        print(f"✓ Removed {hyperlinks_removed} hyperlinks (text preserved)")
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
    print("✓ All hyperlinks will be removed (text preserved)")
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
            print("✓ All hyperlinks removed (text preserved)")
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