#!/usr/bin/env python3
"""
BALANCED AGGRESSIVE Color and Border Removal Patch
Add this to your file_blinder.py or run as standalone script

This version strips:
- All content control styling (the blue sections)
- All paragraph borders (NOT table borders - tables preserved)
- All colors (forces everything to black)
- All shading and backgrounds
- PRESERVES: Tables, table structure, table content
"""

import zipfile
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET


def ultra_aggressive_docx_cleanup(input_path, output_path):
    """
    Balanced aggressive cleanup that removes colors and borders but PRESERVES tables
    """
    print(f"Processing: {input_path}")
    print("Using BALANCED AGGRESSIVE cleanup mode...")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir = Path(temp_dir)

        # Extract DOCX
        with zipfile.ZipFile(input_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)

        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml'
        }

        for prefix, uri in namespaces.items():
            ET.register_namespace(prefix, uri)

        # Remove content controls but preserve content (including tables)
        print("\n1. REMOVING content control structures (preserving ALL content including tables)...")
        document_xml = temp_dir / 'word' / 'document.xml'

        if document_xml.exists():
            tree = ET.parse(document_xml)
            root = tree.getroot()
            parent_map = {c: p for p in tree.iter() for c in p}

            # Find and unwrap ALL content controls (SDTs)
            sdts_removed = 0
            for sdt in root.findall('.//w:sdt', namespaces):
                parent = parent_map.get(sdt)
                if parent is not None:
                    # Get index of SDT
                    sdt_index = list(parent).index(sdt)

                    # Extract content from SDT (this includes tables!)
                    sdt_content = sdt.find('.//w:sdtContent', namespaces)
                    if sdt_content is not None:
                        # Move all children from sdtContent to parent
                        # This preserves tables, paragraphs, everything
                        for child in list(sdt_content):
                            parent.insert(sdt_index, child)
                            sdt_index += 1

                    # Remove the SDT wrapper entirely
                    parent.remove(sdt)
                    sdts_removed += 1

            print(f"   ✓ Removed {sdts_removed} content control wrappers (tables preserved)")

            # FORCE ALL TEXT TO BLACK - SUPER AGGRESSIVE
            print("\n2. FORCING all text to black...")
            colors_fixed = 0
            for element in root.iter():
                # Remove ANY color-related attributes (but not table structure attributes)
                attrs_to_remove = []
                for attr_name in list(element.attrib.keys()):
                    attr_lower = attr_name.lower()
                    # Remove color/fill attributes but preserve table grid attributes
                    if any(x in attr_lower for x in
                           ['color', 'fill', 'theme', 'highlight']) and 'grid' not in attr_lower:
                        attrs_to_remove.append(attr_name)

                for attr in attrs_to_remove:
                    del element.attrib[attr]
                    colors_fixed += 1

            # Force color in all rPr (run properties)
            for rPr in root.findall('.//w:rPr', namespaces):
                # Remove existing color elements
                for color_elem in list(rPr):
                    if 'color' in str(color_elem.tag).lower() or 'highlight' in str(color_elem.tag).lower():
                        rPr.remove(color_elem)

                # Add black color
                color_elem = ET.Element('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color')
                color_elem.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '000000')
                rPr.insert(0, color_elem)
                colors_fixed += 1

            print(f"   ✓ Fixed {colors_fixed} color-related elements")

            # REMOVE ALL SHADING (including table cell shading)
            print("\n3. REMOVING all shading and backgrounds...")
            shading_removed = 0
            for shd in root.findall('.//w:shd', namespaces):
                parent = parent_map.get(shd)
                if parent is not None:
                    parent.remove(shd)
                    shading_removed += 1
            print(f"   ✓ Removed {shading_removed} shading elements")

            # REMOVE PARAGRAPH BORDERS ONLY (keep table structure borders)
            print("\n4. REMOVING paragraph borders (keeping table structure)...")
            borders_removed = 0

            # Remove paragraph borders (these create the blue boxes)
            for pBdr in root.findall('.//w:pBdr', namespaces):
                parent = parent_map.get(pBdr)
                if parent is not None:
                    parent.remove(pBdr)
                    borders_removed += 1

            print(f"   ✓ Removed {borders_removed} paragraph border elements (table borders preserved)")

            # REMOVE ALL PARAGRAPH STYLES (reset to Normal)
            print("\n5. RESETTING all paragraph styles to Normal...")
            styles_reset = 0
            for pStyle in root.findall('.//w:pStyle', namespaces):
                # Reset to Normal style
                pStyle.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'Normal')
                styles_reset += 1
            print(f"   ✓ Reset {styles_reset} paragraph styles")

            # Save document.xml
            tree.write(document_xml, encoding='utf-8', xml_declaration=True)

        # NEUTRALIZE STYLES.XML
        print("\n6. NEUTRALIZING styles.xml...")
        styles_xml = temp_dir / 'word' / 'styles.xml'
        if styles_xml.exists():
            tree = ET.parse(styles_xml)
            root = tree.getroot()
            parent_map = {c: p for p in tree.iter() for c in p}

            # Remove ALL color and shading elements from styles
            for element in root.iter():
                if 'color' in str(element.tag).lower() or 'shd' in str(element.tag).lower():
                    parent = parent_map.get(element)
                    if parent is not None:
                        try:
                            parent.remove(element)
                        except:
                            pass

            tree.write(styles_xml, encoding='utf-8', xml_declaration=True)
            print("   ✓ Neutralized styles.xml")

        # DESTROY THEME FILES
        print("\n7. DESTROYING theme files...")
        theme_dir = temp_dir / 'word' / 'theme'
        if theme_dir.exists():
            import shutil
            shutil.rmtree(theme_dir)
            print("   ✓ Removed all theme files")

        # Rebuild DOCX
        print("\n8. Rebuilding DOCX...")
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zip_out:
            for file_path in temp_dir.rglob('*'):
                if file_path.is_file():
                    arc_path = file_path.relative_to(temp_dir)
                    zip_out.write(file_path, arc_path)

        print(f"\n✅ COMPLETE: {output_path}")
        print("\nAll content controls, colors, and styling removed!")
        print("✓ Tables preserved with structure intact")


if __name__ == "__main__":
    import sys

    print("=" * 70)
    print("BALANCED AGGRESSIVE DOCX COLOR REMOVAL")
    print("Removes: Colors, Content Control Styling, Paragraph Borders")
    print("Preserves: Tables, Table Structure, Document Layout")
    print("=" * 70)
    print()

    if len(sys.argv) > 1:
        input_file = sys.argv[1]
    else:
        input_file = input("Enter input DOCX file path: ").strip().strip('"\'')

    input_path = Path(input_file)
    if not input_path.exists():
        print(f"Error: File not found: {input_file}")
        sys.exit(1)

    output_path = input_path.parent / f"{input_path.stem}_ULTRA_CLEANED{input_path.suffix}"

    try:
        ultra_aggressive_docx_cleanup(input_path, output_path)
        print(f"\n🎉 Success! Check: {output_path}")
    except Exception as e:
        print(f"\n❌ Error: {e}")
        import traceback

        traceback.print_exc()