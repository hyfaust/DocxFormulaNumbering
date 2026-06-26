"""
dfn.py - DocxFormulaNumbering

Word Document Equation Numbering Tool

Features:
- Automatic detection of single-line formulas in Word documents
- Chapter-based numbering format: (1-1), (1-2), (2-1)...
- Simple sequential numbering format: (1), (2), (3)...
- Bookmark creation for cross-referencing
- eqArr format support (professional equation layout)

Usage:
    python dfn.py <input docx> [output docx] [--simple] [--t1]

Examples:
    # Default mode: chapter-number format
    python dfn.py thesis.docx output.docx

    # Simple mode: sequential numbering
    python dfn.py thesis.docx output.docx --simple

    # eqArr format (requires t1.docx template)
    python dfn.py thesis.docx output.docx --t1

    # Overwrite original file
    python dfn.py document.docx

Author: FaustSherpad
Version: 1.0.0
"""

import win32com.client
import re
from pathlib import Path
import argparse
import zipfile
import shutil
import tempfile
import os


# ==================== Utility Functions ====================

def is_display_formula(para) -> bool:
    """Check if paragraph contains a single-line display formula"""
    if para.Range.OMaths.Count == 0:
        return False

    para_xml = para.Range.XML
    texts = []
    for t in re.findall(r'<w:t[^>]*>(.*?)</w:t>', para_xml, re.DOTALL):
        if t.strip():
            texts.append(t.strip())

    visible_texts = [t for t in texts if '<m:' not in t and '</m:' not in t]
    non_formula_text_len = sum(len(t) for t in visible_texts)
    has_label = any(':' in t for t in visible_texts)
    is_inline = non_formula_text_len > 20 or has_label

    return not is_inline


def get_chapter_number(para, word) -> int:
    """Get chapter number for current paragraph using Word COM"""
    try:
        original_range = word.Selection.Range.Duplicate
        para.Range.Select()
        word.Selection.Collapse(1)

        temp_field = word.ActiveDocument.Fields.Add(
            Range=word.Selection.Range,
            Type=-1,
            Text='STYLEREF "Heading 1" \\n',
            PreserveFormatting=False
        )

        temp_field.Update()
        chapter_text = temp_field.Result.Text.strip()
        temp_field.Delete()
        original_range.Select()

        if chapter_text and ('Error' in chapter_text):
            return 0

        if chapter_text:
            try:
                return int(chapter_text)
            except ValueError:
                return 0
        return 0

    except:
        return 0


def check_multilevel_list_configured(doc, word) -> bool:
    """Check if document has multi-level list configured"""
    try:
        original_range = word.Selection.Range.Duplicate
        doc.Range(0, 0).Select()

        temp_field = word.ActiveDocument.Fields.Add(
            Range=word.Selection.Range,
            Type=-1,
            Text='STYLEREF "Heading 1" \\n',
            PreserveFormatting=False
        )

        temp_field.Update()
        chapter_text = temp_field.Result.Text.strip()
        temp_field.Delete()
        original_range.Select()

        if chapter_text:
            if 'Error' in chapter_text:
                pass
            else:
                try:
                    if int(chapter_text) > 0:
                        return True
                except ValueError:
                    pass

        for para in doc.Paragraphs:
            try:
                style_name = para.Style.NameLocal
                if "Heading 1" in style_name:
                    if para.Range.ListFormat.ListValue > 0:
                        return True
            except:
                pass

        return False

    except Exception as e:
        print(f'Failed to check multi-level list configuration: {e}')
        return False


# ==================== eqArr Template Processing ====================

def extract_eqarr_template(template_path: str) -> dict:
    """
    Extract eqArr structure from t1.docx template

    Returns:
        dict: Dictionary containing template components
    """
    with zipfile.ZipFile(template_path, 'r') as zip_ref:
        template_xml = zip_ref.read('word/document.xml').decode('utf-8')

    # Extract oMathPara template
    omath_para_match = re.search(
        r'(<m:oMathPara>.*?</m:oMathPara>)',
        template_xml,
        re.DOTALL
    )

    if not omath_para_match:
        raise ValueError("oMathPara element not found in template")

    omath_template = omath_para_match.group(1)

    # Extract formula content (before #)
    formula_match = re.search(r'<m:t[^>]*>([^#]*)#', omath_template, re.DOTALL)
    if not formula_match:
        raise ValueError("Formula content not found in template")

    original_formula = formula_match.group(1).strip()

    # Extract number
    number_match = re.search(r'<m:t>(\d+)</m:t>', omath_template)
    if not number_match:
        raise ValueError("Formula number not found in template")

    original_number = number_match.group(1)

    # Extract eqArr environment structure
    eqarr_match = re.search(r'(<m:eqArr>.*?</m:eqArr>)', omath_template, re.DOTALL)
    if not eqarr_match:
        raise ValueError("eqArr element not found in template")

    eqarr_template = eqarr_match.group(1)

    return {
        'omath_para': omath_template,
        'eqarr': eqarr_template,
        'original_formula': original_formula,
        'original_number': original_number,
    }


def create_eqarr_formula(template: dict, formula_omml: str, number: str) -> str:
    """
    Create eqArr format formula using template with preserved OMML structure

    Args:
        template: Template dictionary from extract_eqarr_template
        formula_omml: Formula OMML structure (not plain text) - complete XML structure
        number: Formula number

    Returns:
        str: eqArr format oMathPara XML
    """
    # Start with eqarr template
    eqarr_xml = template['eqarr']

    # Template structure:
    # <m:eqArr>
    #   <m:e>
    #     <m:r><m:t>original_formula#</m:t></m:r>
    #     <m:d>  <- Number delimiter (INSIDE <m:e>, at the END)
    #       <m:e><m:r><m:t>1</m:t></m:r></m:e>
    #     </m:d>
    #   </m:e>
    # </m:eqArr>

    # formula_omml structure (extracted from numbered formula):
    # <m:r><m:t>E = m</m:t></m:r><m:sSup>...</m:sSup>
    # (multiple m:r and other elements possible)

    # Strategy:
    # 1. Find the <m:d> element (number delimiter) - keep it as-is but update number
    # 2. Find all content before <m:d> inside <m:e> - this is the formula area to replace
    # 3. Replace formula area with formula_omml + <m:r><m:t>#</m:t></m:r>

    # Find the <m:d> element
    d_match = re.search(r'<m:d>.*?</m:d>', eqarr_xml, re.DOTALL)
    if not d_match:
        # Fallback: just replace number
        eqarr_xml = re.sub(r'<m:t>\d+</m:t>', f'<m:t>{number}</m:t>', eqarr_xml, count=1)
        return f'<m:oMathPara><m:oMath>{eqarr_xml}</m:oMath></m:oMathPara>'

    d_start = d_match.start()
    d_end = d_match.end()
    d_element = d_match.group(0)

    # Content before <m:d> is the formula area
    before_d = eqarr_xml[:d_start]
    after_d = eqarr_xml[d_end:]

    # In before_d, find the <m:e> start tag
    # The structure is: <m:eqArr>...<m:e>...formula...</m:e>...<m:d>
    # But wait, <m:d> is INSIDE <m:e>, so before_d contains: <m:eqArr>...<m:e>...formula...

    # Find the last <m:e> in before_d
    e_start = before_d.rfind('<m:e>')
    if e_start == -1:
        # Fallback
        return f'<m:oMathPara><m:oMath>{eqarr_xml}</m:oMath></m:oMathPara>'

    # Find the end of <m:e> start tag
    e_tag_end = before_d.find('>', e_start) + 1

    # Content before <m:e> tag (including the tag itself)
    before_e_content = before_d[:e_tag_end]

    # Build new formula content: formula_omml + # marker
    new_formula = formula_omml.strip() + '<m:r><m:t>#</m:t></m:r>'

    # Update number in <m:d> element
    updated_d = re.sub(r'<m:t>\d+</m:t>', f'<m:t>{number}</m:t>', d_element, count=1)

    # Build new eqarr XML
    new_eqarr = before_e_content + new_formula + updated_d + after_d

    # Build complete oMathPara
    omath_para = f'<m:oMathPara><m:oMath>{new_eqarr}</m:oMath></m:oMathPara>'

    return omath_para


def extract_formula_omml(omath_content: str) -> str:
    """
    Extract complete OMML structure from oMathPara content before # delimiter

    Returns the formula elements inside m:oMath, preserving all structure
    like m:sSup, m:sSub, m:f, etc. Does not include outer m:oMathPara/m:oMath tags.

    Args:
        omath_content: oMathPara content containing formula with # delimiter

    Returns:
        str: Formula elements (m:r, m:sSup, etc.) without outer wrappers
    """
    # First, extract content inside <m:oMath>...</m:oMath>
    omath_match = re.search(r'<m:oMath>(.*?)</m:oMath>', omath_content, re.DOTALL)
    if not omath_match:
        # Fallback: use content as-is
        omath_inner = omath_content
    else:
        omath_inner = omath_match.group(1)

    # Find # delimiter position
    hash_pos = omath_inner.find('#')
    if hash_pos == -1:
        # No # found, return inner content as-is
        return omath_inner.strip()

    # Strategy: Find all XML elements before # and ensure they are properly closed
    # We need to count open and close tags to ensure balanced XML

    # Find the last </m:t> before #
    before_hash = omath_inner[:hash_pos]
    last_close_pos = before_hash.rfind('</m:t>')
    if last_close_pos == -1:
        # Fallback: just return everything before #
        return before_hash

    # Extract up to and including the last </m:t>
    partial_xml = omath_inner[:last_close_pos + 6]  # +6 for '</m:t>'

    # Now we need to close any open tags
    # Count open vs close tags for m:r, m:sup, m:sSup, m:sSub, m:e, etc.
    # Order matters: we need to close innermost tags first
    # The partial_xml ends with </m:t>, so we need to close tags in the order they were opened
    # e.g., if we have <m:sSup><m:e><m:r><m:t>2</m:t>, we need to close </m:r></m:e></m:sSup>
    tags_to_check = ['m:r', 'm:sup', 'm:sSup', 'm:sSub', 'm:e', 'm:d', 'm:f', 'm:num', 'm:den', 'm:rad']

    # Simple approach: count occurrences
    close_tags = []
    for tag in tags_to_check:  # Forward order: m:r first, then m:sup, then m:sSup
        open_count = len(re.findall(rf'<{tag}(?:\s[^>]*)?>', partial_xml))
        close_count = len(re.findall(rf'</{tag}>', partial_xml))
        # Add missing close tags
        for _ in range(open_count - close_count):
            close_tags.append(f'</{tag}>')

    # Append close tags
    return partial_xml.strip() + ''.join(close_tags)


def extract_formula_text(omath_content: str) -> str:
    """
    Extract plain text formula from oMath content

    Concatenate all <m:t> element contents until # delimiter
    """
    # Find all m:t elements
    texts = re.findall(r'<m:t[^>]*>([^<]*)</m:t>', omath_content)

    # Concatenate text until #
    result = []
    for t in texts:
        if '#' in t:
            # Found #, extract content before it
            idx = t.index('#')
            result.append(t[:idx])
            break
        else:
            result.append(t)

    return ''.join(result).strip()


def extract_formula_number(omath_content: str, simple_mode: bool = False) -> str:
    """
    Extract formula number from oMath content

    Args:
        omath_content: oMath content XML
        simple_mode: Simple mode (no chapter numbers)

    Returns:
        str: Number (plain number in simple mode, "chapter-number" format in default mode)
    """
    chapter = None
    num = None

    # Find bookmark name: eq{chap}_{num} or eq{num}
    bookmark_match = re.search(r'w:name="eq(\d+)_(\d+)"', omath_content)
    if bookmark_match:
        chapter = bookmark_match.group(1)
        num = bookmark_match.group(2)
    else:
        # Try simple mode bookmark: eq{num}
        bookmark_simple_match = re.search(r'w:name="eq(\d+)"', omath_content)
        if bookmark_simple_match:
            num = bookmark_simple_match.group(1)
            chapter = None

    # If bookmark not found, try to extract from SEQ field
    if num is None:
        # Find chapter number in SEQ field: SEQ \\* ARABIC \\s N
        seq_chapter_match = re.search(r'SEQ\s+\\*\s*ARABIC\s+\\s\s+(\d+)', omath_content)
        if seq_chapter_match:
            chapter = seq_chapter_match.group(1)

        # Find sequence number in field result
        # Find first number after SEQ field
        seq_match = re.search(r'SEQ\s+\\*\s*ARABIC', omath_content)
        if seq_match:
            seq_end = seq_match.end()
            after_seq = omath_content[seq_end:seq_end+200]
            num_match = re.search(r'<w:t[^>]*>(\d+)</w:t>', after_seq)
            if num_match:
                num = num_match.group(1)

    # If still not found, return None
    if num is None:
        return None

    # Return number based on mode
    if simple_mode or chapter is None:
        return num
    else:
        return f"{chapter}-{num}"


def convert_formula_in_xml(xml_content: str, template: dict, simple_mode: bool = False) -> str:
    """
    Convert formulas in XML from standard format to eqArr format

    Standard format: <m:r><m:t>formula#</m:t></m:r>...field...<m:r><m:t>)</m:t></m:r>
    eqArr format: <m:eqArr>...</m:eqArr>

    Args:
        xml_content: XML content
        template: Template dictionary
        simple_mode: Simple mode (no chapter numbers)
    """

    # Find all oMathPara elements
    omath_paras = re.findall(r'(<m:oMathPara>.*?</m:oMathPara>)', xml_content, re.DOTALL)

    print(f"Found {len(omath_paras)} oMathPara element(s)")

    for i, para in enumerate(omath_paras):
        # Check if already contains eqArr
        if '<m:eqArr>' in para:
            print(f"  oMathPara {i}: Already contains eqArr, skipping")
            continue

        # Extract formula content with preserved OMML structure
        formula_omml = extract_formula_omml(para)
        if not formula_omml:
            print(f"  oMathPara {i}: Formula content not found, skipping")
            continue

        # Also extract plain text for logging
        formula_text = extract_formula_text(para)

        # Extract number
        number = extract_formula_number(para, simple_mode)
        if not number:
            print(f"  oMathPara {i}: Formula number not found, skipping")
            continue

        print(f"  oMathPara {i}: Formula='{formula_text}', Number={number}")

        # Create eqArr format with preserved OMML structure
        eqarr_xml = create_eqarr_formula(template, formula_omml, number)

        # Replace oMathPara in original XML
        xml_content = xml_content.replace(para, eqarr_xml, 1)

    return xml_content


def postprocess_eqarr_format(docx_path: str, template_path: str, output_path: str = None, simple_mode: bool = False) -> bool:
    """
    Post-process: Convert formulas in docx file to eqArr format

    Args:
        docx_path: Path to numbered docx file
        template_path: Path to t1.docx template
        output_path: Output path (None to overwrite)
        simple_mode: Simple mode (no chapter numbers)

    Returns:
        bool: Success status
    """
    if output_path is None:
        output_path = docx_path

    try:
        # Extract template
        template = extract_eqarr_template(template_path)
        print(f"Template extracted: Formula='{template['original_formula']}', Number='{template['original_number']}'")

        # Read document XML
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            xml_content = zip_ref.read('word/document.xml').decode('utf-8')

        # Convert formulas
        print("Converting formula format...")
        new_xml = convert_formula_in_xml(xml_content, template, simple_mode)

        # Check if conversion occurred
        if new_xml == xml_content:
            print("Warning: No formulas detected for conversion")
        else:
            print("Formula format conversion completed")

        # Write to output file
        with zipfile.ZipFile(output_path, 'w') as zip_out:
            for item in zipfile.ZipFile(docx_path, 'r').infolist():
                if item.filename == 'word/document.xml':
                    zip_out.writestr(item, new_xml)
                else:
                    zip_out.writestr(item, zipfile.ZipFile(docx_path, 'r').read(item.filename))

        # Verify
        with zipfile.ZipFile(output_path, 'r') as zip_ref:
            new_xml_check = zip_ref.read('word/document.xml').decode('utf-8')
            has_eqarr = '<m:eqArr>' in new_xml_check
            has_delim = '<m:d>' in new_xml_check
            print(f"Verification: eqArr={has_eqarr}, m:d={has_delim}")

        return True

    except Exception as e:
        print(f"Post-processing failed: {e}")
        import traceback
        traceback.print_exc()
        return False


# ==================== Main Processing Function ====================

def add_formula_numbering(doc_path: str, output_path: str = None, simple_mode: bool = False, use_eqarr: bool = False):
    """
    Add equation numbers and bookmarks to Word document formulas

    Args:
        doc_path: Input document path
        output_path: Output document path
        simple_mode: Simple mode (sequential numbering without chapter numbers)
        use_eqarr: Use eqArr format (t1 format)
    """
    if output_path is None:
        output_path = doc_path

    word = None
    doc = None
    temp_file = None

    try:
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False

        doc = word.Documents.Open(str(doc_path))

        # If not simple mode, check if document has multi-level list configured
        if not simple_mode:
            print('Checking multi-level list configuration...')
            if not check_multilevel_list_configured(doc, word):
                doc.Close(SaveChanges=False)
                raise ValueError(
                    'Document not configured for multi-level list!\n\n'
                    'Solution:\n'
                    '1. Apply "Heading 1" style to chapter headings\n'
                    '2. Apply numbering via Home > Paragraph > Multilevel List\n'
                    '3. Re-run this script\n\n'
                    'Or use --simple parameter to skip chapter number check'
                )
            print('Multi-level list configuration verified')

        # Count chapters and formulas
        chapter = 0
        formula_paras = []
        formula_count = 0
        global_formula_count = 0

        for i, para in enumerate(doc.Paragraphs):
            para_text = para.Range.Text.strip()

            is_heading = False
            try:
                style_name = para.Style.NameLocal
                if "Heading 1" in style_name:
                    is_heading = True
            except:
                pass

            if is_heading:
                chapter += 1
                formula_count = 0
                continue

            if para.Range.OMaths.Count > 0:
                if is_display_formula(para):
                    para_xml = para.Range.XML
                    if 'SEQ' not in para_xml and 'STYLEREF' not in para_xml:
                        formula_count += 1
                        global_formula_count += 1

                        if simple_mode:
                            formula_paras.append((i, 0, global_formula_count, global_formula_count))
                        else:
                            formula_paras.append((i, chapter, formula_count, global_formula_count))

        if chapter == 0 and not simple_mode:
            doc.Close(SaveChanges=False)
            raise ValueError(
                'Document not configured for multi-level list!\n\n'
                'Solution:\n'
                '1. Apply "Heading 1" style to chapter headings\n'
                '2. Apply numbering via Home > Paragraph > Multilevel List\n'
                '3. Re-run this script\n\n'
                'Or use --simple parameter to skip chapter number check'
            )

        print(f'\nFound {len(formula_paras)} single-line formula(s)')

        # Process from back to front
        processed_count = 0
        for para_idx, chap, num, global_num in reversed(formula_paras):
            para = doc.Paragraphs(para_idx + 1)

            try:
                formula = para.Range.OMaths.Item(1)
                formula.Range.Select()
                word.Selection.Collapse(0)

                # 1. Add # delimiter
                word.Selection.TypeText('#')
                word.Selection.Collapse(0)

                # 2. Add left parenthesis
                word.Selection.TypeText('(')
                start_pos = word.Selection.End
                word.Selection.Collapse(0)

                if simple_mode:
                    field_code = 'SEQ Equation \\* ARABIC'
                    field = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text=field_code,
                        PreserveFormatting=False
                    )
                    field.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)
                else:
                    field1 = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text='STYLEREF "Heading 1" \\n',
                        PreserveFormatting=False
                    )
                    field1.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)

                    word.Selection.TypeText('-')
                    word.Selection.Collapse(0)

                    field_code = f'SEQ Equation \\* ARABIC \\s {chap}'
                    field2 = doc.Fields.Add(
                        Range=word.Selection.Range,
                        Type=-1,
                        Text=field_code,
                        PreserveFormatting=False
                    )
                    field2.Select()
                    word.Selection.MoveRight(Unit=2, Count=1)

                # 3. Add right parenthesis
                end_pos = word.Selection.Start
                word.Selection.TypeText(')')

                # 4. Create bookmark
                bookmark_range = doc.Range(start_pos, end_pos)
                if simple_mode:
                    bookmark_name = f'eq{global_num}'
                else:
                    bookmark_name = f'eq{chap}_{num}'
                doc.Bookmarks.Add(Name=bookmark_name, Range=bookmark_range)

                print(f'  Paragraph {para_idx}: Bookmark {bookmark_name}')
                processed_count += 1

            except Exception as e:
                print(f'Failed to process paragraph {para_idx}: {e}')

        print(f'\nProcessing completed, {processed_count} formula(s) processed')

        # Update all fields
        doc.Fields.Update()

        # If eqarr format enabled, save to temp file first
        if use_eqarr:
            template_path = Path(doc_path).parent / "t1.docx"
            if not template_path.exists():
                print(f"\nWarning: Template file not found: {template_path}")
                print("Using standard format")
                use_eqarr = False
            else:
                # Save to temp file
                temp_dir = tempfile.mkdtemp()
                temp_file = os.path.join(temp_dir, "temp_numbered.docx")
                doc.SaveAs2(temp_file)
                doc.Close()

                print(f"\nConverting to eqArr format...")
                if postprocess_eqarr_format(temp_file, str(template_path), str(output_path), simple_mode):
                    print(f"eqArr format conversion completed: {output_path}")
                    # Clean up temp files
                    try:
                        os.remove(temp_file)
                        os.rmdir(temp_dir)
                    except:
                        pass
                    return 0
                else:
                    print("eqArr conversion failed, using standard format")
                    # Fallback to standard format
                    doc = word.Documents.Open(temp_file)
                    try:
                        os.remove(temp_file)
                        os.rmdir(temp_dir)
                    except:
                        pass

        # Save in standard format
        if Path(output_path).exists():
            Path(output_path).unlink()

        doc.SaveAs2(str(output_path))
        print(f'Saved to: {output_path}')

        # Verify bookmarks
        print(f'\nVerifying bookmarks:')
        for bm in doc.Bookmarks:
            if bm.Name.startswith('eq'):
                print(f'  - {bm.Name}')

        doc.Close()
        return 0

    except ValueError as e:
        print(f'\nError: {e}')
        return 1
    except Exception as e:
        print(f'\nUnexpected error: {e}')
        import traceback
        traceback.print_exc()
        return 1
    finally:
        try:
            if doc:
                doc.Close(SaveChanges=False)
        except:
            pass

        try:
            if word:
                word.Quit()
        except:
            pass


def main():
    parser = argparse.ArgumentParser(
        description='DFN - DocxFormulaNumbering | Word Document Equation Numbering Tool',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Examples:
  # Default mode: chapter-number format (1-1), (1-2), (2-1)...
  python dfn.py input.docx output.docx

  # Simple mode: sequential numbering (1), (2), (3)...
  python dfn.py input.docx output.docx --simple

  # Use t1 format (eqArr environment, requires t1.docx template)
  python dfn.py input.docx output.docx --t1

  # Overwrite original file
  python dfn.py document.docx

Format Description:
  Default/t2 format: <m:t>E=mc^2#(1)</m:t> - Number as plain text
  t1 format: <m:eqArr><m:t>E=mc^2#</m:t><m:d>1</m:d></m:eqArr> - Number in delimiter element

  t1 format requires t1.docx template file

Prerequisites:
  In default mode, document must have multi-level list configured:
  1. Chapter headings use "Heading 1" style
  2. Apply numbering via Home > Paragraph > Multilevel List
  3. Heading style linked to list level
        '''
    )

    parser.add_argument('input', help='Input docx file path')
    parser.add_argument('output', nargs='?', default=None, help='Output docx file path (optional, overwrites input by default)')
    parser.add_argument('--simple', action='store_true', help='Use simple sequential numbering mode (no chapter numbers)')
    parser.add_argument('--t1', action='store_true', help='Use t1 format (eqArr environment, requires t1.docx template)')

    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve() if args.output else None

    if not input_path.exists():
        print(f'Error: File not found - {input_path}')
        return 1

    return add_formula_numbering(
        str(input_path),
        str(output_path) if output_path else None,
        simple_mode=args.simple,
        use_eqarr=args.t1
    )


if __name__ == '__main__':
    import sys
    sys.exit(main())
