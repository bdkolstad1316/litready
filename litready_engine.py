"""
LitReady Formatting Engine
==========================
Takes a messy .docx submission and produces a clean .docx with:
- Named paragraph styles (Story Title, Author Name, Spacer, Body Copy No Indent, Body Copy, Section Break)
- Named character styles (Italic, Bold, Bold Italic, Small Caps, Superscript)
- All inline formatting stripped (fonts, colors, spacing, kerning, ligatures)
- All styles based on Normal independently (no inheritance chains)
- Line spacing set to single (240 DXA)
- First-line indent of 0.25" (360 DXA) on Body Copy

Usage:
    python litready_engine.py input.docx output.docx --genre prose
"""

import argparse
import copy
from pathlib import Path
from lxml import etree
import zipfile
import shutil
import tempfile
import os

# ============================================================
# CONSTANTS
# ============================================================

NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

# Paragraph style definitions
PARA_STYLES = {
    'StoryTitle': {
        'name': 'Story Title',
        'basedOn': 'Normal',
        'next': 'AuthorName',
        'indent_first': 0,
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
    },
    'AuthorName': {
        'name': 'Author Name',
        'basedOn': 'Normal',
        'next': 'BodyCopyNoIndent',
        'indent_first': 0,
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
    },
    'Spacer': {
        'name': 'Spacer',
        'basedOn': 'Normal',
        'next': 'BodyCopyNoIndent',
        'indent_first': 0,
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
    },
    'BodyCopyNoIndent': {
        'name': 'Body Copy No Indent',
        'basedOn': 'Normal',  # NOT based on BodyCopy
        'next': 'BodyCopy',
        'indent_first': 0,
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
    },
    'BodyCopy': {
        'name': 'Body Copy',
        'basedOn': 'Normal',
        'next': 'BodyCopy',
        'indent_first': 360,  # 0.25" = 360 DXA
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
    },
    'SectionBreak': {
        'name': 'Section Break',
        'basedOn': 'Normal',
        'next': 'BodyCopyNoIndent',
        'indent_first': 0,
        'spacing_line': 240,
        'spacing_after': 0,
        'spacing_before': 0,
        'alignment': 'center',
    },
}

# Character style definitions
CHAR_STYLES = {
    'Italic': {'name': 'Italic', 'formatting': {'italic': True}},
    'Bold': {'name': 'Bold', 'formatting': {'bold': True}},
    'BoldItalic': {'name': 'Bold Italic', 'formatting': {'bold': True, 'italic': True}},
    'SmallCaps': {'name': 'Small Caps', 'formatting': {'smallCaps': True}},
    'Superscript': {'name': 'Superscript', 'formatting': {'vertAlign': 'superscript'}},
}

# Section break patterns
SECTION_BREAK_PATTERNS = {'***', '* * *', '***', '#', '##', '###', '—', '***', '⁂', '✱', '✱ ✱ ✱'}


def qn(tag):
    """Convert a namespace-prefixed tag to Clark notation. e.g. 'w:p' -> '{http://...}p'"""
    prefix, local = tag.split(':')
    return f'{{{NSMAP[prefix]}}}{local}'


# ============================================================
# STYLE INJECTION
# ============================================================

def build_paragraph_style_xml(style_id, style_def):
    """Build a <w:style> element for a paragraph style."""
    style_el = etree.SubElement(
        etree.Element('dummy'), qn('w:style'),
        {qn('w:type'): 'paragraph', qn('w:styleId'): style_id, qn('w:customStyle'): '1'}
    )
    
    etree.SubElement(style_el, qn('w:name'), {qn('w:val'): style_def['name']})
    etree.SubElement(style_el, qn('w:basedOn'), {qn('w:val'): style_def['basedOn']})
    etree.SubElement(style_el, qn('w:next'), {qn('w:val'): style_def['next']})
    etree.SubElement(style_el, qn('w:qFormat'))
    
    pPr = etree.SubElement(style_el, qn('w:pPr'))
    
    # Spacing
    spacing_attrs = {
        qn('w:line'): str(style_def['spacing_line']),
        qn('w:lineRule'): 'auto',
        qn('w:after'): str(style_def.get('spacing_after', 0)),
        qn('w:before'): str(style_def.get('spacing_before', 0)),
    }
    etree.SubElement(pPr, qn('w:spacing'), spacing_attrs)
    
    # Indentation
    if style_def['indent_first'] > 0:
        etree.SubElement(pPr, qn('w:ind'), {qn('w:firstLine'): str(style_def['indent_first'])})
    else:
        etree.SubElement(pPr, qn('w:ind'), {qn('w:firstLine'): '0'})
    
    # Alignment
    if 'alignment' in style_def:
        etree.SubElement(pPr, qn('w:jc'), {qn('w:val'): style_def['alignment']})
    
    return style_el


def build_character_style_xml(style_id, style_def):
    """Build a <w:style> element for a character style."""
    style_el = etree.SubElement(
        etree.Element('dummy'), qn('w:style'),
        {qn('w:type'): 'character', qn('w:styleId'): style_id, qn('w:customStyle'): '1'}
    )
    
    etree.SubElement(style_el, qn('w:name'), {qn('w:val'): style_def['name']})
    etree.SubElement(style_el, qn('w:basedOn'), {qn('w:val'): 'DefaultParagraphFont'})
    etree.SubElement(style_el, qn('w:qFormat'))
    
    rPr = etree.SubElement(style_el, qn('w:rPr'))
    fmt = style_def['formatting']
    
    if fmt.get('italic'):
        etree.SubElement(rPr, qn('w:i'))
        etree.SubElement(rPr, qn('w:iCs'))
    if fmt.get('bold'):
        etree.SubElement(rPr, qn('w:b'))
        etree.SubElement(rPr, qn('w:bCs'))
    if fmt.get('smallCaps'):
        etree.SubElement(rPr, qn('w:smallCaps'))
    if fmt.get('vertAlign'):
        etree.SubElement(rPr, qn('w:vertAlign'), {qn('w:val'): fmt['vertAlign']})
    
    return style_el


def inject_styles(styles_tree):
    """Inject all LitReady paragraph and character styles into styles.xml."""
    root = styles_tree.getroot()
    
    # ---- Clean docDefaults: strip fonts, kerning, ligatures ----
    doc_defaults = root.find(qn('w:docDefaults'))
    if doc_defaults is not None:
        # Clean rPrDefault — remove font refs, kerning, ligatures, font sizes
        rPr_default = doc_defaults.find(f'{qn("w:rPrDefault")}/{qn("w:rPr")}')
        if rPr_default is not None:
            for child in list(rPr_default):
                tag_local = child.tag.split('}')[-1]
                # Keep only language settings
                if tag_local in ('rFonts', 'kern', 'sz', 'szCs', 'ligatures'):
                    rPr_default.remove(child)
                # Also strip anything from non-w namespaces (w14:ligatures, etc.)
                elif 'wordprocessingml/2006/main' not in child.tag:
                    rPr_default.remove(child)
        
        # Clean pPrDefault — normalize spacing
        pPr_default = doc_defaults.find(f'{qn("w:pPrDefault")}/{qn("w:pPr")}')
        if pPr_default is not None:
            spacing = pPr_default.find(qn('w:spacing'))
            if spacing is not None:
                pPr_default.remove(spacing)
            # Set clean defaults: no after-spacing, single line
            etree.SubElement(pPr_default, qn('w:spacing'), {
                qn('w:after'): '0',
                qn('w:line'): '240',
                qn('w:lineRule'): 'auto',
            })
    
    # ---- Also clean the theme reference from the font table ----
    # (The rFonts theme refs in docDefaults resolve to Aptos in modern Word)
    
    # Remove any existing LitReady styles (for re-processing)
    all_style_ids = set(PARA_STYLES.keys()) | set(CHAR_STYLES.keys())
    for existing in root.findall(qn('w:style')):
        sid = existing.get(qn('w:styleId'), '')
        if sid in all_style_ids:
            root.remove(existing)
    
    # Add paragraph styles
    for style_id, style_def in PARA_STYLES.items():
        style_el = build_paragraph_style_xml(style_id, style_def)
        root.append(style_el)
    
    # Add character styles
    for style_id, style_def in CHAR_STYLES.items():
        style_el = build_character_style_xml(style_id, style_def)
        root.append(style_el)
    
    return styles_tree


# ============================================================
# DOCUMENT ANALYSIS
# ============================================================

def get_paragraph_text(p):
    """Extract plain text from a paragraph element."""
    texts = []
    for r in p.findall(qn('w:r')):
        t = r.find(qn('w:t'))
        if t is not None and t.text:
            texts.append(t.text)
    return ''.join(texts)


def detect_run_formatting(r):
    """Detect what character formatting a run has. Returns a set of flags."""
    flags = set()
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        return flags
    
    # Check for italic (direct or via style)
    if rPr.find(qn('w:i')) is not None:
        flags.add('italic')
    if rPr.find(qn('w:iCs')) is not None:
        flags.add('italic')
    
    # Check for bold (direct or via style)
    if rPr.find(qn('w:b')) is not None:
        flags.add('bold')
    if rPr.find(qn('w:bCs')) is not None:
        flags.add('bold')
    
    # Check for named styles that imply formatting
    rStyle = rPr.find(qn('w:rStyle'))
    if rStyle is not None:
        val = rStyle.get(qn('w:val'), '')
        if val in ('Emphasis',):
            flags.add('italic')
        elif val in ('Strong',):
            flags.add('bold')
        elif val == 'IntenseEmphasis':
            flags.add('italic')
    
    # Small caps
    if rPr.find(qn('w:smallCaps')) is not None:
        flags.add('smallCaps')
    
    # Superscript
    vertAlign = rPr.find(qn('w:vertAlign'))
    if vertAlign is not None:
        val = vertAlign.get(qn('w:val'), '')
        if val == 'superscript':
            flags.add('superscript')
    
    return flags


def flags_to_char_style(flags):
    """Map a set of formatting flags to a character style ID."""
    if 'bold' in flags and 'italic' in flags:
        return 'BoldItalic'
    elif 'bold' in flags:
        return 'Bold'
    elif 'italic' in flags:
        return 'Italic'
    elif 'smallCaps' in flags:
        return 'SmallCaps'
    elif 'superscript' in flags:
        return 'Superscript'
    return None


def is_section_break(text):
    """Check if a paragraph's text matches a section break pattern."""
    stripped = text.strip()
    return stripped in SECTION_BREAK_PATTERNS


# ============================================================
# PARAGRAPH CLASSIFICATION (Prose)
# ============================================================

def classify_paragraphs_prose(paragraphs):
    """
    Classify each paragraph into a style for prose submissions.
    
    Returns a list of style IDs, one per paragraph.
    
    Logic:
    - P0: Story Title
    - P1: Author Name  
    - P2 (empty): Spacer
    - First body paragraph after title block: Body Copy No Indent
    - All subsequent body paragraphs: Body Copy
    - Section breaks (centered *** or #): Section Break
    - First paragraph after section break: Body Copy No Indent
    """
    classifications = []
    total = len(paragraphs)
    
    # Phase 1: Identify title block
    # Expect: Title, Author, optional Spacer(s)
    title_block_end = 0
    
    if total == 0:
        return classifications
    
    # P0 is always Story Title
    classifications.append('StoryTitle')
    
    if total == 1:
        return classifications
    
    # P1 is Author Name
    classifications.append('AuthorName')
    
    # Check for spacer paragraphs after author
    i = 2
    while i < total:
        text = get_paragraph_text(paragraphs[i])
        if not text.strip():
            classifications.append('Spacer')
            i += 1
        else:
            break
    
    title_block_end = i
    
    # Phase 2: Classify body paragraphs
    after_break = True  # First body paragraph gets No Indent
    
    while i < total:
        text = get_paragraph_text(paragraphs[i])
        stripped = text.strip()
        
        if not stripped:
            # Empty paragraph — could be trailing whitespace or accidental
            classifications.append('Spacer')
            i += 1
            continue
        
        if is_section_break(stripped):
            classifications.append('SectionBreak')
            after_break = True
            i += 1
            continue
        
        if after_break:
            classifications.append('BodyCopyNoIndent')
            after_break = False
        else:
            classifications.append('BodyCopy')
        
        i += 1
    
    return classifications


# ============================================================
# CORE ENGINE: CLEAN & REMAP
# ============================================================

def strip_run_formatting(r):
    """
    Remove ALL inline formatting from a run, preserving only text content.
    Returns the formatting flags that were present (for character style mapping).
    """
    flags = detect_run_formatting(r)
    
    # Remove the entire rPr element
    rPr = r.find(qn('w:rPr'))
    if rPr is not None:
        r.remove(rPr)
    
    return flags


def apply_character_style(r, char_style_id):
    """Apply a named character style to a run via w:rStyle."""
    rPr = r.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(r, qn('w:rPr'))
        # Move rPr to be first child
        r.insert(0, rPr)
    
    # Clear any existing formatting in rPr
    for child in list(rPr):
        rPr.remove(child)
    
    # Add only the named style reference
    etree.SubElement(rPr, qn('w:rStyle'), {qn('w:val'): char_style_id})


def apply_paragraph_style(p, para_style_id):
    """Apply a named paragraph style to a paragraph, stripping all paragraph-level overrides."""
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(p, qn('w:pPr'))
        p.insert(0, pPr)
    
    # Remove everything from pPr
    for child in list(pPr):
        pPr.remove(child)
    
    # Add only the style reference
    etree.SubElement(pPr, qn('w:pStyle'), {qn('w:val'): para_style_id})


def clean_document(doc_tree, genre='prose'):
    """
    Main engine: clean a document tree.
    
    1. Classify paragraphs
    2. Strip all inline formatting from runs
    3. Apply named character styles based on detected formatting
    4. Apply named paragraph styles based on classification
    """
    root = doc_tree.getroot()
    body = root.find(qn('w:body'))
    
    if body is None:
        raise ValueError("No w:body element found in document.xml")
    
    paragraphs = body.findall(qn('w:p'))
    
    # Classify
    if genre in ('prose', 'nonfiction', 'hybrid'):
        classifications = classify_paragraphs_prose(paragraphs)
    else:
        # Poetry classification would go here
        # For now, fall back to prose
        classifications = classify_paragraphs_prose(paragraphs)
    
    # Ensure we have a classification for every paragraph
    while len(classifications) < len(paragraphs):
        classifications.append('BodyCopy')
    
    print(f"\n{'='*60}")
    print(f"  LitReady Engine — Processing {len(paragraphs)} paragraphs")
    print(f"  Genre: {genre}")
    print(f"{'='*60}\n")
    
    # Process each paragraph
    for i, (p, style_id) in enumerate(zip(paragraphs, classifications)):
        text = get_paragraph_text(p)
        display = text[:50] + ('...' if len(text) > 50 else '')
        
        # Process runs: strip formatting, detect character styles
        for r in p.findall(qn('w:r')):
            flags = strip_run_formatting(r)
            char_style = flags_to_char_style(flags)
            if char_style:
                apply_character_style(r, char_style)
        
        # Apply paragraph style
        apply_paragraph_style(p, style_id)
        
        # Log
        char_info = []
        for r in p.findall(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is not None:
                rs = rPr.find(qn('w:rStyle'))
                if rs is not None:
                    t = r.find(qn('w:t'))
                    txt = t.text[:20] if t is not None and t.text else ''
                    char_info.append(f"{rs.get(qn('w:val'))}:'{txt}'")
        
        char_str = f" | chars: {', '.join(char_info)}" if char_info else ""
        print(f"  P{i:3d} -> {style_id:22s} | \"{display}\"{char_str}")
    
    print(f"\n{'='*60}")
    print(f"  Done. {len(paragraphs)} paragraphs cleaned and mapped.")
    print(f"{'='*60}\n")
    
    return doc_tree


# ============================================================
# THEME FONT NEUTRALIZATION
# ============================================================

def neutralize_theme_fonts(theme_path):
    """
    Replace theme font definitions with Times New Roman (universally available).
    
    The theme's font scheme is where Aptos/Aptos Display live. When docDefaults
    references 'minorHAnsi' or 'majorHAnsi', Word resolves them through the theme.
    InDesign tries to do the same and fails if the font isn't installed.
    
    By replacing theme fonts with Times New Roman, any residual theme references
    resolve to a font every system has. Since our styles don't specify fonts at all,
    InDesign will use its own paragraph/character style fonts — which is the goal.
    """
    tree = etree.parse(str(theme_path))
    root = tree.getroot()
    
    # Find all latin typeface declarations and replace with Times New Roman
    for elem in root.iter():
        tag_local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
        if tag_local == 'latin':
            elem.set('typeface', 'Times New Roman')
            # Remove panose attribute if present (it's font-specific)
            if 'panose' in elem.attrib:
                del elem.attrib['panose']
    
    tree.write(str(theme_path), xml_declaration=True, encoding='UTF-8', standalone=True)


def clean_font_table(font_table_path):
    """
    Remove non-standard font entries from fontTable.xml.
    
    InDesign reads this table and complains about missing fonts even if
    no run in the document actually uses them. We strip everything except
    universally available fonts.
    """
    tree = etree.parse(str(font_table_path))
    root = tree.getroot()
    
    # Fonts that are safe to keep (universally available)
    safe_fonts = {
        'Times New Roman', 'Arial', 'Courier New', 'Symbol', 'Wingdings',
        'Calibri', 'Cambria', 'Georgia', 'Verdana', 'Tahoma',
        'Helvetica', 'Helvetica Neue',
    }
    
    for font in list(root):
        tag_local = font.tag.split('}')[-1] if '}' in font.tag else font.tag
        if tag_local == 'font':
            name = font.get(qn('w:name'), '')
            if name and name not in safe_fonts:
                root.remove(font)
    
    tree.write(str(font_table_path), xml_declaration=True, encoding='UTF-8', standalone=True)


# ============================================================
# FILE I/O
# ============================================================

def process_docx(input_path, output_path, genre='prose'):
    """
    Full pipeline: read .docx, clean it, write clean .docx.
    
    Works by unpacking the ZIP, modifying the XML, and repacking.
    """
    input_path = Path(input_path)
    output_path = Path(output_path)
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")
    
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        
        # Unpack
        with zipfile.ZipFile(input_path, 'r') as z:
            z.extractall(tmpdir)
        
        # Parse document.xml
        doc_path = tmpdir / 'word' / 'document.xml'
        doc_tree = etree.parse(str(doc_path))
        
        # Parse styles.xml
        styles_path = tmpdir / 'word' / 'styles.xml'
        styles_tree = etree.parse(str(styles_path))
        
        # Inject styles
        styles_tree = inject_styles(styles_tree)
        
        # Clean document
        doc_tree = clean_document(doc_tree, genre=genre)
        
        # Neutralize theme fonts (Aptos, Aptos Display, etc.)
        theme_path = tmpdir / 'word' / 'theme' / 'theme1.xml'
        if theme_path.exists():
            neutralize_theme_fonts(theme_path)
        
        # Clean fontTable — remove Aptos and other non-standard font entries
        font_table_path = tmpdir / 'word' / 'fontTable.xml'
        if font_table_path.exists():
            clean_font_table(font_table_path)
        
        # Write modified XML back
        doc_tree.write(str(doc_path), xml_declaration=True, encoding='UTF-8', standalone=True)
        styles_tree.write(str(styles_path), xml_declaration=True, encoding='UTF-8', standalone=True)
        
        # Repack as .docx
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for root_dir, dirs, files in os.walk(tmpdir):
                for f in files:
                    file_path = Path(root_dir) / f
                    arcname = file_path.relative_to(tmpdir)
                    zout.write(file_path, arcname)
    
    print(f"  Output: {output_path}")
    return output_path


# ============================================================
# CLI
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description='LitReady Formatting Engine — Clean .docx for InDesign',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python litready_engine.py submission.docx cleaned.docx --genre prose
  python litready_engine.py poem.docx cleaned_poem.docx --genre poetry
        """
    )
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', help='Output .docx file')
    parser.add_argument('--genre', choices=['prose', 'poetry', 'nonfiction', 'hybrid'],
                        default='prose', help='Genre (affects style mapping logic)')
    
    args = parser.parse_args()
    process_docx(args.input, args.output, genre=args.genre)


if __name__ == '__main__':
    main()
