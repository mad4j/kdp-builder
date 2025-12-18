#!/usr/bin/env python3
"""
KDP Builder - Markdown to DOCX and PDF Converter

This script converts Markdown files with style attributes to DOCX or PDF format,
using YAML files for style definitions and document layout.
"""

import argparse
import re
import sys
from pathlib import Path
from typing import Dict, List, Any, Tuple

import yaml
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib.pagesizes import inch
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.lib.colors import HexColor
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak


class StyleDefinition:
    """Represents a style definition from YAML."""
    
    def __init__(self, style_data: Dict[str, Any]):
        self.font_name = style_data.get('font_name', 'Arial')
        self.font_size = style_data.get('font_size', 11)
        self.bold = style_data.get('bold', False)
        self.italic = style_data.get('italic', False)
        self.underline = style_data.get('underline', False)
        self.color = style_data.get('color')
        self.alignment = style_data.get('alignment', 'left')
        self.space_before = style_data.get('space_before', 0)
        self.space_after = style_data.get('space_after', 0)


class LayoutDefinition:
    """Represents document layout from YAML."""
    
    def __init__(self, layout_data: Dict[str, Any]):
        self.page_width = layout_data.get('page_width', 8.5)
        self.page_height = layout_data.get('page_height', 11)
        self.margin_top = layout_data.get('margin_top', 1.0)
        self.margin_bottom = layout_data.get('margin_bottom', 1.0)
        self.margin_left = layout_data.get('margin_left', 1.0)
        self.margin_right = layout_data.get('margin_right', 1.0)
        self.header_text = layout_data.get('header_text')
        self.header_style = layout_data.get('header_style', 'normal')
        self.footer_text = layout_data.get('footer_text')
        self.footer_style = layout_data.get('footer_style', 'normal')


class MarkdownParser:
    """Parses Markdown with custom style attributes."""
    
    # Pattern to match text with style attributes: {text}[style]
    STYLED_TEXT_PATTERN = re.compile(r'\{([^}]+)\}\[([^\]]+)\]')
    
    # Pattern to match Markdown headers
    HEADER_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$')
    
    # Pattern to match page break markers
    PAGEBREAK_PATTERN = re.compile(r'^<<<pagebreak>>>$', re.IGNORECASE)
    
    # Pattern to match index markers: <<<index:term>>>
    INDEX_PATTERN = re.compile(r'^<<<index:(.+)>>>$', re.IGNORECASE)
    
    # Pattern to match table of contents markers: <<<toc>>>
    TOC_PATTERN = re.compile(r'^<<<toc>>>$', re.IGNORECASE)
    
    @staticmethod
    def is_pagebreak(line: str) -> bool:
        """Check if a line is a page break marker."""
        return bool(MarkdownParser.PAGEBREAK_PATTERN.match(line.strip()))
    
    @staticmethod
    def is_index(line: str) -> Tuple[bool, str]:
        """
        Check if a line is an index marker.
        
        Returns (is_index, term) where term is the index entry text.
        """
        match = MarkdownParser.INDEX_PATTERN.match(line.strip())
        if match:
            return (True, match.group(1))
        return (False, '')
    
    @staticmethod
    def is_toc(line: str) -> bool:
        """Check if a line is a table of contents marker."""
        return bool(MarkdownParser.TOC_PATTERN.match(line.strip()))
    
    @staticmethod
    def parse_line(line: str) -> List[Tuple[str, str]]:
        """
        Parse a line of Markdown into text segments with their styles.
        
        Returns a list of tuples: [(text, style_name), ...]
        If no style is specified, style_name is 'normal'.
        """
        segments = []
        pos = 0
        
        # Check if line is a header
        header_match = MarkdownParser.HEADER_PATTERN.match(line)
        if header_match:
            level = len(header_match.group(1))
            text = header_match.group(2)
            style_name = f'heading{level}'
            return [(text, style_name)]
        
        # Process styled text and regular text
        for match in MarkdownParser.STYLED_TEXT_PATTERN.finditer(line):
            # Add any text before this match as normal text
            if match.start() > pos:
                normal_text = line[pos:match.start()]
                if normal_text.strip():
                    segments.append((normal_text, 'normal'))
            
            # Add the styled text
            text = match.group(1)
            style = match.group(2)
            segments.append((text, style))
            pos = match.end()
        
        # Add any remaining text as normal
        if pos < len(line):
            remaining = line[pos:]
            if remaining.strip():
                segments.append((remaining, 'normal'))
        
        # If no segments were found, treat the whole line as normal text
        if not segments and line.strip():
            segments.append((line, 'normal'))
        
        return segments


class DocxBuilder:
    """Builds a DOCX document from parsed Markdown with styles."""
    
    def __init__(self, styles: Dict[str, StyleDefinition], layout: LayoutDefinition):
        self.document = Document()
        self.styles = styles
        self.layout = layout
        # Create a default style if 'normal' is not defined
        if 'normal' not in self.styles:
            self.styles['normal'] = StyleDefinition({})
        self._apply_layout()
    
    @staticmethod
    def _add_field(run, field_name: str):
        """Add a field code to a run."""
        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.set(qn('xml:space'), 'preserve')
        instrText.text = field_name

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
    
    @staticmethod
    def _add_page_number_field(run):
        """Add page number field to a run."""
        DocxBuilder._add_field(run, "PAGE")
    
    @staticmethod
    def _add_page_count_field(run):
        """Add total page count field to a run."""
        DocxBuilder._add_field(run, "NUMPAGES")
    
    def _add_text_with_fields(self, paragraph, text: str, style_def: StyleDefinition):
        """Add text to a paragraph with support for {page} and {total} field codes."""
        if not text:
            return
        
        # Split the text into regular text and placeholder tokens
        # Pattern matches {page} or {total} placeholders
        parts = re.split(r'(\{page\}|\{total\})', text)
        
        # Add runs for each part
        for part in parts:
            if not part:  # Skip empty strings
                continue
            
            run = paragraph.add_run(part if part not in ['{page}', '{total}'] else '')
            self._apply_style_to_run(run, style_def)
            
            if part == '{page}':
                self._add_page_number_field(run)
            elif part == '{total}':
                self._add_page_count_field(run)
    
    def _apply_layout(self):
        """Apply layout settings to the document."""
        sections = self.document.sections
        for section in sections:
            section.page_width = Inches(self.layout.page_width)
            section.page_height = Inches(self.layout.page_height)
            section.top_margin = Inches(self.layout.margin_top)
            section.bottom_margin = Inches(self.layout.margin_bottom)
            section.left_margin = Inches(self.layout.margin_left)
            section.right_margin = Inches(self.layout.margin_right)
    
    def _apply_header_footer(self):
        """Apply header and footer to the document."""
        sections = self.document.sections
        for section in sections:
            # Add header if specified
            if self.layout.header_text:
                header = section.header
                header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
                # Clear any existing content
                header_para.clear()
                
                # Apply style to header
                style_def = self.styles.get(self.layout.header_style, self.styles['normal'])
                header_para.alignment = self._get_alignment(style_def.alignment)
                
                # Add text with support for {page} and {total} tags
                self._add_text_with_fields(header_para, self.layout.header_text, style_def)
            
            # Add footer if specified
            if self.layout.footer_text:
                footer = section.footer
                footer_para = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
                # Clear any existing content
                footer_para.clear()
                
                # Apply style to footer
                style_def = self.styles.get(self.layout.footer_style, self.styles['normal'])
                footer_para.alignment = self._get_alignment(style_def.alignment)
                
                # Add text with support for {page} and {total} tags
                self._add_text_with_fields(footer_para, self.layout.footer_text, style_def)
    
    def _get_alignment(self, alignment_str: str):
        """Convert alignment string to DOCX alignment constant."""
        alignment_map = {
            'left': WD_ALIGN_PARAGRAPH.LEFT,
            'center': WD_ALIGN_PARAGRAPH.CENTER,
            'right': WD_ALIGN_PARAGRAPH.RIGHT,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        return alignment_map.get(alignment_str.lower(), WD_ALIGN_PARAGRAPH.LEFT)
    
    def _parse_color(self, color_str: str) -> RGBColor:
        """Parse color string (hex format like '#FF0000') to RGBColor."""
        if color_str.startswith('#'):
            color_str = color_str[1:]
        
        # Validate color string format
        if len(color_str) != 6:
            raise ValueError(f"Invalid color format: {color_str}. Expected 6 hex characters.")
        
        try:
            r = int(color_str[0:2], 16)
            g = int(color_str[2:4], 16)
            b = int(color_str[4:6], 16)
        except ValueError:
            raise ValueError(f"Invalid hex color: {color_str}. Must contain only hex digits.")
        
        return RGBColor(r, g, b)
    
    def _apply_style_to_run(self, run, style_def: StyleDefinition):
        """Apply a style definition to a text run."""
        run.font.name = style_def.font_name
        run.font.size = Pt(style_def.font_size)
        run.font.bold = style_def.bold
        run.font.italic = style_def.italic
        run.font.underline = style_def.underline
        
        if style_def.color:
            run.font.color.rgb = self._parse_color(style_def.color)
    
    def add_paragraph(self, segments: List[Tuple[str, str]]):
        """Add a paragraph to the document with styled segments."""
        if not segments:
            self.document.add_paragraph()
            return
        
        # Use the style of the first segment to determine paragraph style
        first_style_name = segments[0][1]
        style_def = self.styles.get(first_style_name, self.styles['normal'])
        
        paragraph = self.document.add_paragraph()
        paragraph.alignment = self._get_alignment(style_def.alignment)
        
        if style_def.space_before > 0:
            paragraph.paragraph_format.space_before = Pt(style_def.space_before)
        if style_def.space_after > 0:
            paragraph.paragraph_format.space_after = Pt(style_def.space_after)
        
        # Add each segment with its style
        for text, style_name in segments:
            style_def = self.styles.get(style_name, self.styles['normal'])
            run = paragraph.add_run(text)
            self._apply_style_to_run(run, style_def)
    
    def add_page_break(self):
        """Add a page break to the document."""
        paragraph = self.document.add_paragraph()
        run = paragraph.add_run()
        run.add_break(WD_BREAK.PAGE)
    
    def add_index_entry(self, term: str):
        """
        Add an index entry marker to the document.
        
        This creates an XE (Index Entry) field that marks the term for inclusion
        in a document index. The field is hidden in the document but can be used
        to generate an index in Word.
        
        Args:
            term: The term to add to the index
        """
        paragraph = self.document.add_paragraph()
        run = paragraph.add_run()
        
        # Create XE field for index entry
        # Field code format: XE "term"
        # In Word field codes, quotes are escaped by doubling them
        escaped_term = term.replace('"', '""')
        field_code = f'XE "{escaped_term}"'
        
        # Add the field using the existing _add_field method
        DocxBuilder._add_field(run, field_code)
        
        # Make the paragraph hidden so the XE field doesn't show in the document
        # XE fields are typically hidden in Word documents
        pPr = paragraph._element.get_or_add_pPr()
        vanish = OxmlElement('w:vanish')
        pPr.append(vanish)
    
    def add_toc(self):
        """
        Add a table of contents field to the document.
        
        This creates a TOC field that can be updated in Word to generate
        a table of contents based on the document headings.
        """
        paragraph = self.document.add_paragraph()
        run = paragraph.add_run()
        
        # Create TOC field
        # Field code format: TOC \o "1-3" \h \z \u
        # \o "1-3" = include heading levels 1-3
        # \h = make TOC entries hyperlinks
        # \z = hide tab leader and page numbers in Web Layout view
        # \u = use applied paragraph outline level
        field_code = 'TOC \\o "1-3" \\h \\z \\u'
        
        # Add the field using the existing _add_field method
        DocxBuilder._add_field(run, field_code)
    
    def save(self, output_path: str):
        """Save the document to a file."""
        self.document.save(output_path)


class PdfBuilder:
    """Builds a PDF document from parsed Markdown with styles."""
    
    def __init__(self, styles: Dict[str, StyleDefinition], layout: LayoutDefinition):
        self.styles = styles
        self.layout = layout
        self.story = []
        self.pdf_styles = {}
        self.toc_entries = []
        self.page_count = 0
        
        # Create a default style if 'normal' is not defined
        if 'normal' not in self.styles:
            self.styles['normal'] = StyleDefinition({})
        
        # Convert styles to ReportLab format
        self._create_pdf_styles()
    
    def _create_pdf_styles(self):
        """Convert StyleDefinition objects to ReportLab ParagraphStyle objects."""
        for style_name, style_def in self.styles.items():
            # Map alignment
            alignment_map = {
                'left': TA_LEFT,
                'center': TA_CENTER,
                'right': TA_RIGHT,
                'justify': TA_JUSTIFY,
            }
            alignment_str = (style_def.alignment or 'left').lower()
            alignment = alignment_map.get(alignment_str, TA_LEFT)
            
            # Create ReportLab style
            pdf_style = ParagraphStyle(
                name=style_name,
                fontName=self._get_font_name(style_def),
                fontSize=style_def.font_size,
                alignment=alignment,
                spaceBefore=style_def.space_before,
                spaceAfter=style_def.space_after,
            )
            
            # Add color if specified
            if style_def.color:
                pdf_style.textColor = self._parse_color(style_def.color)
            
            self.pdf_styles[style_name] = pdf_style
    
    def _get_font_name(self, style_def: StyleDefinition) -> str:
        """Get ReportLab font name based on style definition."""
        # Map common font names to ReportLab built-in fonts
        font_map = {
            'arial': 'Helvetica',
            'times new roman': 'Times-Roman',
            'courier': 'Courier',
            'georgia': 'Times-Roman',
        }
        
        font_name = (style_def.font_name or 'Arial').lower()
        base_font = font_map.get(font_name, 'Helvetica')
        
        # Apply bold and italic modifiers
        if style_def.bold and style_def.italic:
            if base_font == 'Helvetica':
                return 'Helvetica-BoldOblique'
            elif base_font == 'Times-Roman':
                return 'Times-BoldItalic'
            elif base_font == 'Courier':
                return 'Courier-BoldOblique'
        elif style_def.bold:
            if base_font == 'Helvetica':
                return 'Helvetica-Bold'
            elif base_font == 'Times-Roman':
                return 'Times-Bold'
            elif base_font == 'Courier':
                return 'Courier-Bold'
        elif style_def.italic:
            if base_font == 'Helvetica':
                return 'Helvetica-Oblique'
            elif base_font == 'Times-Roman':
                return 'Times-Italic'
            elif base_font == 'Courier':
                return 'Courier-Oblique'
        
        return base_font
    
    def _parse_color(self, color_str: str) -> HexColor:
        """Parse color string (hex format like '#FF0000') to HexColor."""
        if not color_str.startswith('#'):
            color_str = '#' + color_str
        try:
            return HexColor(color_str)
        except (ValueError, AttributeError) as e:
            # Return black color as fallback for invalid color strings
            return HexColor('#000000')
    
    def add_paragraph(self, segments: List[Tuple[str, str]]):
        """Add a paragraph to the PDF document with styled segments."""
        if not segments:
            self.story.append(Spacer(1, 0.2 * inch))
            return
        
        # Use the style of the first segment to determine paragraph style
        first_style_name = segments[0][1]
        pdf_style = self.pdf_styles.get(first_style_name, self.pdf_styles['normal'])
        
        # Build the paragraph text with inline styling
        text_parts = []
        for text, style_name in segments:
            style_def = self.styles.get(style_name, self.styles['normal'])
            
            # Build inline style tags
            style_tags = []
            if style_def.bold:
                text = f"<b>{text}</b>"
            if style_def.italic:
                text = f"<i>{text}</i>"
            if style_def.underline:
                text = f"<u>{text}</u>"
            if style_def.color:
                color = style_def.color
                text = f'<font color="{color}">{text}</font>'
            
            text_parts.append(text)
        
        full_text = ''.join(text_parts)
        
        # Add as heading if it's a heading style
        if first_style_name.startswith('heading'):
            # Extract heading level safely
            try:
                level = int(first_style_name.replace('heading', ''))
            except (ValueError, IndexError):
                level = 1
            # Extract plain text for TOC
            plain_text = ''.join([seg[0] for seg in segments])
            self.toc_entries.append((level, plain_text))
        
        paragraph = Paragraph(full_text, pdf_style)
        self.story.append(paragraph)
    
    def add_page_break(self):
        """Add a page break to the PDF document."""
        self.story.append(PageBreak())
    
    def add_index_entry(self, term: str):
        """Add an index entry marker (placeholder for PDF - not fully supported in ReportLab)."""
        # ReportLab doesn't have native index support like DOCX
        # This is a placeholder that could be extended with custom implementation
        pass
    
    def add_toc(self):
        """Add a table of contents placeholder to the PDF document."""
        # Add a simple TOC using stored entries
        toc_style = self.pdf_styles.get('heading1', self.pdf_styles['normal'])
        toc_para = Paragraph("<b>Table of Contents</b>", toc_style)
        self.story.append(toc_para)
        self.story.append(Spacer(1, 0.3 * inch))
        
        # Add TOC entries
        for level, title in self.toc_entries:
            # Use proper indentation with left indent
            indent_amount = (level - 1) * 20  # 20 points per level
            toc_entry_style = ParagraphStyle(
                name=f'toc_level_{level}',
                parent=self.pdf_styles.get('normal'),
                leftIndent=indent_amount
            )
            toc_entry = Paragraph(title, toc_entry_style)
            self.story.append(toc_entry)
    
    def save(self, output_path: str):
        """Save the PDF document to a file."""
        # Calculate page size in points (1 inch = 72 points)
        page_width = self.layout.page_width * inch
        page_height = self.layout.page_height * inch
        pagesize = (page_width, page_height)
        
        # Calculate margins in points
        margin_top = self.layout.margin_top * inch
        margin_bottom = self.layout.margin_bottom * inch
        margin_left = self.layout.margin_left * inch
        margin_right = self.layout.margin_right * inch
        
        # Create PDF document
        doc = SimpleDocTemplate(
            output_path,
            pagesize=pagesize,
            topMargin=margin_top,
            bottomMargin=margin_bottom,
            leftMargin=margin_left,
            rightMargin=margin_right,
        )
        
        # Build the document
        doc.build(self.story)


def load_styles(yaml_path: str) -> Dict[str, StyleDefinition]:
    """Load style definitions from YAML file."""
    with open(yaml_path, 'r', encoding='utf-8') as f:
        styles_data = yaml.safe_load(f)
    
    styles = {}
    for style_name, style_data in styles_data.get('styles', {}).items():
        styles[style_name] = StyleDefinition(style_data)
    
    return styles


def load_layout(yaml_path: str) -> LayoutDefinition:
    """Load layout definition from YAML file."""
    with open(yaml_path, 'r', encoding='utf-8') as f:
        layout_data = yaml.safe_load(f)
    
    return LayoutDefinition(layout_data.get('layout', {}))


def convert_markdown_to_docx(markdown_path: str, styles_path: str, 
                             layout_path: str, output_path: str):
    """
    Convert a Markdown file to DOCX using style and layout definitions.
    
    Args:
        markdown_path: Path to input Markdown file
        styles_path: Path to YAML file with style definitions
        layout_path: Path to YAML file with layout definition
        output_path: Path to output DOCX file
    """
    # Load configuration
    styles = load_styles(styles_path)
    layout = load_layout(layout_path)
    
    # Create document builder
    builder = DocxBuilder(styles, layout)
    
    # Parse and add Markdown content
    with open(markdown_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.rstrip('\n')
            if MarkdownParser.is_pagebreak(line):  # Page break marker
                builder.add_page_break()
            elif MarkdownParser.is_toc(line):  # Table of contents marker
                builder.add_toc()
            else:
                # Check for index marker
                is_idx, term = MarkdownParser.is_index(line)
                if is_idx:  # Index marker
                    builder.add_index_entry(term)
                elif line.strip():  # Non-empty line
                    segments = MarkdownParser.parse_line(line)
                    builder.add_paragraph(segments)
                else:  # Empty line
                    builder.add_paragraph([])
    
    # Apply header and footer after content is added
    builder._apply_header_footer()
    
    # Save document
    builder.save(output_path)
    print(f"Document saved to: {output_path}")


def convert_markdown_to_pdf(markdown_path: str, styles_path: str, 
                            layout_path: str, output_path: str):
    """
    Convert a Markdown file to PDF using style and layout definitions.
    
    Args:
        markdown_path: Path to input Markdown file
        styles_path: Path to YAML file with style definitions
        layout_path: Path to YAML file with layout definition
        output_path: Path to output PDF file
    """
    # Load configuration
    styles = load_styles(styles_path)
    layout = load_layout(layout_path)
    
    # Create document builder
    builder = PdfBuilder(styles, layout)
    
    # Parse and add Markdown content
    with open(markdown_path, 'r', encoding='utf-8') as f:
        for line in f:
            line = line.rstrip('\n')
            if MarkdownParser.is_pagebreak(line):  # Page break marker
                builder.add_page_break()
            elif MarkdownParser.is_toc(line):  # Table of contents marker
                builder.add_toc()
            else:
                # Check for index marker
                is_idx, term = MarkdownParser.is_index(line)
                if is_idx:  # Index marker
                    builder.add_index_entry(term)
                elif line.strip():  # Non-empty line
                    segments = MarkdownParser.parse_line(line)
                    builder.add_paragraph(segments)
                else:  # Empty line
                    builder.add_paragraph([])
    
    # Save document
    builder.save(output_path)
    print(f"Document saved to: {output_path}")


def main():
    """Main entry point for the script."""
    parser = argparse.ArgumentParser(
        description='Convert Markdown with style attributes to DOCX or PDF format.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s -m input.md -s styles.yaml -l layout.yaml -o output.docx
  %(prog)s -m input.md -s styles.yaml -l layout.yaml -o output.pdf
  
The Markdown file supports custom style attributes:
  {text}[style_name] - Apply custom style to text
  # Heading 1 - Standard Markdown heading
  
Output format is determined by the file extension (.docx or .pdf)
        """
    )
    
    parser.add_argument('-m', '--markdown', required=True,
                       help='Input Markdown file')
    parser.add_argument('-s', '--styles', required=True,
                       help='YAML file with style definitions')
    parser.add_argument('-l', '--layout', required=True,
                       help='YAML file with layout definition')
    parser.add_argument('-o', '--output', required=True,
                       help='Output file (DOCX or PDF format based on extension)')
    
    args = parser.parse_args()
    
    # Validate input files exist
    for file_path, name in [(args.markdown, 'Markdown'),
                            (args.styles, 'Styles'),
                            (args.layout, 'Layout')]:
        if not Path(file_path).exists():
            print(f"Error: {name} file not found: {file_path}", file=sys.stderr)
            sys.exit(1)
    
    # Determine output format based on file extension
    output_path = Path(args.output)
    output_ext = output_path.suffix.lower()
    
    try:
        if output_ext == '.pdf':
            convert_markdown_to_pdf(args.markdown, args.styles, 
                                   args.layout, args.output)
        elif output_ext == '.docx':
            convert_markdown_to_docx(args.markdown, args.styles, 
                                    args.layout, args.output)
        else:
            print(f"Error: Unsupported output format '{output_ext}'. Use .docx or .pdf", 
                  file=sys.stderr)
            sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
