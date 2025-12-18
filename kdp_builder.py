#!/usr/bin/env python3
"""
KDP Builder - Markdown to DOCX Converter

This script converts Markdown files with style attributes to DOCX format,
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


class MarkdownParser:
    """Parses Markdown with custom style attributes."""
    
    # Pattern to match text with style attributes: {text}[style]
    STYLED_TEXT_PATTERN = re.compile(r'\{([^}]+)\}\[([^\]]+)\]')
    
    # Pattern to match Markdown headers
    HEADER_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$')
    
    # Pattern to match page break markers
    PAGEBREAK_PATTERN = re.compile(r'^<<<pagebreak>>>$', re.IGNORECASE)
    
    @staticmethod
    def is_pagebreak(line: str) -> bool:
        """Check if a line is a page break marker."""
        return bool(MarkdownParser.PAGEBREAK_PATTERN.match(line.strip()))
    
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
    
    def save(self, output_path: str):
        """Save the document to a file."""
        self.document.save(output_path)


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
        description='Convert Markdown with style attributes to DOCX format.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Example usage:
  %(prog)s -m input.md -s styles.yaml -l layout.yaml -o output.docx
  
The Markdown file supports custom style attributes:
  {text}[style_name] - Apply custom style to text
  # Heading 1 - Standard Markdown heading
        """
    )
    
    parser.add_argument('-m', '--markdown', required=True,
                       help='Input Markdown file')
    parser.add_argument('-s', '--styles', required=True,
                       help='YAML file with style definitions')
    parser.add_argument('-l', '--layout', required=True,
                       help='YAML file with layout definition')
    parser.add_argument('-o', '--output', required=True,
                       help='Output DOCX file')
    
    args = parser.parse_args()
    
    # Validate input files exist
    for file_path, name in [(args.markdown, 'Markdown'),
                            (args.styles, 'Styles'),
                            (args.layout, 'Layout')]:
        if not Path(file_path).exists():
            print(f"Error: {name} file not found: {file_path}", file=sys.stderr)
            sys.exit(1)
    
    try:
        convert_markdown_to_docx(args.markdown, args.styles, 
                                args.layout, args.output)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
