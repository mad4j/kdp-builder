# kdp-builder
Script for Amazon Kindle Direct Publishing (KDP) service

## Description

KDP Builder is a Python script that converts Markdown files with custom style attributes into DOCX format, suitable for Amazon KDP publishing. The script uses YAML files to define styles and document layout, providing a flexible way to create professionally formatted documents.

## Features

- Convert Markdown to DOCX format
- Custom style attributes in Markdown: `{text}[style]`
- YAML-based style definitions (fonts, sizes, colors, alignment, spacing)
- YAML-based layout definitions (page size, margins, headers, footers)
- Support for standard Markdown headers (`#`, `##`, etc.)
- Header and footer sections with custom styling and embedded page numbers using `{page}` and `{total}` tags
- Page break support using `<<<pagebreak>>>` marker
- Index entry support using `<<<index:term>>>` marker
- Command-line interface for easy usage

## Installation

1. Clone the repository:
```bash
git clone https://github.com/mad4j/kdp-builder.git
cd kdp-builder
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Basic usage:
```bash
python kdp_builder.py -m input.md -s styles.yaml -l layout.yaml -o output.docx
```

### Command-line Arguments

- `-m`, `--markdown`: Input Markdown file (required)
- `-s`, `--styles`: YAML file with style definitions (required)
- `-l`, `--layout`: YAML file with layout definition (required)
- `-o`, `--output`: Output DOCX file (required)

### Example

Try the included example:
```bash
python kdp_builder.py -m examples/example.md -s examples/styles.yaml -l examples/layout.yaml -o output.docx
```

## File Formats

### Markdown Format

The Markdown file supports custom style attributes using the syntax `{text}[style_name]`:

```markdown
{My Title}[title]

# Chapter 1

This is normal text with {emphasized text}[emphasis] and {highlighted text}[highlight].

<<<pagebreak>>>

# Chapter 2

Content on a new page after the page break.
```

#### Page Breaks

To insert a page break, use the `<<<pagebreak>>>` marker on its own line:

```markdown
Content on page 1

<<<pagebreak>>>

Content on page 2
```

#### Index Entries

To mark a term for the document index, use the `<<<index:term>>>` marker on its own line:

```markdown
# Chapter 1

This paragraph discusses important concepts.

<<<index:important concepts>>>

The term "important concepts" will be marked for inclusion in the index.
```

The index markers create hidden XE (Index Entry) fields in the DOCX file. To generate the actual index in Microsoft Word:
1. Position your cursor where you want the index to appear
2. Go to References → Insert Index
3. Configure your index settings and click OK

### Styles YAML Format

Define styles in YAML format:

```yaml
styles:
  normal:
    font_name: "Arial"
    font_size: 11
    bold: false
    italic: false
    underline: false
    alignment: "left"
    space_before: 0
    space_after: 6
  
  highlight:
    font_name: "Arial"
    font_size: 11
    bold: true
    color: "#FF0000"
    alignment: "left"
```

Available style properties:
- `font_name`: Font family name (e.g., "Arial", "Times New Roman")
- `font_size`: Font size in points
- `bold`: Boolean for bold text
- `italic`: Boolean for italic text
- `underline`: Boolean for underlined text
- `color`: Hex color code (e.g., "#FF0000" for red)
- `alignment`: Text alignment ("left", "center", "right", "justify")
- `space_before`: Space before paragraph in points
- `space_after`: Space after paragraph in points

### Layout YAML Format

Define document layout:

```yaml
layout:
  page_width: 6.0    # in inches
  page_height: 9.0   # in inches
  margin_top: 0.75   # in inches
  margin_bottom: 0.75
  margin_left: 0.75
  margin_right: 0.75
  
  # Optional header and footer
  # You can use {page} and {total} tags directly in the text
  header_text: "My Book Title"                          # text to display in header
  header_style: "subtitle"                              # style to apply (from styles.yaml)
  footer_text: "© 2025 Author Name - Page {page} of {total}"  # text to display in footer
  footer_style: "normal"                                # style to apply (from styles.yaml)
```

Available layout properties:
- `page_width`: Page width in inches
- `page_height`: Page height in inches
- `margin_top`: Top margin in inches
- `margin_bottom`: Bottom margin in inches
- `margin_left`: Left margin in inches
- `margin_right`: Right margin in inches
- `header_text`: Optional text to display in the document header. You can use `{page}` for the current page number and `{total}` for total pages (e.g., "Chapter 1 - Page {page}")
- `header_style`: Style name to apply to header text (must exist in styles.yaml)
- `footer_text`: Optional text to display in the document footer. You can use `{page}` for the current page number and `{total}` for total pages (e.g., "© 2025 Author - Page {page} of {total}")
- `footer_style`: Style name to apply to footer text (must exist in styles.yaml)

## License

This project is licensed under the GNU General Public License v3.0 - see the LICENSE file for details.
