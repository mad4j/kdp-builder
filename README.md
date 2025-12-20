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
- Table of contents support using `<<<toc>>>` marker
- Index entry support using `<<<index:term>>>` marker
- Cross-link support using standard Markdown link syntax `[text](#bookmark)` and `<<<bookmark:name>>>` marker
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
python -m kdpbuilder -m input.md -s styles.yaml -l layout.yaml -o output.docx
```

### Command-line Arguments

- `-m`, `--markdown`: Input Markdown file (required)
- `-s`, `--styles`: YAML file with style definitions (required)
- `-l`, `--layout`: YAML file with layout definition (required)
- `-o`, `--output`: Output DOCX file (required)

### Example

Try the included example which demonstrates all features (custom styling, table of contents, cross-links, index entries, and page breaks):
```bash
python -m kdpbuilder -m examples/example.md -s examples/styles.yaml -l examples/layout.yaml -o output.docx
```

The layout.yaml file includes comments showing how to use different units (inches, millimeters, or centimeters) for page dimensions and margins.

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

#### Table of Contents

To insert a table of contents, use the `<<<toc>>>` marker on its own line:

```markdown
{My Book Title}[title]

<<<toc>>>

<<<pagebreak>>>

# Chapter 1

Content for chapter 1...
```

The TOC marker creates a TOC field in the DOCX file. To generate the actual table of contents in Microsoft Word:
1. Right-click on the TOC field
2. Select "Update Field"
3. Choose whether to update page numbers only or the entire table

The table of contents will automatically include all headings in the document.

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

#### Cross-Links (Internal Hyperlinks)

Cross-links allow you to create internal hyperlinks within your document, enabling readers to jump between sections. This is particularly useful for KDP eBooks.

**Linking to Headings:**
All headings automatically get bookmarks based on their text. You can link to any heading using standard Markdown link syntax:

```markdown
# Chapter 1: Introduction

See [Chapter 2](#chapter_2_advanced_topics) for more details.

# Chapter 2: Advanced Topics

Content here. Go back to [Chapter 1](#chapter_1_introduction).
```

**Creating Custom Bookmarks:**
For non-heading sections, create custom bookmarks using the `<<<bookmark:name>>>` marker:

```markdown
# Chapter 1

<<<bookmark:important_section>>>

## Important Information

This section has a custom bookmark. 

You can reference it from anywhere: [See the important section](#important_section).
```

**Features:**
- Headings automatically generate bookmarks (e.g., "Chapter 1: Introduction" → `chapter_1_introduction`)
- Use `[link text](#bookmark_name)` to create a hyperlink
- Use `<<<bookmark:name>>>` to create a custom bookmark at any location
- Hyperlinks appear blue and underlined in the generated document
- Perfect for table of contents, cross-references, and navigation in KDP eBooks

**Example with Table of Contents:**
```markdown
{My Book}[title]

# Table of Contents

- [Chapter 1: Getting Started](#chapter_1_getting_started)
- [Chapter 2: Advanced Topics](#chapter_2_advanced_topics)
- [Appendix](#appendix)

<<<pagebreak>>>

# Chapter 1: Getting Started

Content here...

<<<pagebreak>>>

# Chapter 2: Advanced Topics

See [Chapter 1](#chapter_1_getting_started) for basics.

<<<pagebreak>>>

# Appendix

Additional information.
```

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
  unit: inches       # Unit for dimensions: "inches", "mm", or "cm" (default: inches)
  page_width: 6.0    # Page width
  page_height: 9.0   # Page height
  margin_top: 0.75   # Top margin
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

Example with millimeters:

```yaml
layout:
  unit: mm           # All dimensions will be in millimeters
  page_width: 152.4  # 6 inches = 152.4 mm
  page_height: 228.6 # 9 inches = 228.6 mm
  margin_top: 19.05  # 0.75 inches = 19.05 mm
  margin_bottom: 19.05
  margin_left: 19.05
  margin_right: 19.05
```

Example with centimeters:

```yaml
layout:
  unit: cm           # All dimensions will be in centimeters
  page_width: 15.24  # 6 inches = 15.24 cm
  page_height: 22.86 # 9 inches = 22.86 cm
  margin_top: 1.905  # 0.75 inches = 1.905 cm
  margin_bottom: 1.905
  margin_left: 1.905
  margin_right: 1.905
```

Available layout properties:
- `unit`: Unit for all dimension values - either `"inches"`, `"mm"`, or `"cm"` (default: `"inches"` for backward compatibility)
- `page_width`: Page width in the specified unit
- `page_height`: Page height in the specified unit
- `margin_top`: Top margin in the specified unit
- `margin_bottom`: Bottom margin in the specified unit
- `margin_left`: Left margin in the specified unit
- `margin_right`: Right margin in the specified unit
- `header_text`: Optional text to display in the document header. You can use `{page}` for the current page number and `{total}` for total pages (e.g., "Chapter 1 - Page {page}")
- `header_style`: Style name to apply to header text (must exist in styles.yaml)
- `footer_text`: Optional text to display in the document footer. You can use `{page}` for the current page number and `{total}` for total pages (e.g., "© 2025 Author - Page {page} of {total}")
- `footer_style`: Style name to apply to footer text (must exist in styles.yaml)

## License

This project is licensed under the GNU General Public License v3.0 - see the LICENSE file for details.
