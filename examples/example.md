{KDP Builder - Complete Example}[title]

{Demonstrating All Features}[subtitle]

<<<pagebreak>>>

{Table of Contents}[heading1]

- [Chapter 1: Introduction](#chapter_1_introduction)
- [Chapter 2: Text Styling](#chapter_2_text_styling)
- [Chapter 3: Cross-Links and Navigation](#chapter_3_cross_links_and_navigation)
- [Chapter 4: Conclusion](#chapter_4_conclusion)

<<<toc>>>

<<<pagebreak>>>

# Chapter 1: Introduction

Welcome to {KDP Builder}[strong], a powerful tool for converting Markdown files to professionally formatted DOCX documents for {Amazon Kindle Direct Publishing}[emphasis].

This example demonstrates all the key features including custom styling, table of contents, cross-links, index entries, and page breaks.

<<<index:KDP Builder>>>
<<<index:Markdown>>>

## 1.1 Getting Started

To use this tool, you need three files:
- A Markdown file with your content
- A styles.yaml file defining text styles
- A layout.yaml file defining page layout

For more advanced features, see [Chapter 3](#chapter_3_cross_links_and_navigation).

<<<pagebreak>>>

# Chapter 2: Text Styling

KDP Builder supports {custom text styles}[highlight] using a simple syntax: `{text}[style_name]`

## 2.1 Available Styles

Here are examples of different text styles:

- {Normal text}[normal] - the default style
- {Emphasized text}[emphasis] - for subtle emphasis
- {Strong text}[strong] - for important points
- {Highlighted text}[highlight] - to draw attention
- {Quoted text}[quote] - for quotations

<<<index:text styles>>>
<<<index:formatting>>>

## 2.2 Headers

You can use standard Markdown headers from level 1 to 3:

### 2.2.1 Level 3 Header Example

Each header level has its own predefined style in the styles.yaml file.

<<<pagebreak>>>

# Chapter 3: Cross-Links and Navigation

This chapter demonstrates {internal hyperlinks}[emphasis] within the document.

<<<bookmark:navigation_section>>>

## 3.1 Linking to Sections

You can link to any heading using standard Markdown syntax:
- Back to [Chapter 1](#chapter_1_introduction)
- Jump to [Text Styling](#chapter_2_text_styling)
- See the [Conclusion](#chapter_4_conclusion)

<<<index:cross-links>>>
<<<index:hyperlinks>>>

## 3.2 Custom Bookmarks

You can also create custom bookmarks at any location using `<<<bookmark:name>>>` and reference them with `[text](#name)`.

For example, go back to the [navigation section](#navigation_section) above.

<<<pagebreak>>>

# Chapter 4: Conclusion

This example has demonstrated:

1. **Custom styling** with multiple text styles
2. **Table of contents** generation with the `<<<toc>>>` marker
3. **Page breaks** using the `<<<pagebreak>>>` marker
4. **Cross-links** for internal navigation
5. **Index entries** for creating document indexes
6. **Multiple heading levels** with automatic styling

{Thank you for using KDP Builder!}[strong]

For more information, visit the [project documentation](#table_of_contents) or go back to [Chapter 1](#chapter_1_introduction) to start over.

<<<index:documentation>>>
