{Cross-Link Example Document}[title]

{Demonstrating Internal Hyperlinks}[subtitle]

<<<pagebreak>>>

{Table of Contents}[heading1]

Jump directly to any section:

- [Chapter 1: Getting Started](#chapter_1_getting_started)
- [Chapter 2: Main Content](#chapter_2_main_content)  
- [Important Note](#important_note)
- [Chapter 3: Conclusion](#chapter_3_conclusion)

<<<pagebreak>>>

# Chapter 1: Getting Started

Welcome to this example document that demonstrates the cross-link feature.

Cross-links allow you to create {internal hyperlinks}[emphasis] within your document, making it easier for readers to navigate between sections.

For more details, see [Chapter 2: Main Content](#chapter_2_main_content).

## How It Works

- All headings automatically get bookmarks based on their text
- Use standard Markdown link syntax: `[text](#bookmark)`
- Create custom bookmarks with: `<<<bookmark:name>>>`

<<<pagebreak>>>

# Chapter 2: Main Content

This chapter contains the main content. You can reference earlier sections like [Chapter 1](#chapter_1_getting_started).

## Creating Links

Simply use the standard Markdown link syntax with a `#` prefix:

{[link text](#bookmark_name)}[normal]

The bookmark name is automatically generated from heading text, or you can create custom bookmarks.

<<<bookmark:important_note>>>

## Important Note

This section has a {custom bookmark}[strong] named "important_note". It can be referenced from anywhere in the document using `[text](#important_note)`.

Go back to the [Table of Contents](#table_of_contents).

<<<pagebreak>>>

# Chapter 3: Conclusion

That's all! Cross-links make your documents more interactive and easier to navigate, especially important for {KDP eBooks}[emphasis].

Quick links:
- [Back to start](#table_of_contents)
- [Review Chapter 1](#chapter_1_getting_started)
- [See the important note](#important_note)
