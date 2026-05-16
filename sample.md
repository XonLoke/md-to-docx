# Project Report

## Introduction

This document was generated from **Markdown** using the *MD → DOCX Converter*.
It demonstrates all supported formatting features including headings, lists,
tables, code blocks, and more.

---

## Text Formatting

You can use **bold**, *italic*, ***bold italic***, ~~strikethrough~~, and `inline code`.

Links are also supported: [Visit OpenAI](https://openai.com)

---

## Lists

### Unordered List

- Item one
- Item two
  - Nested item A
  - Nested item B
- Item three

### Ordered List

1. First step — install dependencies
2. Second step — run the application
3. Third step — open your browser

---

## Blockquote

> "Any sufficiently advanced technology is indistinguishable from magic."
> — Arthur C. Clarke

---

## Code Block

```python
def convert_markdown(text: str) -> bytes:
    """Convert markdown string to DOCX bytes."""
    doc = Document()
    # ... processing logic ...
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
```

```bash
# Run the web app
py app.py

# Or use the CLI
py cli.py README.md -o output.docx
```

---

## Table

| Feature          | Supported | Notes                        |
|------------------|-----------|------------------------------|
| Headings (H1–H6) | ✅        | With colour styling          |
| Bold / Italic    | ✅        | Inline formatting            |
| Code blocks      | ✅        | Monospace with grey bg       |
| Tables           | ✅        | Styled header + alt rows     |
| Blockquotes      | ✅        | Left border accent           |
| Ordered lists    | ✅        | Auto-numbered                |
| Unordered lists  | ✅        | Bullet points                |
| Horizontal rules | ✅        | Thin grey line               |
| Strikethrough    | ✅        | `~~text~~` syntax            |
| Inline code      | ✅        | Red monospace style          |

---

## Conclusion

The converter handles all common Markdown elements and exports them to a
clean, professional Word document ready for sharing or printing.
