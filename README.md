# Markdown → DOCX Converter

Convert any Markdown (`.md`) file into a nicely formatted Microsoft Word (`.docx`) document — via a web browser UI or the command line.

---

## Program Location

All application files are located at:

```
D:\sep_venv\md-to-docx\
```

### File Structure

```
md-to-docx/
│
├── app.py            ← Flask web server (run this to start the app)
├── converter.py      ← Core Markdown-to-DOCX conversion engine
├── cli.py            ← Command-line interface (no browser needed)
├── sample.md         ← Sample Markdown file for testing
│
├── templates/
│   └── index.html    ← Web UI (browser interface)
│
├── uploads/          ← (unused — files are not saved to disk)
└── outputs/          ← (unused — files are not saved to disk)
```

---

## Requirements

| Requirement | Version |
|-------------|---------|
| Python      | 3.14    |
| Flask       | 3.1.3   |
| python-docx | 1.2.0   |
| Markdown    | 3.10.2  |

All dependencies are already installed. If you ever need to reinstall them:

```bash
py -m pip install flask python-docx markdown
```

---

## Method 1 — Web Browser UI (Recommended)

### Step 1 — Open a terminal

Press `Win + R`, type `cmd`, press Enter.  
Or open **Windows Terminal** / **PowerShell**.

### Step 2 — Navigate to the app folder

```cmd
cd D:\sep_venv\md-to-docx
```

### Step 3 — Start the server

```cmd
py app.py
```

You should see:

```
✅  Markdown → DOCX converter running at http://127.0.0.1:5000
 * Running on http://127.0.0.1:5000
```

> Keep this terminal window open. Closing it will stop the server.

### Step 4 — Open your browser

Open any browser (Chrome, Edge, Firefox) and go to:

```
http://127.0.0.1:5000
```

### Step 5 — Convert a Markdown file

You have three ways to input your Markdown:

| Method | How |
|--------|-----|
| **Type / Paste** | Click the left text area and type or paste your Markdown directly |
| **Upload file** | Click the **"Upload .md file"** button in the toolbar and select a `.md` file |
| **Drag & Drop** | Drag a `.md` file from File Explorer and drop it onto the left text area |

The right panel shows a **live preview** of how your document will look.

### Step 6 — Download the DOCX

1. (Optional) Change the output filename in the **"Output filename"** field
2. Click the **"Download DOCX"** button
3. The `.docx` file will be saved to your browser's default **Downloads folder**  
   (usually `C:\Users\<YourName>\Downloads\`)

### Step 7 — Stop the server

When you are done, go back to the terminal and press:

```
Ctrl + C
```

---

## Method 2 — Command Line (No browser needed)

Use this if you want to convert files directly without opening a browser.

### Basic usage

```cmd
cd D:\sep_venv\md-to-docx
py cli.py <input_file.md>
```

The output `.docx` is saved in the **same folder** as the input file, with the same name.

**Example:**

```cmd
py cli.py sample.md
```

Produces: `sample.docx` in `D:\sep_venv\md-to-docx\`

### Custom output path

```cmd
py cli.py sample.md -o C:\Users\Lenovo\Documents\report.docx
```

### More examples

```cmd
# Convert a file in another folder
py cli.py C:\Users\Lenovo\Documents\notes.md

# Convert and save to a specific location
py cli.py C:\Users\Lenovo\Documents\notes.md -o C:\Users\Lenovo\Desktop\notes.docx

# Use the short -o flag
py cli.py myfile.md -o output\myfile.docx
```

---

## Supported Markdown Features

| Feature | Syntax | Example |
|---------|--------|---------|
| Heading 1 | `# Title` | Large blue heading |
| Heading 2 | `## Section` | Medium blue heading |
| Heading 3–6 | `### Sub` | Smaller headings |
| Bold | `**text**` | **text** |
| Italic | `*text*` | *text* |
| Bold + Italic | `***text***` | ***text*** |
| Strikethrough | `~~text~~` | ~~text~~ |
| Inline code | `` `code` `` | red monospace |
| Code block | ` ```lang ` | grey background block |
| Unordered list | `- item` | bullet points |
| Ordered list | `1. item` | numbered list |
| Blockquote | `> text` | left blue border |
| Table | `\| col \| col \|` | styled with header |
| Horizontal rule | `---` | thin grey line |
| Link | `[text](url)` | blue underlined |

---

## Troubleshooting

**"py" is not recognised**  
→ Try `python app.py` instead of `py app.py`

**Browser shows "This site can't be reached"**  
→ The server is not running. Go back to Step 3 and start it again.

**Port 5000 already in use**  
→ Another program is using port 5000. Edit the last line of `app.py` and change the port:
```python
app.run(debug=True, port=5001)   # change 5000 to any free port
```
Then access `http://127.0.0.1:5001` instead.

**Downloaded file won't open**  
→ Make sure Microsoft Word or a compatible app (LibreOffice, Google Docs) is installed.
