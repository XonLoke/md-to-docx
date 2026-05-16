"""
Convert Sprint Breakdown v3.md to docx
"""
import sys
sys.path.insert(0, r"D:\sep_venv\md-to-docx")
from converter import convert_markdown_to_docx

input_path = r"D:\Developer Courseware\_sem4term1-RP-C3000C-Capstone Projects\Sprint Management\Sprint Breakdown v3.md"
output_path = r"D:\Developer Courseware\_sem4term1-RP-C3000C-Capstone Projects\Sprint Management\Sprint Breakdown v3.docx"

with open(input_path, "r", encoding="utf-8") as f:
    md_content = f.read()

docx_bytes = convert_markdown_to_docx(md_content)

with open(output_path, "wb") as f:
    f.write(docx_bytes)

print(f"Converted: {input_path}")
print(f"Saved to: {output_path}")
