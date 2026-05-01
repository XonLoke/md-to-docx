"""
Flask web application for Markdown → DOCX conversion.
"""

import os
import uuid
from flask import (Flask, render_template, request,
                   send_file, jsonify, after_this_request)
from werkzeug.utils import secure_filename
from converter import convert_markdown_to_docx
import io

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 10 * 1024 * 1024   # 10 MB upload limit
UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    """
    Accepts either:
      - JSON body  { "markdown": "...", "filename": "output" }
      - multipart form with a .md file upload
    Returns the generated .docx as a file download.
    """
    filename = "output"
    markdown_text = None

    # ── JSON / text body ─────────────────────────────────────────────────────
    if request.is_json:
        data = request.get_json(silent=True) or {}
        markdown_text = data.get("markdown", "")
        filename      = data.get("filename", "output") or "output"

    # ── File upload ──────────────────────────────────────────────────────────
    elif "file" in request.files:
        f = request.files["file"]
        if f.filename == "":
            return jsonify({"error": "No file selected"}), 400
        if not f.filename.lower().endswith(".md"):
            return jsonify({"error": "Only .md files are accepted"}), 400
        markdown_text = f.read().decode("utf-8", errors="replace")
        filename      = os.path.splitext(secure_filename(f.filename))[0]

    # ── Form text area ────────────────────────────────────────────────────────
    elif "markdown" in request.form:
        markdown_text = request.form.get("markdown", "")
        filename      = request.form.get("filename", "output") or "output"

    else:
        return jsonify({"error": "No markdown content provided"}), 400

    if not markdown_text.strip():
        return jsonify({"error": "Markdown content is empty"}), 400

    # ── Convert ───────────────────────────────────────────────────────────────
    try:
        docx_bytes = convert_markdown_to_docx(markdown_text)
    except Exception as e:
        return jsonify({"error": f"Conversion failed: {str(e)}"}), 500

    # ── Stream back ───────────────────────────────────────────────────────────
    buf = io.BytesIO(docx_bytes)
    buf.seek(0)
    safe_name = secure_filename(filename) or "output"
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"{safe_name}.docx",
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    print("✅  Markdown → DOCX converter running at http://127.0.0.1:5000")
    app.run(debug=True, port=5000)
