#!/usr/bin/env python3
"""Flask backend for Drawio to PPTX conversion service."""

import os
import tempfile
import uuid
from pathlib import Path

from flask import Flask, request, send_file, jsonify, send_from_directory
from flask_cors import CORS

from drawio_to_pptx import convert

BASE_DIR = Path(__file__).resolve().parent

app = Flask(__name__, static_folder=str(BASE_DIR), static_url_path="")
CORS(app)

# Temp directory for intermediate files
TEMP_DIR = Path(tempfile.gettempdir()) / "drawio_to_pptx"
TEMP_DIR.mkdir(parents=True, exist_ok=True)


@app.route("/")
def index():
    return send_from_directory(str(BASE_DIR), "index.html")


@app.route("/api/convert", methods=["POST"])
def convert_file():
    """Receive a .drawio file, convert to .pptx, and return the result."""
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Empty filename"}), 400

    ext = Path(file.filename).suffix.lower()
    if ext not in (".drawio", ".dio", ".xml"):
        return jsonify({"error": f"Unsupported file type: {ext}"}), 400

    # Save uploaded file to temp
    task_id = uuid.uuid4().hex[:8]
    input_path = TEMP_DIR / f"{task_id}_input{ext}"
    output_path = TEMP_DIR / f"{task_id}_output.pptx"

    try:
        file.save(str(input_path))
        convert(input_path, output_path, fix_mojibake=True)

        if not output_path.exists():
            return jsonify({"error": "Conversion produced no output"}), 500

        output_name = Path(file.filename).stem + ".pptx"
        return send_file(
            str(output_path),
            mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            as_attachment=True,
            download_name=output_name,
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        # Cleanup
        try:
            input_path.unlink(missing_ok=True)
            output_path.unlink(missing_ok=True)
        except Exception:
            pass


@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8765)))
