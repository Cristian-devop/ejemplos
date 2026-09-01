"""
app.py - Frontend web para el convertidor doc → Markdown
Ejecutar:  python app.py
"""

import os
import shutil
import zipfile
from pathlib import Path
from flask import (
    Flask, render_template, request,
    send_file, jsonify, after_this_request
)
from werkzeug.utils import secure_filename

from doc_to_md import docx_to_markdown, doc_to_md_mammoth, pdf_to_markdown

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB máx

UPLOAD_FOLDER = Path("uploads")
OUTPUT_FOLDER = Path("outputs")
UPLOAD_FOLDER.mkdir(exist_ok=True)
OUTPUT_FOLDER.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".doc", ".docx", ".pdf"}


def allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No se recibió ningún archivo."}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Nombre de archivo vacío."}), 400

    filename = secure_filename(file.filename)
    if not allowed_file(filename):
        return jsonify({"error": "Solo se permiten archivos .doc, .docx o .pdf"}), 400

    use_mammoth = request.form.get("mammoth") == "true"

    # Guardar archivo subido
    input_path = UPLOAD_FOLDER / filename
    file.save(str(input_path))

    stem = Path(filename).stem
    output_dir = OUTPUT_FOLDER / stem
    output_dir.mkdir(exist_ok=True)
    output_md = output_dir / f"{stem}.md"

    try:
        ext = Path(filename).suffix.lower()
        if ext == ".pdf":
            pdf_to_markdown(input_path, output_md)
        elif use_mammoth or ext == ".doc":
            doc_to_md_mammoth(input_path, output_md)
        else:
            docx_to_markdown(input_path, output_md)
    except Exception as e:
        return jsonify({"error": f"Error al convertir: {str(e)}"}), 500
    finally:
        input_path.unlink(missing_ok=True)

    # Si hay imágenes, empaquetar en ZIP; si no, devolver el .md
    images_dir = output_dir / "images"
    has_images = images_dir.exists() and any(images_dir.iterdir())

    if has_images:
        zip_path = OUTPUT_FOLDER / f"{stem}.zip"
        with zipfile.ZipFile(zip_path, "w") as zf:
            zf.write(output_md, output_md.name)
            for img in images_dir.iterdir():
                zf.write(img, f"images/{img.name}")

        @after_this_request
        def cleanup(response):
            shutil.rmtree(output_dir, ignore_errors=True)
            zip_path.unlink(missing_ok=True)
            return response

        return send_file(
            zip_path,
            as_attachment=True,
            download_name=f"{stem}.zip",
            mimetype="application/zip",
        )
    else:
        md_content = output_md.read_text(encoding="utf-8")

        @after_this_request
        def cleanup(response):
            shutil.rmtree(output_dir, ignore_errors=True)
            return response

        return jsonify({"markdown": md_content, "filename": f"{stem}.md"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
