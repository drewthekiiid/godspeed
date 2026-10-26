from flask import Blueprint, request, jsonify
from services.utils.parser import extract_text
from services.utils.compression import compress_if_large
from services.gpt_handler import analyze_with_gpt

chat_bp = Blueprint("chat", __name__)

@chat_bp.route("/analyze", methods=["POST"])
def analyze_chat():
    data = request.json
    content = data.get("content")
    query = data.get("query")
    if not query and not content:  # Allow either to be present
        return jsonify({"error": "Missing 'content' or 'query'."}), 400
    try:
        result = analyze_with_gpt(query or "", content or "")
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"OpenAI error: {str(e)}"}), 502

@chat_bp.route("/upload-analyze", methods=["POST"])
def upload_analyze():
    query = request.form.get("query", "")
    uploaded_files = request.files.getlist("files")

    if not uploaded_files or not query:
        return jsonify({"error": "Missing 'files' or 'query'"}), 400

    try:
        texts = []

        for file in uploaded_files:
            mime_type = file.mimetype
            file_bytes = file.read()

            # Compress if needed (for metrics or future zip fallback)
            file_bytes, _, _ = compress_if_large(file_bytes, file.filename)

            # Parse text
            parsed_text = extract_text(file_bytes, mime_type)
            texts.append(f"\n\n--- File: {file.filename} ---\n\n{parsed_text}")

        full_text = "\n\n".join(texts)
        result = analyze_with_gpt(query, full_text)

        return jsonify(result)

    except Exception as e:
        return jsonify({"error": f"Multi-file analysis failed: {str(e)}"}), 500