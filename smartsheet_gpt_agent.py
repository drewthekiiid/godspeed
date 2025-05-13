from flask import Flask, request, jsonify
import requests
import os
import time
import io
import zipfile
from dotenv import load_dotenv
from openai import OpenAI

import fitz  # PyMuPDF for PDFs
import pandas as pd
import docx  # python-docx for Word documents
import smartsheet

load_dotenv()

app = Flask(__name__)

SMARTSHEET_API_KEY = os.getenv("SMARTSHEET_API_KEY")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
SECRET_CODE = os.getenv("INTERNAL_SECRET_CODE")
HEADERS = {"Authorization": f"Bearer {SMARTSHEET_API_KEY}"}
BASE_URL = "https://api.smartsheet.com/2.0"

ss_client = smartsheet.Smartsheet(SMARTSHEET_API_KEY)
openai = OpenAI(api_key=OPENAI_API_KEY)

def fetch_json(url, method="GET", **kwargs):
    try:
        res = requests.get(url, headers=HEADERS, **kwargs) if method == "GET" else requests.post(url, headers=HEADERS, **kwargs)
    except requests.exceptions.RequestException as e:
        return None, jsonify({"error": str(e)}), 502
    if not res.ok:
        return None, jsonify({"error": res.text}), res.status_code
    return res.json(), None, None

def compress_if_large(file_bytes, filename):
    if len(file_bytes) > 75 * 1024 * 1024:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr(filename, file_bytes)
        return buf.getvalue(), 'application/zip', f"{filename}.zip"
    return file_bytes, None, None

@app.before_request
def auth_check():
    if request.endpoint not in ("healthcheck",):
        if request.headers.get("X-Internal-Access") != SECRET_CODE:
            return jsonify({"error": "Unauthorized. Invalid internal access code."}), 403

@app.route("/test-download/<attachment_id>", methods=["GET"])
def test_download_debug(attachment_id):
    url = f"{BASE_URL}/attachments/{attachment_id}/download"
    try:
        res = requests.get(url, headers=HEADERS, allow_redirects=False)
        debug_info = {
            "status_code": res.status_code,
            "headers": dict(res.headers),
            "text_snippet": res.text[:200]
        }
        return jsonify(debug_info)
    except requests.exceptions.RequestException as e:
        return jsonify({"error": str(e)}), 502

@app.route("/debug-attachment/<sheet_id>/<attachment_id>", methods=["GET"])
def fallback_get_attachment(sheet_id, attachment_id):
    try:
        att = ss_client.Attachments.get_attachment(sheet_id, attachment_id)
        return jsonify({
            "name": att.name,
            "type": att.mime_type,
            "url": att.url,
            "size": att.size_in_kb
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/analyze", methods=["POST"])
def analyze_content():
    data = request.json
    content = data.get("content")
    query = data.get("query")
    if not content or not query:
        return jsonify({"error": "Both 'content' and 'query' are required."}), 400
    try:
        chat_response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a document analyst."},
                {"role": "user", "content": f"""{query}\n\n{content[:8000]}"""}
            ]
        )
        result = chat_response.choices[0].message.content
        return jsonify({"response": result, "usage": chat_response.usage})
    except Exception as e:
        return jsonify({"error": f"OpenAI error: {str(e)}"}), 502

@app.route("/analyze-attachment/<sheet_id>/<attachment_id>", methods=["GET"])
def analyze_attachment(sheet_id, attachment_id):
    try:
        attachment = ss_client.Attachments.get_attachment(sheet_id, attachment_id)
        if not hasattr(attachment, 'url') or not attachment.url:
            return jsonify({"error": "Attachment object missing URL."}), 404

        file_res = requests.get(attachment.url)
        file_bytes = file_res.content
        file_bytes, compressed_type, compressed_name = compress_if_large(file_bytes, attachment.name)
        if compressed_type:
            return jsonify({"note": "File was compressed due to size.", "filename": compressed_name}), 200

    except Exception as e:
        return jsonify({"error": f"Attachment fetch failed: {str(e)}"}), 400

    content_type = attachment.mime_type
    file_stream = io.BytesIO(file_bytes)

    try:
        if "pdf" in content_type:
            doc = fitz.open(stream=file_stream.read(), filetype="pdf")
            raw_text = "\n".join([page.get_text() for page in doc])
        elif "text" in content_type:
            raw_text = file_stream.read().decode("utf-8")
        elif "spreadsheet" in content_type or "excel" in content_type:
            df = pd.read_excel(file_stream)
            raw_text = df.to_string(index=False)
        elif "wordprocessingml" in content_type or "msword" in content_type:
            docx_obj = docx.Document(file_stream)
            raw_text = "\n".join([para.text.strip() for para in docx_obj.paragraphs if para.text.strip()])
        else:
            return jsonify({"error": f"Unsupported file type: {content_type}"}), 415
    except Exception as e:
        return jsonify({"error": f"Document parsing failed: {str(e)}"}), 422

    if not raw_text:
        return jsonify({"error": "Failed to extract text."}), 422

    try:
        chat_response = openai.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a document analyst."},
                {"role": "user", "content": f"""Analyze this document:\n\n{raw_text[:8000]}"""}
            ]
        )
        analysis = chat_response.choices[0].message.content
        return jsonify({"analysis": analysis})
    except Exception as e:
        return jsonify({"error": f"OpenAI error: {str(e)}"}), 502

@app.route("/health", methods=["GET"])
def healthcheck():
    return jsonify({"status": "ok"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)

