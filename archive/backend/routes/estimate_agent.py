from flask import Blueprint, request, jsonify, send_file
import io
from docx import Document
from fpdf import FPDF  # fpdf2 installed
import tempfile

estimate_bp = Blueprint("estimate", __name__)

@estimate_bp.route("/export", methods=["POST"])
def export_estimate():
    data = request.get_json()
    content = data.get("content", "")
    file_format = data.get("format", "pdf")

    if not content or file_format not in ["pdf", "docx"]:
        return jsonify({"error": "Missing content or invalid format"}), 400

    try:
        if file_format == "pdf":
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Helvetica", size=12)
            for line in content.splitlines():
                pdf.multi_cell(0, 10, line)
            temp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
            pdf.output(temp.name)
            temp.close()
            return send_file(temp.name, mimetype="application/pdf", as_attachment=True, download_name="estimate.pdf")

        elif file_format == "docx":
            doc = Document()
            for line in content.splitlines():
                doc.add_paragraph(line)
            temp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
            doc.save(temp.name)
            temp.close()
            return send_file(temp.name, mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document", as_attachment=True, download_name="estimate.docx")

    except Exception as e:
        return jsonify({"error": f"Export failed: {str(e)}"}), 500

    return jsonify({"error": "Unknown export format"}), 400
