from flask import Blueprint, request, jsonify
import os
import requests
import smartsheet
from services.utils.parser import extract_text
from services.utils.compression import compress_if_large
from services.gpt_handler import analyze_with_gpt

smartsheet_bp = Blueprint("smartsheet", __name__)
ss_client = smartsheet.Smartsheet(os.getenv("SMARTSHEET_API_KEY"))

@smartsheet_bp.route("/list-attachments", methods=["POST"])
def list_attachments():
    data = request.get_json()
    url = data.get("sheet_url")
    if not url or "/sheets/" not in url:
        return jsonify({"error": "Invalid Smartsheet URL."}), 400

    try:
        sheet_id = url.split("/sheets/")[-1].split("?")[0].strip()
        sheet = ss_client.Sheets.get_sheet(sheet_id)
        result = {}

        for row in sheet.rows:
            try:
                attachments = ss_client.Attachments.list_row_attachments(sheet.id, row.id).data
                if attachments:
                    result[row.id] = [
                        {"name": a.name, "id": a.id, "mime_type": a.mime_type} for a in attachments
                    ]
            except Exception:
                continue

        return jsonify({"sheet_id": sheet.id, "attachments": result})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@smartsheet_bp.route("/analyze-attachment/<sheet_id>/<attachment_id>", methods=["GET"])
def analyze_attachment(sheet_id, attachment_id):
    try:
        attachment = ss_client.Attachments.get_attachment(sheet_id, attachment_id)
        file_res = requests.get(attachment.url)
        file_bytes = file_res.content
        file_bytes, _, _ = compress_if_large(file_bytes, attachment.name)
        parsed_text = extract_text(file_bytes, attachment.mime_type)
        result = analyze_with_gpt("Analyze this document:", parsed_text)
        return jsonify({"analysis": result["response"], "usage": result["usage"]})
    except Exception as e:
        return jsonify({"error": f"Smartsheet agent error: {str(e)}"}), 500


@smartsheet_bp.route("/analyze-batch", methods=["POST"])
def analyze_batch_attachments():
    try:
        data = request.get_json()
        print("analyze-batch incoming data:", data)  # Debug line

        # Accept both possible field names for compatibility
        sheet_id = data.get("sheet_id")
        if not sheet_id and "sheet_url" in data:
            # Parse from URL
            try:
                sheet_id = data["sheet_url"].split("/sheets/")[-1].split("?")[0].strip()
            except Exception:
                sheet_id = None

        attachment_ids = data.get("attachment_ids", [])
        additional_text = data.get("additional_text", "") or data.get("notes", "")
        query = data.get("query") or data.get("notes", "") or "Analyze these attachments:"

        if not sheet_id or not attachment_ids:
            return jsonify({"error": "Missing required fields: sheet_id or attachment_ids"}), 400

        # For now, if query is still blank, just use notes or fallback text
        if not query.strip():
            query = "Analyze these attachments:"

        full_text = additional_text.strip() + "\n\n"

        for att_id in attachment_ids:
            att = ss_client.Attachments.get_attachment(sheet_id, att_id)
            file_res = requests.get(att.url)
            file_bytes = file_res.content
            file_bytes, _, _ = compress_if_large(file_bytes, att.name)
            parsed = extract_text(file_bytes, att.mime_type)
            full_text += f"\n\n--- File: {att.name} ---\n\n{parsed}"

        result = analyze_with_gpt(query, full_text)
        return jsonify(result)

    except Exception as e:
        return jsonify({"error": f"Batch analysis failed: {str(e)}"}), 500
    
@smartsheet_bp.route("/list-rows", methods=["POST"])
def list_rows():
    try:
        data = request.get_json()
        sheet_id = data.get("sheet_id")

        if not sheet_id:
            return jsonify({"error": "Missing sheet_id"}), 400

        sheet = ss_client.Sheets.get_sheet(sheet_id)
        project_rows = []

        for row in sheet.rows:
            # Try to pull something human-readable from first non-blank cell
            name = None
            for cell in row.cells:
                if cell.value:
                    name = str(cell.value)
                    break
            project_rows.append({
                "id": row.id,
                "name": name or f"Row {row.id}"
            })

        return jsonify({"rows": project_rows})

    except Exception as e:
        return jsonify({"error": f"Row listing failed: {str(e)}"}), 500
    
@smartsheet_bp.route("/upload", methods=["POST"])
def upload_to_smartsheet():
    try:
        data = request.get_json()
        row_id = data.get("row_id")
        filename = data.get("filename")
        file_buffer = data.get("file_buffer")  # Base64 encoded

        if not row_id or not filename or not file_buffer:
            return jsonify({"error": "Missing fields: row_id, filename, or file_buffer"}), 400

        import base64
        file_bytes = base64.b64decode(file_buffer)

        upload_response = ss_client.Attachments.attach_file_to_row(
            row_id,
            (
                filename,
                io.BytesIO(file_bytes),
                "application/octet-stream"
            )
        )

        if upload_response.message != "SUCCESS":
            return jsonify({"error": "Smartsheet API upload failed"}), 500

        return jsonify({
            "success": True,
            "attachment_id": upload_response.result.id
        })

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500

