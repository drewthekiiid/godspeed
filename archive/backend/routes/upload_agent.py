from flask import Blueprint, request, jsonify
import os
import smartsheet
from werkzeug.utils import secure_filename
import io
import base64

# Instantiate the blueprint
upload_bp = Blueprint("upload", __name__)

# Initialize Smartsheet client
ss_client = smartsheet.Smartsheet(os.getenv("SMARTSHEET_API_KEY"))

@upload_bp.route("/smartsheet/upload", methods=["POST"])
def upload_to_smartsheet():
    try:
        data = request.get_json()
        sheet_id = data.get("sheet_id")
        filename = data.get("filename", "Estimate.docx")
        file_data_b64 = data.get("file_data")
        row_id = data.get("row_id")  # Optional

        if not sheet_id or not file_data_b64:
            return jsonify({"error": "Missing required fields: sheet_id or file_data"}), 400

        # Decode the Base64 file data
        file_bytes = base64.b64decode(file_data_b64)

        # Attach to existing row if given, otherwise create one
        if row_id:
            target_row_id = row_id
        else:
            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            row_response = ss_client.Sheets.add_rows(sheet_id, [new_row])
            if row_response.message != "SUCCESS":
                return jsonify({"error": f"Row creation failed: {row_response.message}"}), 500
            target_row_id = row_response.data[0].id

        # Upload the file to that row
        upload_response = ss_client.Attachments.attach_file_to_row(
            sheet_id,
            target_row_id,
            (
                secure_filename(filename),
                io.BytesIO(file_bytes),
                'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            )
        )

        if upload_response.message != "SUCCESS":
            return jsonify({"error": f"Attachment upload failed: {upload_response.message}"}), 500

        return jsonify({
            "row_id": target_row_id,
            "attachment_id": upload_response.result.id,
            "message": "Upload successful"
        })

    except Exception as e:
        return jsonify({"error": f"Upload failed: {str(e)}"}), 500
