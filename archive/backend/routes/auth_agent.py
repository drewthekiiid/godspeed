from flask import Blueprint, request, jsonify
import os
from dotenv import load_dotenv

load_dotenv()

auth_bp = Blueprint("auth", __name__)

@auth_bp.route("/auth", methods=["POST"])
def auth():
    data = request.get_json()
    if data.get("code") == os.getenv("INTERNAL_SECRET_CODE"):
        return jsonify({"status": "ok"}), 200
    return jsonify({"error": "invalid code"}), 403

@auth_bp.route("/health", methods=["GET"])
def healthcheck():
    return jsonify({"status": "ok"})