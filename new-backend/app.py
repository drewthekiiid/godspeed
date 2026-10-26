from flask import Flask, request, jsonify
import os
from dotenv import load_dotenv
from flask_cors import CORS  # CORS for cross-origin requests

# === Load Environment ===
load_dotenv()
SECRET_CODE = os.getenv("INTERNAL_SECRET_CODE")
print(f"Loaded INTERNAL_SECRET_CODE: {SECRET_CODE}")

# === Initialize App ===
app = Flask(__name__)
CORS(app, supports_credentials=True)  # Enable CORS on all routes

# === Register Flask Blueprints (Agents) ===
from routes.auth_agent import auth_bp
from routes.chat_agent import chat_bp
from routes.smartsheet_agent import smartsheet_bp
from routes.estimate_agent import estimate_bp
from routes.export_agent import export_bp
from routes.upload_agent import upload_bp

app.register_blueprint(auth_bp)                                 # /auth, /health
app.register_blueprint(chat_bp, url_prefix="/chat")             # /chat/analyze, /chat/upload-analyze
app.register_blueprint(smartsheet_bp, url_prefix="/smartsheet") # /smartsheet/...
app.register_blueprint(estimate_bp, url_prefix="/estimate")     # /estimate/generate
app.register_blueprint(export_bp, url_prefix="/estimate")       # /estimate/export
app.register_blueprint(upload_bp)                               # /smartsheet/upload

# === Auth Middleware ===
@app.before_request
def auth_check():
    exempt_endpoints = (
        "auth.auth",           # POST /auth
        "auth.healthcheck",    # GET /health
        "static"               # Needed for local dev/static
    )
    # --- NEW: Allow all OPTIONS preflight requests ---
    if request.method == "OPTIONS":
        return  # Just let Flask-CORS handle it
    # ------------------------------------------------
    print(f"\n--- Request: {request.method} {request.path}")
    print("Headers:", dict(request.headers))
    if request.endpoint not in exempt_endpoints:
        access_code = request.headers.get("X-Internal-Access")
        print("X-Internal-Access:", access_code)
        if not access_code or access_code != SECRET_CODE:
            print("❌ Unauthorized: Invalid internal access code.")
            return jsonify({"error": "Unauthorized. Invalid internal access code."}), 403


# === Run Server ===
if __name__ == "__main__":
    print(f"SECRET_CODE LOADED: {SECRET_CODE}")
    app.run(host="127.0.0.1", port=5001, debug=True)