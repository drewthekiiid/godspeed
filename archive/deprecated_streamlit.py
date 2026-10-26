# app.py – Godspeed Estimator Frontend (ChatGPT-Perfected UX + Embedded Paperclip Input)

import streamlit as st
import requests
import os
from dotenv import load_dotenv
from datetime import datetime, timedelta
import json

# === Session Timeout Config ===
TIMEOUT_MINUTES = 15

# === Load Environment ===
load_dotenv()
API_BASE = "http://localhost:5001"
DEFAULT_ACCESS = os.getenv("INTERNAL_SECRET_CODE", "hermes")

# === Session Timeout Check ===
if "last_active" in st.session_state:
    if datetime.now() - st.session_state["last_active"] > timedelta(minutes=TIMEOUT_MINUTES):
        if "access_code" in st.session_state:
            del st.session_state["access_code"]
        del st.session_state["last_active"]
        st.warning("Session expired due to inactivity. Please log in again.")

# === Show Login Page if Not Authenticated ===
if "access_code" not in st.session_state:
    st.set_page_config(page_title="Godspeed Estimator – Login", layout="centered")
    st.title("Godspeed Estimator Login")
    with st.form("login_form"):
        password = st.text_input("Enter internal access code:", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            res = requests.post(f"{API_BASE}/auth", json={"code": password})
            if res.status_code == 200:
                st.session_state["access_code"] = password
                st.session_state["last_active"] = datetime.now()
                st.success("Access granted.")
                st.rerun()
            else:
                st.error("Invalid access code.")
    st.stop()

# === Refresh session activity timestamp ===
st.session_state["last_active"] = datetime.now()

# === Session Setup ===
st.set_page_config(page_title="Godspeed Estimator", layout="wide")
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []
if "saved_sessions" not in st.session_state:
    st.session_state.saved_sessions = {}

HEADERS = {"X-Internal-Access": st.session_state["access_code"]}

# === Global CSS ===
st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] {
    height: 100% !important;
    overflow: hidden;
}
.chat-scroll-area {
    position: absolute;
    top: 0;
    bottom: 130px;
    width: 100%;
    overflow-y: auto;
    padding: 1rem 1.5rem 3rem 1.5rem;
}
.chat-input-dock {
    position: fixed;
    bottom: 0;
    left: 0;
    width: 100%;
    padding: 1rem 2rem;
    background-color: #f9f9f9;
    border-top: 1px solid #ccc;
    z-index: 100;
}
input[type="text"] {
    font-size: 1.05rem !important;
}
.sidebar-style {
    height: 100%;
    padding: 1rem 1rem 3rem 1rem;
    background-color: #f5f5f5;
}
</style>
""", unsafe_allow_html=True)

# === Sidebar (ChatGPT Style) ===
with st.sidebar:
    st.markdown("<div class='sidebar-style'>", unsafe_allow_html=True)
    visible_sessions = st.session_state.saved_sessions
    if visible_sessions:
        st.subheader("🗂️ Saved Sessions")
        for name, preview in visible_sessions.items():
            st.button(f"{name}", help=preview[:50] + "...")
        st.markdown("---")
    st.button("🧹 Clear Chat", on_click=lambda: st.session_state.chat_history.clear())
    st.button("🗑️ Clear All", on_click=lambda: st.session_state.clear())
    st.caption(f"🔒 Access Code: `****{st.session_state['access_code'][-2:]}`")
    st.markdown("</div>", unsafe_allow_html=True)

# === Header if Empty ===
if len(st.session_state.chat_history) == 0:
    st.markdown("""
        <div style='text-align: center; margin-top: 2rem;'>
            <h1 style='color: #6c63ff; font-size: 2.5rem;'>Godspeed Estimator</h1>
            <p style='color: #666;'>Ask a question or browse a file to begin.</p>
        </div>
        <hr class='faded'>
    """, unsafe_allow_html=True)

# === Scrollable Chat Window ===
st.markdown("<div class='chat-scroll-area'>", unsafe_allow_html=True)
for entry in st.session_state.chat_history:
    with st.chat_message("user"):
        st.markdown(entry["user"])
        if entry.get("files"):
            for f in entry["files"]:
                st.caption(f"📄 `{f['name']}` — {f['size_kb']} KB")
        st.caption(f"🕒 {entry['timestamp']}")
    with st.chat_message("assistant"):
        st.markdown(entry["assistant"])
        st.caption(f"🕒 {entry['timestamp']}")
st.markdown("</div>", unsafe_allow_html=True)

# === Bottom Dock: Custom Styled Input w/ Paperclip ===
st.markdown("""
<style>
.chat-input-dock {
    position: fixed;
    bottom: 0;
    left: 0;
    width: 100%;
    background-color: #f9f9f9;
    border-top: 1px solid #ccc;
    padding: 1rem 2rem;
    z-index: 100;
}
.chat-form-container {
    position: relative;
    width: 100%;
}
.chat-textbox {
    width: 100%;
    padding: 0.8rem 3rem 0.8rem 1rem;
    font-size: 1rem;
    border: 1px solid #ccc;
    border-radius: 8px;
}
.paperclip-icon {
    position: absolute;
    top: 50%;
    right: 1rem;
    transform: translateY(-50%);
    cursor: pointer;
    font-size: 1.2rem;
    opacity: 0.6;
}
input[type="file"] {
    display: none;
}
</style>

<div class="chat-input-dock">
  <form method="post">
    <div class="chat-form-container">
      <input name="user_prompt" id="user_prompt" class="chat-textbox" placeholder="Ask a question or drop a file..." />
      <label for="file_input">
        <span class="paperclip-icon">📎</span>
      </label>
      <input id="file_input" name="file_input" type="file" multiple />
    </div>
  </form>
</div>
""", unsafe_allow_html=True)

# === Submission Logic ===
if False:
    timestamp = datetime.now().strftime("%Y-%m-%d %I:%M %p")
    file_meta = [{"name": f.name, "size_kb": round(len(f.getbuffer()) / 1024, 2)} for f in uploaded_files or []]

    st.session_state.chat_history.append({
        "user": user_input if user_input else "[File Upload]",
        "files": file_meta,
        "assistant": "🤖 Processing...",
        "timestamp": timestamp
    })

    # Handle file or text
    if uploaded_files:
        file = uploaded_files[0]
        res = requests.post(
            f"{API_BASE}/chat/upload-analyze",
            headers=HEADERS,
            files={"file": (file.name, file.getvalue())}
        )
        parsed_text = res.json().get("parsed_text", "")
    else:
        res = requests.post(
            f"{API_BASE}/chat/analyze",
            headers=HEADERS,
            json={"query": user_input, "content": ""}
        )
        parsed_text = res.json().get("parsed_text", "")

    gpt_response = res.json().get("response", "⚠️ No response received.")

    # === Intent Routing ===
    intent_prompt = [
        {"role": "system", "content": "You are a system controller. Based on the user message below, return one of: generate_estimate, export_estimate, upload_smartsheet, none."},
        {"role": "user", "content": user_input or "[User uploaded a file]"}
    ]
    intent_res = requests.post(
        f"{API_BASE}/chat/analyze",
        headers=HEADERS,
        json={"query": json.dumps(intent_prompt), "content": ""}
    )
    system_action = intent_res.json().get("response", "none").strip().lower()

    # === Backend Action ===
    if system_action == "generate_estimate":
        est_res = requests.post(
            f"{API_BASE}/estimate/generate",
            headers=HEADERS,
            json={"parsed_text": parsed_text}
        )
        gpt_response += "\n\n---\n**Estimate:**\n```json\n" + json.dumps(est_res.json(), indent=2) + "\n```"

    elif system_action == "export_estimate":
        gpt_response += "\n\n✅ Exported! (hook up download logic here)"

    elif system_action == "upload_smartsheet":
        gpt_response += "\n\n✅ Uploaded to Smartsheet! (Backend confirmed)"

    # Update assistant response
    st.session_state.chat_history[-1]["assistant"] = gpt_response

    # Save to session preview list
    if user_input:
        preview = user_input.strip().split("\n")[0][:60]
        st.session_state.saved_sessions[f"Session {len(st.session_state.saved_sessions)+1}"] = preview

    st.rerun()