#!/bin/bash

AUTH_HEADER="X-Internal-Access: hermes"
CONTENT_TYPE="Content-Type: application/json"
SHEET_ID="6058615036137348"
ATTACHMENT_1="3939186555785092"
ATTACHMENT_2="5628036416049028"

echo "===== Godspeed Estimator: Test Sequence ====="
read -p "Press [Enter] to start..."

function pause() {
  echo ""
  read -p "Press [Enter] to continue..."
  echo ""
}

echo "[1] Healthcheck"
curl http://localhost:5001/health
pause

echo "[2] Auth Test"
curl -X POST http://localhost:5001/auth \
  -H "$CONTENT_TYPE" \
  -d '{"secret_code": "hermes"}'
pause

echo "[3] GPT Text Analysis"
curl -X POST http://localhost:5001/chat/analyze \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d '{"query": "Summarize this document", "content": "This is a test spec for drywall, electrical, and plumbing."}'
pause

echo "[4] File Upload + GPT Analysis"
curl -X POST http://localhost:5001/chat/upload-analyze \
  -H "$AUTH_HEADER" \
  -F "query=Extract the scope of work" \
  -F "file=@./estimate.docx"
pause

echo "[5] List Smartsheet Attachments"
curl -X POST http://localhost:5001/smartsheet/list-attachments \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d "{\"sheet_url\": \"https://app.smartsheet.com/sheets/$SHEET_ID\"}"
pause

echo "[6] Analyze a Single Attachment"
curl -X GET http://localhost:5001/smartsheet/analyze-attachment/$SHEET_ID/$ATTACHMENT_1 \
  -H "$AUTH_HEADER"
pause

echo "[7] Analyze Multiple Attachments + Notes"
curl -X POST http://localhost:5001/smartsheet/analyze-batch \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d "{\"sheet_id\": \"$SHEET_ID\", \"attachment_ids\": [$ATTACHMENT_1, $ATTACHMENT_2], \"additional_text\": \"Also include ADA compliance.\", \"query\": \"Extract a combined scope of work\"}"
pause

echo "[8] Estimate Generator"
curl -X POST http://localhost:5001/estimate/generate \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d '{ "parsed_text": "Install 150 LF of drywall, 12 light fixtures, 1 ADA restroom." }'
pause

echo "[9] Export Estimate to DOCX"
curl -X POST http://localhost:5001/estimate/export \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d '{ "estimate": [ { "division": "09", "scope_item": "Drywall Partition", "quantity": 150, "unit": "LF", "unit_price": 12.5, "total": 1875.00 } ] }' --output estimate.docx
echo "File saved as: estimate.docx"
pause

echo "[10] Upload DOCX Back to Smartsheet"
BASE64_FILE=$(base64 -i estimate.docx)
curl -X POST http://localhost:5001/smartsheet/upload \
  -H "$AUTH_HEADER" \
  -H "$CONTENT_TYPE" \
  -d "{ \"sheet_id\": \"$SHEET_ID\", \"filename\": \"Estimate.docx\", \"file_data\": \"$BASE64_FILE\" }"
pause

echo "✅ All tests completed."