import fitz  # PyMuPDF
import pandas as pd
import docx
import io

def extract_text(file_bytes: bytes, mime_type: str) -> str:
    file_stream = io.BytesIO(file_bytes)
    try:
        if "pdf" in mime_type:
            doc = fitz.open(stream=file_stream.read(), filetype="pdf")
            return "\n".join([page.get_text() for page in doc])
        elif "text" in mime_type:
            return file_stream.read().decode("utf-8")
        elif "spreadsheet" in mime_type or "excel" in mime_type:
            df = pd.read_excel(file_stream)
            return df.to_string(index=False)
        elif "wordprocessingml" in mime_type or "msword" in mime_type:
            docx_obj = docx.Document(file_stream)
            return "\n".join([p.text.strip() for p in docx_obj.paragraphs if p.text.strip()])
        else:
            raise ValueError(f"Unsupported MIME type: {mime_type}")
    except Exception as e:
        raise ValueError(f"Parsing failed: {str(e)}")