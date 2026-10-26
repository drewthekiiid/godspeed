import io
import zipfile

def compress_if_large(file_bytes: bytes, filename: str):
    if len(file_bytes) > 75 * 1024 * 1024:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zipf:
            zipf.writestr(filename, file_bytes)
        return buf.getvalue(), 'application/zip', f"{filename}.zip"
    return file_bytes, None, None