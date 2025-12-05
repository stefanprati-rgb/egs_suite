import re
import unicodedata
import hashlib

HTML_TAGS_RE = re.compile(r"<[^>]+>")

def normaliza(s: str) -> str:
    if not s: return ""
    try:
        s = HTML_TAGS_RE.sub(" ", s); s = s.lower(); s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        s = re.sub(r"\s+", " ", s).strip()
    except UnicodeDecodeError:
        s = s.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')
    return s

def _to_bytes(x):
    if isinstance(x, (bytes, bytearray)): return bytes(x)
    try: return bytes(x)
    except Exception: return None
    
def _iso(dt):
    return dt.strftime("%Y-%m-%dT%H:%M:%S")

def hash_bytes(b): return hashlib.sha256(b).hexdigest()
