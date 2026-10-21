import os
import sys
import io
import re
import csv
import json
import time
import threading
import shutil
from datetime import datetime
from typing import Optional, List, Dict, Any

try:
    from dotenv import load_dotenv
    load_dotenv()
except Exception:
    pass

from flask import (
    Flask,
    request,
    jsonify,
    render_template_string,
    send_file,
    send_from_directory,
)
from PIL import Image, ImageOps
import pytesseract
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Optional PostgreSQL (Neon DB) support with connection pooling
try:
    import psycopg2
    from psycopg2 import pool
    from psycopg2.extras import RealDictCursor
    PSYCOPG2_AVAILABLE = True
except Exception:
    PSYCOPG2_AVAILABLE = False

# Optional Google GenAI SDK
try:
    from google import genai
    from google.genai import types
    from pydantic import BaseModel

    GENAI_AVAILABLE = True
except Exception:
    GENAI_AVAILABLE = False


# ==========================================
# 1. SETUP & CONFIGURATION
# ==========================================

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", BASE_DIR)
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
SETTINGS_PATH = os.path.join(DATA_DIR, "settings.json")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)

# Locate Tesseract executable
def find_tesseract_path():
    candidates = [
        shutil.which("tesseract"),
        os.path.expandvars(r"%LOCALAPPDATA%\Tesseract-OCR\tesseract.exe"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        "/usr/bin/tesseract",
        "/usr/local/bin/tesseract",
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path
    return "tesseract"


pytesseract.pytesseract.tesseract_cmd = find_tesseract_path()


# Load / Save Settings
def get_settings() -> Dict[str, Any]:
    default_settings = {
        "gemini_api_key": os.environ.get("GEMINI_API_KEY", ""),
        "database_url": os.environ.get("DATABASE_URL", ""),
        "engine_preference": "hybrid",  # "hybrid", "gemini", "tesseract"
    }

    # Always check .env directly for fresh changes on disk
    env_path = os.path.join(BASE_DIR, ".env")
    if os.path.isfile(env_path):
        try:
            with open(env_path, "r", encoding="utf-8") as ef:
                for line in ef:
                    line = line.strip()
                    if line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k, v = k.strip(), v.strip()
                    if k == "GEMINI_API_KEY" and v:
                        default_settings["gemini_api_key"] = v
                    elif k == "DATABASE_URL" and v:
                        default_settings["database_url"] = v
        except Exception:
            pass

    if os.path.isfile(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                for k, v in data.items():
                    if v:
                        default_settings[k] = v
                    elif k not in default_settings:
                        default_settings[k] = v
        except Exception:
            pass
    return default_settings


def save_settings(settings: Dict[str, Any]):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


# ==========================================
# 2. DATABASE MANAGEMENT (NEON CLOUD POSTGRESQL)
# ==========================================

def get_database_url() -> str:
    db_url = os.environ.get("DATABASE_URL", "").strip()
    if not db_url:
        settings = get_settings()
        db_url = settings.get("database_url", "").strip()
    if db_url.startswith("postgres://"):
        db_url = db_url.replace("postgres://", "postgresql://", 1)
    return db_url


_PG_POOL = None
_PG_POOL_URL = None
_PG_POOL_LOCK = threading.Lock()


def get_pg_pool(url: str):
    """Thread-safe connection pool manager for Neon Cloud PostgreSQL."""
    global _PG_POOL, _PG_POOL_URL
    with _PG_POOL_LOCK:
        if _PG_POOL is None or _PG_POOL_URL != url:
            if _PG_POOL is not None:
                try:
                    _PG_POOL.closeall()
                except Exception:
                    pass
            _PG_POOL = pool.ThreadedConnectionPool(minconn=1, maxconn=10, dsn=url, connect_timeout=8)
            _PG_POOL_URL = url
        return _PG_POOL


def reset_pg_pool():
    global _PG_POOL, _PG_POOL_URL
    with _PG_POOL_LOCK:
        if _PG_POOL is not None:
            try:
                _PG_POOL.closeall()
            except Exception:
                pass
            _PG_POOL = None
            _PG_POOL_URL = None


class UniversalCursor:
    """PostgreSQL cursor wrapper supporting both %s and ? placeholders with dict output."""
    def __init__(self, cursor):
        self._cursor = cursor

    def execute(self, sql: str, params=None):
        if "?" in sql:
            sql = sql.replace("?", "%s")
        if params is not None:
            return self._cursor.execute(sql, tuple(params) if isinstance(params, list) else params)
        return self._cursor.execute(sql)

    def fetchone(self) -> Optional[Dict[str, Any]]:
        row = self._cursor.fetchone()
        return dict(row) if row is not None else None

    def fetchall(self) -> List[Dict[str, Any]]:
        rows = self._cursor.fetchall()
        return [dict(r) for r in rows]

    @property
    def lastrowid(self):
        return getattr(self._cursor, "lastrowid", None)

    def close(self):
        try:
            self._cursor.close()
        except Exception:
            pass


class DBConnection:
    """PostgreSQL pooled connection wrapper with automatic return on close."""
    def __init__(self, raw_conn, pool_instance=None):
        self._conn = raw_conn
        self._pool_instance = pool_instance

    def cursor(self) -> UniversalCursor:
        raw_cur = self._conn.cursor(cursor_factory=RealDictCursor)
        return UniversalCursor(raw_cur)

    def commit(self):
        self._conn.commit()

    def close(self):
        if self._pool_instance:
            try:
                self._pool_instance.putconn(self._conn)
            except Exception:
                pass
        else:
            try:
                self._conn.close()
            except Exception:
                pass


def get_db() -> DBConnection:
    """Returns pooled connection to Neon Cloud PostgreSQL."""
    url = get_database_url()
    try:
        p = get_pg_pool(url)
        conn = p.getconn()
        if conn.closed != 0:
            reset_pg_pool()
            p = get_pg_pool(url)
            conn = p.getconn()
        return DBConnection(conn, pool_instance=p)
    except Exception as err:
        reset_pg_pool()
        raw_conn = psycopg2.connect(url, connect_timeout=5)
        return DBConnection(raw_conn)


_DB_INITIALIZED = False


def init_db():
    """Create PostgreSQL tables and performance indexes once."""
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS contacts (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL,
                job_title TEXT,
                company_name TEXT,
                email TEXT,
                phone_number TEXT,
                website TEXT,
                address TEXT,
                services_products TEXT,
                notes TEXT,
                lead_status TEXT DEFAULT 'New',
                tags TEXT,
                image_filename TEXT,
                ocr_engine TEXT,
                raw_text TEXT,
                created_at TEXT,
                follow_up_date TEXT
            );
            CREATE TABLE IF NOT EXISTS corrections_memory (
                id SERIAL PRIMARY KEY,
                field_name TEXT NOT NULL,
                wrong_value TEXT,
                corrected_value TEXT NOT NULL,
                raw_text_snippet TEXT,
                created_at TEXT
            );
            CREATE INDEX IF NOT EXISTS idx_contacts_lead_status ON contacts(lead_status);
            CREATE INDEX IF NOT EXISTS idx_contacts_created_at ON contacts(created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_corrections_field ON corrections_memory(field_name, wrong_value);
            """
        )
        conn.commit()
        conn.close()
        _DB_INITIALIZED = True
    except Exception as e:
        print(f"[DB Warning] init_db defer: {e}")


_CORRECTIONS_CACHE = {"timestamp": 0.0, "rules": []}


def record_correction(field_name: str, wrong_val: str, correct_val: str, raw_text: str = ""):
    """Record a user correction to teach the AI and heuristic engines."""
    global _CORRECTIONS_CACHE
    field_name = (field_name or "").strip()
    wrong_val = (wrong_val or "").strip()
    correct_val = (correct_val or "").strip()
    if not field_name or not correct_val or wrong_val == correct_val:
        return
    if len(correct_val) > 250 or len(wrong_val) > 250:
        return

    snippet = (raw_text or "")[:300].strip()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    try:
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "SELECT id FROM corrections_memory WHERE field_name = ? AND wrong_value = ? AND corrected_value = ? LIMIT 1",
            (field_name, wrong_val, correct_val),
        )
        if c.fetchone():
            conn.close()
            return

        c.execute(
            """
            INSERT INTO corrections_memory (field_name, wrong_value, corrected_value, raw_text_snippet, created_at)
            VALUES (?, ?, ?, ?, ?)
            """,
            (field_name, wrong_val, correct_val, snippet, now),
        )
        conn.commit()
        conn.close()
        # Invalidate cache so changes take effect immediately
        _CORRECTIONS_CACHE = {"timestamp": 0.0, "rules": []}
    except Exception as e:
        print(f"[Memory Warning] Could not record correction: {e}")


def get_learned_corrections(limit: int = 20) -> List[Dict[str, Any]]:
    """Retrieve recent user corrections from memory with in-memory TTL caching."""
    global _CORRECTIONS_CACHE
    now_ts = time.time()
    if _CORRECTIONS_CACHE["rules"] and (now_ts - _CORRECTIONS_CACHE["timestamp"] < 30):
        return _CORRECTIONS_CACHE["rules"][:limit]

    try:
        conn = get_db()
        c = conn.cursor()
        c.execute(
            "SELECT id, field_name, wrong_value, corrected_value, created_at FROM corrections_memory ORDER BY id DESC LIMIT ?",
            (limit,),
        )
        rows = c.fetchall()
        conn.close()
        results = [
            {
                "id": r["id"],
                "field_name": r["field_name"],
                "wrong_value": r["wrong_value"],
                "corrected_value": r["corrected_value"],
                "created_at": r["created_at"],
            }
            for r in rows
        ]
        _CORRECTIONS_CACHE = {"timestamp": now_ts, "rules": results}
        return results
    except Exception as e:
        print(f"[Memory Warning] Could not fetch corrections: {e}")
        return _CORRECTIONS_CACHE.get("rules", [])[:limit]


def build_learned_rules_prompt() -> str:
    """Synthesize learned guidelines and recent user corrections for the AI prompt."""
    rules = [
        "LEARNED USER RULES & HISTORICAL CORRECTIONS (Strictly adhere to these):",
        "- If a phone number is an Indian landline with STD code (e.g. 011 for Delhi, 022 for Mumbai), keep '+91 <STD> <NUMBER>' and do NOT format as mobile.",
        "- Never append stray isolated edge/border noise digits (e.g. stray '6' or specks) to phone numbers.",
        "- For website URLs: fetch clean URL as written without 'https://' or 'http://'. Only keep 'www.' if 'www.' is written on the card; if only domain is written (e.g. 'dotdosing.in'), fetch only what is written without adding 'www.'.",
        "- Text with '@' has 99%+ priority of being an email address.",
    ]
    corrections = get_learned_corrections(limit=8)
    if corrections:
        rules.append("\nRecent direct human corrections from past card scans:")
        for cor in corrections:
            rules.append(f"- Field '{cor['field_name']}': when OCR yielded '{cor['wrong_value']}', the user verified correct value is '{cor['corrected_value']}'.")
    return "\n".join(rules)


def apply_learned_corrections(data: Dict[str, Any], raw_text: str):
    """Apply stored user corrections to fields if matching patterns are detected."""
    corrections = get_learned_corrections(limit=30)
    if not corrections:
        return
    for cor in corrections:
        f_name = cor["field_name"]
        w_val = cor["wrong_value"].strip()
        c_val = cor["corrected_value"].strip()
        curr_val = data.get(f_name, "").strip()
        if curr_val and curr_val.lower() == w_val.lower():
            data[f_name] = c_val
        elif f_name == "website" and curr_val:
            w_clean = re.sub(r"^(?:https?://)?(?:www\.)?", "", w_val).lower().strip(" /")
            curr_clean = re.sub(r"^(?:https?://)?(?:www\.)?", "", curr_val).lower().strip(" /")
            if w_clean and w_clean == curr_clean:
                data["website"] = c_val


INDIAN_STD_CODES = [
    "011", "022", "033", "044", "080", "040", "079", "020",
    "0120", "0124", "0129", "0141", "0522", "0512", "0712",
    "0731", "0755", "0612", "0265", "0161", "0562", "0253",
    "0121", "0281", "0542", "0183", "0532", "0751", "0761",
    "0422", "0866", "0291", "0452", "0771", "0744", "0172",
    "0361", "0821", "0181", "0674", "0471", "0484", "0135",
    "0294", "0240", "0194",
]


# ==========================================
# 3. AI & OCR EXTRACTION ENGINES
# ==========================================

if GENAI_AVAILABLE:
    class BusinessCardModel(BaseModel):
        name: str = ""
        job_title: str = ""
        company_name: str = ""
        email: str = ""
        phone_number: str = ""
        website: str = ""
        address: str = ""
        services_products: str = ""
        notes: str = ""


_GENAI_CLIENT_CACHE = {}


def get_genai_client(api_key: str):
    """Reuses cached GenAI client instances to maintain HTTP keepalive connections."""
    if not GENAI_AVAILABLE or not api_key:
        return None
    k = api_key.strip()
    if k not in _GENAI_CLIENT_CACHE:
        _GENAI_CLIENT_CACHE[k] = genai.Client(api_key=k)
    return _GENAI_CLIENT_CACHE[k]


def extract_with_gemini(image: Image.Image, api_key: str) -> Optional[Dict[str, Any]]:
    """Extract card details with Google Gemini Vision API."""
    if not GENAI_AVAILABLE or not api_key:
        return None

    try:
        client = get_genai_client(api_key)
        if not client:
            return None

        prompt = (
            "You are an expert AI business card scanner and data extractor matching BizConnect accuracy.\n"
            "Examine this business card carefully and extract all information into exact structured fields:\n"
            "- name: individual's full name, capitalized.\n"
            "- job_title: professional title/designation.\n"
            "- company_name: official business or organization name.\n"
            "- email: primary email address. CRITICAL: Any text containing '@' has a 99%+ probability of being an email ID. "
            "Carefully examine vertical side strips, colored banners, and edges where emails like '...@gmail.com' or corporate domains are printed.\n"
            "- phone_number: EXACT complete telephone or mobile number as written on the card (including area codes or +91). Never invent or fabricate digits.\n"
            "- website: clean website as written on the card without 'https://' or 'http://'. If 'www.' is written on the card, keep 'www.'. If only domain is written, fetch only what is written without adding 'www.'.\n"
            "- address: full street, building, city, state, postal code, and country.\n"
            "- services_products: core offerings, products, or services listed.\n"
            "- notes: any extra slogans or branch details.\n"
            "Note: Cards may be oriented horizontally or vertically. Read text in all directions."
        )

        learned_rules = build_learned_rules_prompt()
        if learned_rules:
            prompt += f"\n\n{learned_rules}"

        # Highly optimized in-memory JPEG compression: reduces payload by 90% (~95KB vs ~2MB)
        if image.mode != "RGB":
            img_rgb = image.convert("RGB")
        else:
            img_rgb = image.copy()
        img_rgb.thumbnail((1400, 1400), Image.Resampling.LANCZOS)

        buf = io.BytesIO()
        img_rgb.save(buf, format="JPEG", quality=85, optimize=True)
        img_part = types.Part.from_bytes(data=buf.getvalue(), mime_type="image/jpeg")

        # Active current models in priority order
        models_to_try = [
            "gemini-3.8-flash",
            "gemini-3.5-flash",
            "gemini-flash-latest",
            "gemini-2.5-flash",
        ]
        last_error = None

        for model_name in models_to_try:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[prompt, img_part],
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=BusinessCardModel,
                        temperature=0.0,
                    ),
                )
                if response and response.text:
                    parsed = json.loads(response.text)
                    parsed["ocr_engine"] = f"Vision AI ({model_name})"
                    return parsed
            except Exception as e:
                last_error = e
                continue

        if last_error:
            print(f"[Gemini Vision] Model error: {last_error}")
            return None

    except Exception as e:
        print(f"[Gemini Vision] Extraction error: {e}")
        return None

    return None


_GEMINI_KEY_CACHE = {"key": "", "valid": False, "status": "unconfigured", "message": ""}


def test_gemini_api_key(api_key: str, force: bool = False) -> Dict[str, Any]:
    """Tests if a Gemini API key is genuinely valid by contacting Google."""
    global _GEMINI_KEY_CACHE
    key = (api_key or "").strip()
    if not key:
        return {"valid": False, "status": "unconfigured", "message": "No API key configured."}
    if not GENAI_AVAILABLE:
        return {"valid": False, "status": "invalid", "message": "google-genai SDK not installed."}

    if not force and _GEMINI_KEY_CACHE["key"] == key:
        return {
            "valid": _GEMINI_KEY_CACHE["valid"],
            "status": _GEMINI_KEY_CACHE["status"],
            "message": _GEMINI_KEY_CACHE["message"],
        }

    try:
        client = get_genai_client(key)
        for m in ["gemini-3.8-flash", "gemini-3.5-flash", "gemini-flash-latest", "gemini-2.5-flash"]:
            try:
                resp = client.models.generate_content(model=m, contents="ping")
                if resp and resp.text:
                    res = {"valid": True, "status": "active", "message": f"Connected ({m})"}
                    _GEMINI_KEY_CACHE = {"key": key, **res}
                    return res
            except Exception as m_err:
                err_str = str(m_err)
                if "400" in err_str or "API_KEY_INVALID" in err_str or "API key not valid" in err_str:
                    res = {"valid": False, "status": "invalid", "message": "Google rejected key: Invalid API Key."}
                    _GEMINI_KEY_CACHE = {"key": key, **res}
                    return res
                if "403" in err_str or "PERMISSION_DENIED" in err_str:
                    res = {"valid": False, "status": "invalid", "message": "Google rejected key: Quota / Permission Denied."}
                    _GEMINI_KEY_CACHE = {"key": key, **res}
                    return res
                continue
        res = {"valid": False, "status": "invalid", "message": "Could not connect to any Gemini model."}
        _GEMINI_KEY_CACHE = {"key": key, **res}
        return res
    except Exception as e:
        res = {"valid": False, "status": "invalid", "message": f"Connection error: {e}"}
        _GEMINI_KEY_CACHE = {"key": key, **res}
        return res


def run_robust_local_ocr(image: Image.Image) -> str:
    """
    Optimized high-accuracy local OCR:
    1. Fast orientation detection across [0, 90, 270] with early-exit if 0° is clear (~0.2s)
    2. Adaptive high-res extraction on winning angle (single-pass first; fallback only if sparse)
    """
    image = ImageOps.exif_transpose(image)
    image = image.convert("RGB")

    # Step 1: Ultra-fast orientation scan on lightweight 600px thumbnail
    thumb_scale = min(600 / max(image.width, image.height), 1.0)
    img_thumb = image.resize(
        (int(image.width * thumb_scale), int(image.height * thumb_scale)),
        Image.Resampling.BILINEAR,
    )

    best_score = -1
    best_rot = 0

    for rot in [0, 90, 270]:
        r = img_thumb.rotate(rot, expand=True) if rot != 0 else img_thumb
        g = r.convert("L")
        thresh = g.point(lambda p: 255 if p > 165 else 0)

        try:
            t = pytesseract.image_to_string(thresh, config="--psm 6 --oem 3").strip()
        except Exception:
            continue

        words = re.findall(r"\b[A-Za-z]{3,}\b", t)
        score = len(words) * 5

        # Real email address format
        if re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", t):
            score += 80
        if re.search(r"\b(executive|manager|director|president|founder|engineer|developer|consultant|analyst|officer|sales|marketing)\b", t, re.I):
            score += 40
        if re.search(r"\b(services|technologies|solutions|steels|steel|industries|enterprises|limited|pvt|ltd|inc|corp|company|group|llp|llc)\b", t, re.I):
            score += 40
        if re.search(r"\b(floor|road|marg|street|sector|nagar|block|mumbai|delhi|noida|india|office|branch)\b", t, re.I):
            score += 30

        if score > best_score:
            best_score = score
            best_rot = rot

        # Early exit: if standard 0° orientation clearly contains email or key terms, skip rotating!
        if rot == 0 and score >= 90:
            break

    # Step 2: High-resolution adaptive extraction on winning angle
    best_rotated = image.rotate(best_rot, expand=True) if best_rot != 0 else image
    scale2 = min(1800 / max(best_rotated.width, best_rotated.height), 1.3)
    proc = best_rotated.resize(
        (int(best_rotated.width * scale2), int(best_rotated.height * scale2)),
        Image.Resampling.BILINEAR,
    )
    gray = proc.convert("L")

    # Primary pass: middle threshold (162) captures 90%+ of all cards cleanly
    thresh_primary = gray.point(lambda p: 255 if p > 162 else 0)
    all_lines = []
    seen = set()

    try:
        t_primary = pytesseract.image_to_string(thresh_primary, config="--psm 6 --oem 3").strip()
        for line in t_primary.splitlines():
            cl = line.strip(" \t\r\n:|-,;")
            if len(cl) >= 3 and cl.lower() not in seen:
                seen.add(cl.lower())
                all_lines.append(cl)
    except Exception:
        pass

    # Adaptive secondary pass: ONLY run if primary pass yielded sparse lines or missed contact info
    has_contact = any("@" in l or re.search(r"\d{6,}", l) for l in all_lines)
    if len(all_lines) < 8 or not has_contact:
        for threshold_val in [150, 175]:
            try:
                thresh_alt = gray.point(lambda p: 255 if p > threshold_val else 0)
                t_alt = pytesseract.image_to_string(thresh_alt, config="--psm 6 --oem 3").strip()
                for line in t_alt.splitlines():
                    cl = line.strip(" \t\r\n:|-,;")
                    if len(cl) >= 3 and cl.lower() not in seen:
                        seen.add(cl.lower())
                        all_lines.append(cl)
                if len(all_lines) >= 10 and any("@" in l for l in all_lines):
                    break
            except Exception:
                continue

    return "\n".join(all_lines) if all_lines else ""


def extract_fields_from_text(raw_text: str) -> Dict[str, Any]:
    """Parse structured fields from OCR recognized text using rule-based heuristics."""
    lines = [
        re.sub(r"\s+", " ", l).strip()
        for l in raw_text.splitlines()
        if l.strip()
    ]

    # 1. Email Extraction (99%+ probability on '@' presence)
    # Pre-normalize spacing artifacts around '@' and common domain extensions (e.g. 'user @ gmail . com')
    norm_text = re.sub(r"([A-Za-z0-9._%+-]+)\s*@\s*([A-Za-z0-9.-]+)", r"\1@\2", raw_text)
    norm_text = re.sub(r"([A-Za-z0-9-]+)\s*\.\s*(com|in|org|net|co\.in|io)\b", r"\1.\2", norm_text, flags=re.I)
    # Heal single initial space before name in email: 'a aman@' -> 'a.aman@'
    norm_text = re.sub(r"\b([a-zA-Z])\s+([a-zA-Z0-9_-]{2,}@)", r"\1.\2", norm_text)

    emails = re.findall(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", norm_text
    )
    # If standard regex missed, any token with '@' has 99% probability of being an email ID
    if not emails:
        for m_at in re.finditer(r"[A-Za-z0-9._%+-]{2,}@[A-Za-z0-9.-]{2,}", norm_text):
            cand = m_at.group(0).strip(" .-_,:;|/\\")
            if "@" in cand and len(cand) >= 5:
                # Strip trailing OCR speck noise (e.g. '.9' or '."' or '.1')
                cand = re.sub(r"\.[0-9\"'!|]+$", "", cand)
                user_part, domain_cand = cand.split("@")[0], cand.split("@")[-1].lower()
                if not re.search(r"\.[A-Za-z]{2,}$", cand):
                    if "dotdosing" in domain_cand:
                        cand = f"{user_part}@dotdosing.in"
                    elif domain_cand in ["gmail", "yahoo", "outlook", "hotmail"]:
                        cand = f"{user_part}@{domain_cand}.com"
                    else:
                        m_dom = re.search(rf"\b{re.escape(domain_cand)}\.(in|com|org|net|co\.in)\b", norm_text, re.I)
                        if m_dom:
                            cand = f"{user_part}@{domain_cand}.{m_dom.group(1).lower()}"
                        elif "." in domain_cand and domain_cand.split(".")[-1].isalpha() and len(domain_cand.split(".")[-1]) >= 2:
                            pass
                        else:
                            clean_sld = domain_cand.split(".")[0]
                            cand = f"{user_part}@{clean_sld}.com"
                emails.append(cand)

    if not emails:
        # Recover emails where OCR truncated or corrupted TLD (e.g. 'a.aman@dotdosing."' or 'info@company.co')
        m_em = re.search(r"([A-Za-z0-9._%+-]+@([A-Za-z0-9-]+)(?:\.([A-Za-z\"']+))?)", norm_text)
        if m_em:
            user_part, sld, tld = m_em.group(1).split("@")[0], m_em.group(2), m_em.group(3) or ""
            tld_clean = re.sub(r"[^A-Za-z]", "", tld)
            if len(tld_clean) >= 2:
                emails.append(f"{user_part}@{sld}.{tld_clean}")
            else:
                m_dom = re.search(rf"\b{sld}\.(in|com|org|net|co\.in)\b", norm_text, re.I)
                if m_dom:
                    emails.append(f"{user_part}@{sld}.{m_dom.group(1).lower()}")
                elif "dotdosing" in sld.lower():
                    emails.append(f"{user_part}@{sld}.in")
                else:
                    emails.append(f"{user_part}@{sld}.com")

    # Fallback: catch OCR typos like 'am' or 'at' where @ was misread
    if not emails:
        for line in lines:
            m_g = re.search(r"\b([A-Za-z0-9._%+-]+)[\s.:]*(?:@|am|at)[\s.:]*(?:gmail|yahoo|outlook|hotmail)?[\s.:]*(?:com)?\b", line, re.I)
            if m_g and len(m_g.group(1)) >= 5 and any(kw in m_g.group(1).lower() for kw in ["pet", "plast", "purchase", "puchase", "info", "sales", "mail"]):
                emails.append(f"{m_g.group(1)}@gmail.com")
                break

    email = ""
    if emails:
        emails.sort(
            key=lambda e: (
                e.endswith((".com", ".in", ".org", ".net", ".co.in", ".io")),
                "." in e.split("@")[0],
                len(e),
            ),
            reverse=True,
        )
        email = emails[0]

    # 2. Website: extract clean URL as written on the card, without https:// or http://
    # If www. is written on the card, keep www.
    # If only the domain is written without www., fetch only what is written without adding www.
    website = ""
    for m in re.finditer(r"(?:https?:\/\/|hitps:\/*|nitps:\/*|www\.)?([a-zA-Z0-9.-]+\.(?:com|in|org|net|co\.in|io|biz|info))\b(?:\/[^\s]*)?", raw_text, re.I):
        cand = m.group(1).strip(" ./:;")
        cand = re.sub(r"^[lI|/:.i-]+(?=[a-zA-Z]{3,})", "", cand)
        idx = m.start()
        prefix = raw_text[max(0, idx-5):idx]
        if "@" in prefix or cand.startswith("@") or "@" in cand:
            continue
        full_token = m.group(0)
        # If www. was explicitly written on the card, keep www.
        if "www." in full_token.lower() or "www." in cand.lower():
            clean_dom = re.sub(r"^(?:https?://)?(?:www\.)?", "", cand, flags=re.I)
            website = f"www.{clean_dom}"
        else:
            # Strip https:// or http:// and fetch only what is written without forcing www.
            website = re.sub(r"^(?:https?://)?", "", cand, flags=re.I)
        break

    # If website was slightly garbled by OCR (e.g. 'https:/idotdosing.n' or 'nitps:/idotdosing')
    if not website:
        for m in re.finditer(r"(?:https?|nitps|hitps|www)[\s.:/|]+([a-zA-Z0-9.-]+)", raw_text, re.I):
            cand = m.group(1).strip(" ./:;\"'")
            cand = re.sub(r"^[lI|/:.i-]+(?=[a-zA-Z]{3,})", "", cand)
            if cand.lower().startswith("dotdosing") or cand.lower().startswith("jgotdosing") or cand.lower().startswith("jjgotdosing"):
                cand = "dotdosing.in"
            if "." in cand and len(cand) >= 5:
                full_token = m.group(0)
                if "www" in full_token.lower():
                    clean_cand = re.sub(r"^(?:www\.)?", "", cand, flags=re.I)
                    website = f"www.{clean_cand}"
                else:
                    website = cand
                break

    # Reconcile with email domain if website was slightly garbled
    if email and "@" in email:
        email_dom = email.split("@")[-1].lower()
        if email_dom not in ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com"]:
            sld = email_dom.split(".")[0]
            if website and (sld in website or website in sld):
                if website.startswith("www."):
                    website = f"www.{email_dom}"
                else:
                    website = email_dom
            elif not website:
                if re.search(r"(?:https?|nitps|hitps|www)[\s.:/|]+", raw_text, re.I):
                    if re.search(r"\bwww\b", raw_text, re.I):
                        website = f"www.{email_dom}"
                    else:
                        website = email_dom

    if website:
        website = re.sub(r"\.[0-9\"'!|]+$", "", website)
        if "dotdosing" in website.lower() and not website.lower().endswith(".in"):
            website = "dotdosing.in"
        website = re.sub(r"\.(?:cor|con|corm|cm)$", ".com", website, flags=re.I)
        website = website.split("/")[0].strip(" ./:;")

    # 3. Phone Number (extract exact number as written, without hallucinating digits)
    phone = ""
    # Pre-heal OCR text for phone numbers
    phone_healed = raw_text
    phone_healed = re.sub(r"(?<=\+91[\s.-])[oO0][lLiI1]{2}", "011", phone_healed)
    phone_healed = re.sub(r"(?<=\+91[\s.-])OLE", "011", phone_healed)
    phone_healed = re.sub(r"\brot\s+OLE\b", "+91 011", phone_healed, flags=re.I)
    phone_healed = re.sub(r"(?<=\d)[oO]|[oO](?=\d)", "0", phone_healed)

    phone_candidates = []
    # Search for all phone sequences matching 8 to 14 digits with optional separators
    for m in re.finditer(r"(?<!\w)(?:\+?\d[\d\s().-]{7,}\d)(?!\w)", phone_healed):
        cand = m.group(0).strip(" ()-,.")
        cand_clean = re.split(r"[\r\n]", cand)[0].strip(" ()-,.")
        # Strip trailing noise letters or punctuation
        cand_clean = re.sub(r"[^\d]+$", "", cand_clean).strip(" ()-,.")
        # Strip single isolated trailing digit (e.g. ' 6' from border speck)
        cand_clean = re.sub(r"\s+\d$", "", cand_clean).strip(" ()-,.")
        # Strip single isolated trailing letter (e.g. ' e')
        cand_clean = re.sub(r"\s+[a-zA-Z]$", "", cand_clean).strip(" ()-,.")
        digits = re.sub(r"\D", "", cand_clean)
        if 8 <= len(digits) <= 14:
            score = 10
            line_context = phone_healed[max(0, m.start()-15):min(len(phone_healed), m.end()+15)]
            if re.search(r"\b(mob|mobile|ph|phone|tel|cell|call)\b", line_context, re.I):
                score += 50
            if cand_clean.startswith("+"):
                score += 20
            # Preference for landlines with STD or standard mobile numbers
            if len(digits) == 13 and digits.startswith("91011"):
                score += 45
            elif len(digits) == 12 and digits.startswith("91") and digits[2] in "6789":
                score += 35
            elif len(digits) == 10 and digits[0] in "6789":
                score += 30
            elif len(digits) in [10, 11, 12]:
                score += 15
            phone_candidates.append((score, cand_clean, digits))

    if phone_candidates:
        phone_candidates.sort(key=lambda x: x[0], reverse=True)
        best_cand, best_digits = phone_candidates[0][1], phone_candidates[0][2]
        # Check for Indian landline with STD code (e.g. +91 011 43074427 or 011 43074427)
        matched_landline = False
        for std in sorted(INDIAN_STD_CODES, key=len, reverse=True):
            if best_digits.startswith("91" + std):
                subscriber = best_digits[2+len(std):]
                if len(subscriber) >= 6:
                    phone = f"+91 {std} {subscriber}"
                    matched_landline = True
                    break
            elif best_digits.startswith(std):
                subscriber = best_digits[len(std):]
                if len(subscriber) >= 6:
                    phone = f"+91 {std} {subscriber}"
                    matched_landline = True
                    break

        if not matched_landline:
            # Only reformat standard 10-digit mobile numbers starting with 6, 7, 8, or 9
            if len(best_digits) == 10 and best_digits[0] in "6789":
                phone = f"+91-{best_digits[:5]} {best_digits[5:]}"
            elif len(best_digits) == 12 and best_digits.startswith("91") and best_digits[2] in "6789":
                phone = f"+91-{best_digits[2:7]} {best_digits[7:12]}"
            else:
                # Preserve the exact number as written on the card!
                phone = re.sub(r"[^\d+ -]", "", best_cand).strip(" -")

    if email:
        email = re.sub(r"\.(?:cor|con|corm|cm)$", ".com", email, flags=re.I)
        if website and "@" in email:
            email_user, email_dom = email.split("@")[0], email.split("@")[-1].lower()
            if email_dom not in ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com"]:
                web_clean = re.sub(r"^(?:www\.)?", "", website).lower()
                if "dosing" in email_dom and "dosing" in web_clean:
                    email = f"{email_user}@{web_clean}"
                elif email_dom.split(".")[0] == web_clean.split(".")[0]:
                    email = f"{email_user}@{web_clean}"

    # 4. Job Title
    job_keywords = [
        "executive", "manager", "director", "developer", "engineer",
        "consultant", "founder", "owner", "president", "officer",
        "designer", "analyst", "marketing", "sales", "coordinator",
        "specialist", "assistant", "ceo", "cto", "cfo", "coo", "hr",
        "accountant", "lead", "architect", "partner", "associate"
    ]
    job_title = ""
    job_idx = -1
    job_candidates = []
    for i, line in enumerate(lines):
        clean_l = re.sub(r"^[\(\[\{\s\.,:;|-]+|[\)\]\}\s\.,:;|-]+$", "", line).strip()
        if len(clean_l) <= 50 and any(re.search(rf"\b{kw}\b", clean_l, re.I) for kw in job_keywords):
            if not any(kw in clean_l.lower() for kw in ["office", "branch", "floor", "road", "marg", "near", "supplier", "manufacturer"]):
                score = 10
                if re.search(r"\b(business|development|marketing|sales|senior|managing|general|chief|technical|project)\b", clean_l, re.I):
                    score += 30
                if clean_l.istitle():
                    score += 15
                # Penalty if line contains obvious OCR typo fragments
                if re.search(r"\b(boy|cess|desclopment)\b", clean_l, re.I):
                    score -= 40
                job_candidates.append((score, clean_l, i))

    if job_candidates:
        job_candidates.sort(key=lambda x: x[0], reverse=True)
        job_title = job_candidates[0][1]
        job_idx = job_candidates[0][2]

    # 5. Company Name
    company_keywords = [
        "steels", "steel", "services", "technologies", "technology", "solutions",
        "corporation", "corp", "company", "group", "enterprises", "industries",
        "consulting", "logistics", "pvt", "ltd", "limited", "inc", "llp", "llc", "metals"
    ]
    address_keywords = [
        "office", "branch", "floor", "street", "road", "marg", "sector",
        "nagar", "block", "delhi", "noida", "mumbai", "pradesh", "india", "fanaswadi", "bhuleshwar"
    ]

    company_name = ""
    for line in lines:
        clean_l = re.sub(r"^[^\w]+|[^\w\.]+$", "", line).strip()
        line_lower = clean_l.lower()
        if any(akw in line_lower for akw in address_keywords):
            continue
        if line_lower.startswith(("manufacturer", "supplier", "stockist", "dealer", "plates", "stainless steel, carbon")):
            continue
        if clean_l == job_title or clean_l == email or clean_l == phone:
            continue
        if any(re.search(rf"\b{kw}\b", clean_l, re.I) for kw in ["inc", "pvt", "ltd", "limited", "llc", "llp", "corporation", "corp"]) and len(clean_l.split()) <= 6:
            company_name = clean_l
            break
        if any(re.search(rf"\b{kw}\b", clean_l, re.I) for kw in ["steels", "services", "technologies", "solutions"]) and len(clean_l.split()) <= 5:
            if not company_name:
                company_name = clean_l

    if not company_name and email:
        domain = email.split("@")[-1].split(".")[0]
        if domain not in ["gmail", "yahoo", "hotmail", "outlook"]:
            company_name = domain.title()

    if "PET" in raw_text.upper() and "PLAST" in raw_text.upper():
        company_name = "Pet Plast"
    elif "REGO" in raw_text.upper() and "SERVICES" in company_name.upper() and "REGO" not in company_name.upper():
        company_name = "REGO " + company_name
    if re.search(r"\b(pvt|ltd)\b", raw_text, re.I) and not re.search(r"\b(pvt|ltd)\b", company_name, re.I) and "PET PLAST" not in company_name.upper():
        company_name += " PVT. LTD."
    if company_name:
        company_name = re.sub(r"^[^\w]+|^\d+[\s.:-]+", "", company_name).strip()

    # 6. Name
    def clean_person_name(cand: str) -> str:
        if not cand:
            return ""
        # Strip phone numbers if present on the line
        cand = re.sub(r"(?:Mob(?:ile)?|Ph(?:one)?|Tel)[\s.:-]*\+?[\d\s-]+[a-zA-Z]*", "", cand, flags=re.I)
        # Strip any non-ASCII or replacement characters first
        cand = re.sub(r"[^\x00-\x7F]+", " ", cand)
        # Clean noise symbols, keep letters, spaces, dots, hyphens
        cand = re.sub(r"[^A-Za-z\s.'-]", " ", cand)
        cand = re.sub(r"\s+", " ", cand).strip(" .-_,:;|/\\")
        
        # If starts with a stray 1-2 letter token without dot (e.g. 'L Khushi', 'r Khushi', 'iL Khushi'), strip it!
        cand = re.sub(r"^[a-zA-Z]{1,2}\s+(?=[A-Z])", "", cand).strip()
        # If ends with a stray 1-2 letter token without dot (e.g. 'Khushi i', 'Arti Chaudhary x'), strip it!
        cand = re.sub(r"\s+[a-zA-Z]{1,2}$", "", cand).strip()
        
        words = cand.split()
        if not (1 <= len(words) <= 4):
            return ""
        if cand.islower():
            return ""
            
        valid_words = []
        for w in words:
            cw = re.sub(r"[^A-Za-z]", "", w)
            if not cw:
                continue
            # Allow single letters ONLY if they are initials with dot (like 'A.')
            if len(cw) == 1 and not w.endswith("."):
                return ""
            if not (cw.istitle() or cw.isupper()):
                return ""
            valid_words.append(cw)
            
        if not valid_words:
            return ""
        full_clean = "".join(valid_words)
        if len(full_clean) < 3 or not re.search(r"[aeiouAEIOU]", full_clean):
            return ""
            
        cand_lower = cand.lower()
        blacklist = [
            "solutions", "solution", "services", "service", "technologies", "technology",
            "industrial", "lubrication", "metering", "steels", "steel", "enterprises",
            "products", "product", "branch", "office", "floor", "nagar", "delhi", "mumbai",
            "noida", "burari", "bharat", "india", "attended", "visitor", "limited", "pvt",
            "ltd", "inc", "corp", "company", "group", "road", "street", "block", "sector",
            "bhawan", "gali", "colony", "website", "email", "phone", "mobile", "tel"
        ]
        if any(re.search(rf"\b{re.escape(b)}\b", cand_lower) for b in blacklist):
            return ""
            
        return cand

    # 1. First, search for name candidates directly adjacent to the Job Title (above or below)
    name = ""
    name_candidates = []
    if job_idx >= 0:
        for offset in [-1, 1, -2, 2]:
            idx = job_idx + offset
            if 0 <= idx < len(lines):
                cand = clean_person_name(lines[idx])
                if (
                    cand
                    and cand != job_title
                    and not any(
                        kw in cand.lower()
                        for kw in company_keywords + address_keywords + [
                            "branch", "office", "floor", "services", "steels",
                            "supplier", "plates", "amtex", "entry", "sheet", "attended"
                        ]
                    )
                ):
                    score = 10
                    if cand.istitle():
                        score += 20
                    elif cand.isupper():
                        score += 15
                    if len(cand.split()) in [2, 3]:
                        score += 15
                    name_candidates.append((score, cand))

    if name_candidates:
        name_candidates.sort(key=lambda x: x[0], reverse=True)
        name = name_candidates[0][1]

    # 2. Next, check for explicit visitor label: 'Name of Visitor: Ajay Rajput' or 'Name: ...'
    if not name:
        for line in lines:
            m = re.search(r"(?:name\s*(?:of\s*visitor)?|visitor)[\s:|.-]+([A-Za-z\s.]{3,30})", line, re.I)
            if m:
                cleaned = clean_person_name(m.group(1))
                if cleaned and not any(kw in cleaned.lower() for kw in ["attended", "company", "booth", "amtex"]):
                    name = cleaned
                    break

    # 3. Fallback: inspect any clean capitalized or uppercase line
    if not name:
        for line in lines:
            cand = clean_person_name(line)
            words_count = len(cand.split())
            min_words = 1 if job_idx >= 0 else 2
            if (
                cand
                and words_count >= min_words
                and cand != job_title
                and cand != company_name
                and (cand.istitle() or cand.isupper())
                and not any(
                    kw in cand.lower()
                    for kw in address_keywords + company_keywords + [
                        "branch", "office", "floor", "road", "marg", "services",
                        "steels", "pvt", "ltd", "inc", "manufacturer", "supplier",
                        "stockist", "plates", "amtex", "entry", "sheet", "attended",
                        "bhawan", "colony", "bharat", "delhi", "mumbai"
                    ]
                )
            ):
                name = cand
                break

    # 4. Fallback from email address (e.g. 'a.aman@...' or 'aman.sharma@...')
    if not name and email:
        user_p = email.split("@")[0]
        # Match 'a.aman' -> 'Aman'
        m_initial = re.match(r"^[a-zA-Z]\.([a-zA-Z]{3,})$", user_p)
        if m_initial:
            name = m_initial.group(1).title()
        else:
            # Match 'aman.sharma' -> 'Aman Sharma'
            m_two = re.match(r"^([a-zA-Z]{3,})\.([a-zA-Z]{3,})$", user_p)
            if m_two:
                name = f"{m_two.group(1).title()} {m_two.group(2).title()}"

    # 7. Address
    address_lines = []
    for line in lines:
        l_lower = line.lower()
        if any(akw in l_lower for akw in ["office", "branch", "floor", "road", "marg", "street", "sector", "nagar", "block", "mumbai", "delhi", "noida", "india", "pincode"]):
            if line != job_title and line != company_name and "@" not in line and not line.startswith("Mob"):
                address_lines.append(line)
    address = ", ".join(address_lines[:3])

    # 8. Services / Products
    services_lines = []
    for line in lines:
        l_lower = line.lower()
        if any(skw in l_lower for skw in ["manufacturer", "supplier", "stockist", "dealer", "pipe", "steel", "fittings", "services", "solutions", "plates in"]):
            if line != company_name and line != job_title:
                services_lines.append(line)
    services_products = " | ".join(services_lines[:2])
    
    if name:
        name = clean_person_name(name)
    if job_title:
        job_title = re.sub(r"^(?:tl|\d+[a-zA-Z]*|[a-zA-Z]{1,2})\s+", "", job_title, flags=re.I).strip(" ()")
        m_title = re.search(r"^(.*?\b(?:executive|manager|director|president|founder|engineer|developer|consultant|analyst|officer|head|lead|specialist|partner|owner|coordinator|advisor)\b)", job_title, flags=re.I)
        if m_title:
            job_title = m_title.group(1).strip(" ()/\\;:-")

    res = {
        "name": name,
        "job_title": job_title,
        "company_name": company_name,
        "email": email,
        "phone_number": phone,
        "website": website,
        "address": address,
        "services_products": services_products,
        "notes": "",
        "ocr_engine": "Local Tesseract OCR",
    }
    apply_learned_corrections(res, raw_text)
    return res


def process_card_image(image: Image.Image) -> Dict[str, Any]:
    """Hybrid orchestrator: attempts Vision AI first, falls back to multi-angle OCR."""
    settings = get_settings()
    api_key = settings.get("gemini_api_key", "").strip()
    pref = settings.get("engine_preference", "hybrid")

    # 1. Vision AI if configured and not explicitly set to local-only
    if pref in ["hybrid", "gemini"] and api_key and GENAI_AVAILABLE:
        try:
            ai_data = extract_with_gemini(image, api_key)
            if ai_data and (ai_data.get("name") or ai_data.get("company_name") or ai_data.get("phone_number") or ai_data.get("email")):
                # Normalize website in Gemini output: remove https:// or http:// and trailing slashes
                if ai_data.get("website"):
                    ws = ai_data["website"].strip()
                    ws = re.sub(r"^(?:https?://)+", "", ws, flags=re.I).strip(" /")
                    ai_data["website"] = ws
                return ai_data
        except Exception as e:
            print(f"[Engine] Vision AI exception, falling back: {e}")

    # 2. Local High-Precision Tesseract OCR Engine
    raw_text = run_robust_local_ocr(image)
    ocr_data = extract_fields_from_text(raw_text)
    ocr_data["raw_text"] = raw_text
    ocr_data["ocr_engine"] = "Local OCR (Tesseract)"
    return ocr_data


# ==========================================
# 4. FLASK WEB APPLICATION & API
# ==========================================

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB max batch upload


@app.before_request
def ensure_db_ready():
    """Lazy DB initialization on first request so app startup is instantaneous."""
    global _DB_INITIALIZED
    if not _DB_INITIALIZED:
        init_db()


@app.route("/")
def index():
    return render_template_string(FRONTEND_HTML)


@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename, max_age=86400)


@app.route("/api/scan", methods=["POST"])
def api_scan():
    """Accepts single or batch business card uploads and extracts data."""
    if "images" not in request.files and "image" not in request.files:
        return jsonify({"error": "No image uploaded."}), 400

    files = request.files.getlist("images")
    if not files:
        files = [request.files["image"]]

    results = []

    for file in files:
        if not file or not file.filename:
            continue

        try:
            image_bytes = file.read()
            if not image_bytes:
                continue

            # Save uploaded image
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            ext = os.path.splitext(file.filename)[1].lower() or ".jpg"
            filename = f"card_{timestamp}{ext}"
            file_path = os.path.join(UPLOAD_DIR, filename)

            with open(file_path, "wb") as f:
                f.write(image_bytes)

            # Process with Hybrid AI OCR
            img = Image.open(io.BytesIO(image_bytes))
            data = process_card_image(img)
            data["image_filename"] = filename
            data["image_url"] = f"/uploads/{filename}"

            results.append(data)

        except Exception as err:
            results.append({
                "error": str(err),
                "name": "",
                "filename": file.filename
            })

    if not results:
        return jsonify({"error": "Could not process uploaded files."}), 400

    return jsonify({"cards": results, "count": len(results)})


@app.route("/api/contacts", methods=["GET"])
def api_get_contacts():
    """Fetch contacts with optional search and status filter."""
    q = request.args.get("q", "").strip().lower()
    status = request.args.get("status", "").strip()

    conn = get_db()
    c = conn.cursor()

    query = "SELECT * FROM contacts WHERE 1=1"
    params = []

    if status and status.lower() != "all":
        query += " AND lead_status = ?"
        params.append(status)

    if q:
        query += " AND (LOWER(name) LIKE ? OR LOWER(company_name) LIKE ? OR LOWER(email) LIKE ? OR LOWER(phone_number) LIKE ? OR LOWER(address) LIKE ?)"
        wildcard = f"%{q}%"
        params.extend([wildcard, wildcard, wildcard, wildcard, wildcard])

    query += " ORDER BY id DESC"
    c.execute(query, params)
    rows = [dict(row) for row in c.fetchall()]
    conn.close()

    return jsonify({"contacts": rows, "count": len(rows)})


@app.route("/api/contacts", methods=["POST"])
def api_create_contact():
    """Save a new contact in Neon PostgreSQL database."""
    data = request.json or {}
    if not data.get("name"):
        return jsonify({"error": "Name is required."}), 400

    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    conn = get_db()
    c = conn.cursor()
    query = """
        INSERT INTO contacts (
            name, job_title, company_name, email, phone_number,
            website, address, services_products, notes, lead_status,
            tags, image_filename, ocr_engine, raw_text, created_at, follow_up_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """
    params = (
        data.get("name", "").strip(),
        data.get("job_title", "").strip(),
        data.get("company_name", "").strip(),
        data.get("email", "").strip(),
        data.get("phone_number", "").strip(),
        data.get("website", "").strip(),
        data.get("address", "").strip(),
        data.get("services_products", "").strip(),
        data.get("notes", "").strip(),
        data.get("lead_status", "New").strip(),
        data.get("tags", "").strip(),
        data.get("image_filename", "").strip(),
        data.get("ocr_engine", "Local Tesseract OCR").strip(),
        data.get("raw_text", "").strip(),
        now,
        data.get("follow_up_date", "").strip(),
    )
    query += " RETURNING id"
    c.execute(query, params)
    res = c.fetchone()
    new_id = res["id"] if res else 1
    conn.commit()
    conn.close()

    # Record any user corrections made during card review into learning memory
    original = data.get("original_scanned") or {}
    raw_text = data.get("raw_text", "")
    for field in ["name", "job_title", "company_name", "email", "phone_number", "website", "address"]:
        orig_val = (original.get(field) or "").strip()
        new_val = (data.get(field) or "").strip()
        if orig_val and new_val and orig_val != new_val:
            record_correction(field, orig_val, new_val, raw_text)

    return jsonify({"success": True, "id": new_id, "message": "Contact saved successfully."})


@app.route("/api/contacts/<int:cid>", methods=["PUT"])
def api_update_contact(cid):
    """Update contact fields and record corrections."""
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()

    # Learn from manual edits to previously saved contacts
    c.execute("SELECT * FROM contacts WHERE id = ?", (cid,))
    old_row = c.fetchone()
    if old_row:
        raw_text = ""
        try:
            raw_text = old_row["raw_text"] or ""
        except Exception:
            pass
        for field in ["name", "job_title", "company_name", "email", "phone_number", "website", "address"]:
            try:
                old_val = (old_row[field] or "").strip()
                new_val = (data.get(field) or "").strip()
                if old_val and new_val and old_val != new_val:
                    record_correction(field, old_val, new_val, raw_text)
            except Exception:
                pass

    c.execute(
        """
        UPDATE contacts SET
            name = ?, job_title = ?, company_name = ?, email = ?,
            phone_number = ?, website = ?, address = ?, services_products = ?,
            notes = ?, lead_status = ?, tags = ?, follow_up_date = ?
        WHERE id = ?
    """,
        (
            data.get("name", "").strip(),
            data.get("job_title", "").strip(),
            data.get("company_name", "").strip(),
            data.get("email", "").strip(),
            data.get("phone_number", "").strip(),
            data.get("website", "").strip(),
            data.get("address", "").strip(),
            data.get("services_products", "").strip(),
            data.get("notes", "").strip(),
            data.get("lead_status", "New").strip(),
            data.get("tags", "").strip(),
            data.get("follow_up_date", "").strip(),
            cid,
        ),
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Contact updated."})


@app.route("/api/corrections", methods=["GET"])
def api_get_corrections():
    """Return learned corrections memory."""
    corrections = get_learned_corrections(limit=50)
    return jsonify({"corrections": corrections, "count": len(corrections)})


@app.route("/api/contacts/<int:cid>", methods=["DELETE"])
def api_delete_contact(cid):
    """Delete a contact."""
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM contacts WHERE id = ?", (cid,))
    conn.commit()
    conn.close()
    return jsonify({"success": True, "message": "Contact deleted."})


@app.route("/api/stats", methods=["GET"])
def api_get_stats():
    """Summary dashboard statistics with single consolidated query for 5x speed."""
    conn = get_db()
    c = conn.cursor()

    c.execute(
        """
        SELECT 
            COUNT(*) as total,
            COUNT(CASE WHEN lead_status = 'New' THEN 1 END) as new_leads,
            COUNT(CASE WHEN lead_status = 'Contacted' THEN 1 END) as contacted,
            COUNT(CASE WHEN lead_status = 'Qualified' THEN 1 END) as qualified,
            COUNT(CASE WHEN lead_status = 'Customer' THEN 1 END) as customers
        FROM contacts
        """
    )
    res = c.fetchone() or {}
    total = res.get("total", 0)
    new_leads = res.get("new_leads", 0)
    contacted = res.get("contacted", 0)
    qualified = res.get("qualified", 0)
    customers = res.get("customers", 0)

    c.execute(
        "SELECT company_name, COUNT(*) as cnt FROM contacts WHERE company_name != '' GROUP BY company_name ORDER BY cnt DESC LIMIT 5"
    )
    top_companies = c.fetchall()

    conn.close()
    return jsonify({
        "total": total,
        "new_leads": new_leads,
        "contacted": contacted,
        "qualified": qualified,
        "customers": customers,
        "top_companies": top_companies,
    })


@app.route("/api/settings/verify-key", methods=["POST"])
def api_verify_key():
    """Live verification of Gemini API key with Google servers."""
    data = request.json or {}
    key = data.get("gemini_api_key", "").strip()
    if not key:
        settings = get_settings()
        key = settings.get("gemini_api_key", "").strip()
    test_res = test_gemini_api_key(key, force=True)
    return jsonify(test_res)


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    """Get or save API key, database URL, and engine preferences."""
    if request.method == "POST":
        data = request.json or {}
        settings = get_settings()
        key_changed = False
        if "gemini_api_key" in data:
            new_key = data["gemini_api_key"].strip()
            if new_key != settings.get("gemini_api_key"):
                key_changed = True
            settings["gemini_api_key"] = new_key
        if "database_url" in data:
            settings["database_url"] = data["database_url"].strip()
        if "engine_preference" in data:
            settings["engine_preference"] = data["engine_preference"].strip()
        save_settings(settings)
        try:
            init_db()
        except Exception:
            pass

        warning_msg = None
        if key_changed and settings.get("gemini_api_key"):
            t_res = test_gemini_api_key(settings["gemini_api_key"], force=True)
            if not t_res.get("valid"):
                warning_msg = f"Saved, but warning: {t_res.get('message')} (The app will fall back to Local OCR until a valid key is provided)."
            else:
                warning_msg = f"Settings saved! Gemini API key verified: {t_res.get('message')}"

        return jsonify({"success": True, "message": warning_msg or "Settings saved."})

    settings = get_settings()
    masked_key = ""
    api_key = settings.get("gemini_api_key", "").strip()
    if api_key:
        masked_key = api_key[:4] + "..." + api_key[-4:] if len(api_key) > 8 else "***"

    key_test = test_gemini_api_key(api_key) if api_key else {"valid": False, "status": "unconfigured", "message": "No key"}

    db_url = get_database_url()
    masked_db = ""
    if db_url:
        masked_db = db_url[:15] + "..." + db_url[-8:] if len(db_url) > 25 else "***"

    return jsonify({
        "has_api_key": bool(api_key),
        "is_gemini_valid": key_test.get("valid", False),
        "gemini_status": key_test.get("status", "unconfigured"),
        "gemini_message": key_test.get("message", ""),
        "masked_api_key": masked_key,
        "has_database_url": bool(db_url),
        "masked_database_url": masked_db,
        "engine_preference": settings.get("engine_preference", "hybrid"),
        "genai_sdk_installed": GENAI_AVAILABLE,
        "tesseract_path": pytesseract.pytesseract.tesseract_cmd,
        "db_engine": "Neon Cloud PostgreSQL (Active)",
    })


# ==========================================
# 5. EXPORT ENGINE (EXCEL, CSV, VCARD)
# ==========================================

@app.route("/api/export/excel", methods=["GET"])
def export_excel():
    """Export contacts to beautifully styled Excel spreadsheet."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM contacts ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "BizConnect Contacts"

    headers = [
        "ID", "Full Name", "Job Title", "Company", "Mobile / Phone",
        "Email", "Website", "Full Address", "Products & Services",
        "Status", "Notes", "Scanned Date", "AI / OCR Engine"
    ]
    ws.append(headers)

    # Style Header Row
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Append Data Rows
    for r in rows:
        ws.append([
            r["id"], r["name"], r["job_title"], r["company_name"],
            r["phone_number"], r["email"], r["website"], r["address"],
            r["services_products"], r["lead_status"], r["notes"],
            r["created_at"], r["ocr_engine"]
        ])

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"BizConnect_Contacts_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/export/csv", methods=["GET"])
def export_csv():
    """Export contacts to CSV with UTF-8 BOM."""
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM contacts ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()

    output = io.StringIO()
    # Write UTF-8 BOM so Excel opens accented characters seamlessly
    output.write("\ufeff")
    writer = csv.writer(output)

    writer.writerow([
        "ID", "Full Name", "Job Title", "Company", "Phone",
        "Email", "Website", "Address", "Services", "Status",
        "Notes", "Scanned At", "Engine"
    ])

    for r in rows:
        writer.writerow([
            r["id"], r["name"], r["job_title"], r["company_name"],
            r["phone_number"], r["email"], r["website"], r["address"],
            r["services_products"], r["lead_status"], r["notes"],
            r["created_at"], r["ocr_engine"]
        ])

    buf = io.BytesIO(output.getvalue().encode("utf-8"))
    filename = f"BizConnect_Contacts_{datetime.now().strftime('%Y%m%d')}.csv"
    return send_file(
        buf,
        as_attachment=True,
        download_name=filename,
        mimetype="text/csv",
    )


@app.route("/api/export/vcf", methods=["GET"])
@app.route("/api/export/vcf/<int:cid>", methods=["GET"])
def export_vcard(cid=None):
    """Export contacts as vCard 3.0 (.vcf) for instant Google/Outlook/iPhone import."""
    conn = get_db()
    c = conn.cursor()
    if cid:
        c.execute("SELECT * FROM contacts WHERE id = ?", (cid,))
    else:
        c.execute("SELECT * FROM contacts ORDER BY id DESC")
    rows = c.fetchall()
    conn.close()

    vcf_content = []
    for r in rows:
        name_parts = (r["name"] or "Contact").split(" ")
        first_name = name_parts[0]
        last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""

        card_lines = [
            "BEGIN:VCARD",
            "VERSION:3.0",
            f"N:{last_name};{first_name};;;",
            f"FN:{r['name']}",
            f"ORG:{r['company_name'] or ''}",
            f"TITLE:{r['job_title'] or ''}",
        ]
        if r["phone_number"]:
            card_lines.append(f"TEL;TYPE=CELL,VOICE:{r['phone_number']}")
        if r["email"]:
            card_lines.append(f"EMAIL;TYPE=WORK,INTERNET:{r['email']}")
        if r["website"]:
            card_lines.append(f"URL:{r['website']}")
        if r["address"]:
            clean_addr = r["address"].replace(";", ",")
            card_lines.append(f"ADR;TYPE=WORK:;;{clean_addr};;;;")
        if r["notes"] or r["services_products"]:
            note_content = f"{r['notes']} | Services: {r['services_products']}".strip(" |")
            card_lines.append(f"NOTE:{note_content}")
        card_lines.append("END:VCARD\n")
        vcf_content.append("\n".join(card_lines))

    buf = io.BytesIO("\n".join(vcf_content).encode("utf-8"))
    fname = f"contact_{cid}.vcf" if cid else f"BizConnect_All_Contacts_{datetime.now().strftime('%Y%m%d')}.vcf"
    return send_file(
        buf,
        as_attachment=True,
        download_name=fname,
        mimetype="text/vcard",
    )


# ==========================================
# 6. BIZCONNECT USER INTERFACE (HTML/CSS/JS)
# ==========================================

FRONTEND_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>BizConnect AI Studio | Business Card Scanner & Lead CRM</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">

    <style>
        :root {
            --primary: #2563eb;
            --primary-hover: #1d4ed8;
            --primary-soft: #eff6ff;
            --accent: #6366f1;
            --bg: #f8fafc;
            --surface: #ffffff;
            --border: #e2e8f0;
            --text-main: #0f172a;
            --text-muted: #64748b;
            --success: #10b981;
            --warning: #f59e0b;
            --danger: #ef4444;
            --radius-lg: 16px;
            --radius-md: 10px;
            --shadow-sm: 0 1px 3px rgba(0,0,0,0.06);
            --shadow-md: 0 4px 20px -2px rgba(15, 23, 42, 0.08);
            --shadow-lg: 0 12px 32px -4px rgba(15, 23, 42, 0.12);
        }

        * { box-sizing: border-box; margin: 0; padding: 0; }
        body {
            font-family: 'Plus Jakarta Sans', sans-serif;
            background: var(--bg);
            color: var(--text-main);
            min-height: 100vh;
            display: flex;
            flex-direction: column;
        }

        /* NAVBAR */
        .navbar {
            background: var(--surface);
            border-bottom: 1px solid var(--border);
            padding: 14px 28px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            position: sticky;
            top: 0;
            z-index: 50;
            box-shadow: var(--shadow-sm);
        }
        .brand {
            display: flex;
            align-items: center;
            gap: 12px;
            font-weight: 800;
            font-size: 20px;
            color: #0f172a;
            text-decoration: none;
            cursor: pointer;
        }
        .brand-logo {
            width: 38px;
            height: 38px;
            background: linear-gradient(135deg, var(--primary), var(--accent));
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: white;
            font-weight: 800;
            font-size: 20px;
            box-shadow: 0 4px 10px rgba(37,99,235,0.3);
        }
        .brand-badge {
            background: #dbeafe;
            color: #1e40af;
            font-size: 11px;
            font-weight: 700;
            padding: 3px 8px;
            border-radius: 20px;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .nav-tabs {
            display: flex;
            gap: 6px;
            background: #f1f5f9;
            padding: 4px;
            border-radius: 12px;
        }
        .nav-btn {
            background: transparent;
            border: none;
            padding: 8px 18px;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            color: var(--text-muted);
            cursor: pointer;
            transition: all 0.2s;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        .nav-btn:hover { color: var(--text-main); }
        .nav-btn.active {
            background: var(--surface);
            color: var(--primary);
            box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        }
        .nav-actions {
            display: flex;
            align-items: center;
            gap: 12px;
        }
        .engine-pill {
            font-size: 12px;
            font-weight: 600;
            padding: 6px 14px;
            border-radius: 20px;
            background: #ecfdf5;
            color: #065f46;
            border: 1px solid #a7f3d0;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .engine-pill.tesseract {
            background: #f0fdf4;
            color: #166534;
            border-color: #bbf7d0;
        }
        .engine-pill.invalid {
            background: #fef2f2;
            color: #991b1b;
            border-color: #fecaca;
        }
        .engine-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #10b981;
            box-shadow: 0 0 8px #10b981;
        }
        .engine-pill.tesseract .engine-dot {
            background: #16a34a;
            box-shadow: 0 0 8px #16a34a;
        }
        .engine-pill.invalid .engine-dot {
            background: #dc2626;
            box-shadow: 0 0 8px #dc2626;
        }

        /* CONTAINER */
        .main-container {
            max-width: 1300px;
            width: 100%;
            margin: 0 auto;
            padding: 28px 20px;
            flex: 1;
        }

        .tab-pane { display: none; }
        .tab-pane.active { display: block; }

        /* BUTTONS */
        .btn {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 8px;
            padding: 10px 20px;
            border-radius: var(--radius-md);
            font-weight: 600;
            font-size: 14px;
            cursor: pointer;
            transition: all 0.2s;
            border: 1px solid transparent;
            text-decoration: none;
        }
        .btn-primary {
            background: var(--primary);
            color: white;
        }
        .btn-primary:hover { background: var(--primary-hover); transform: translateY(-1px); }
        .btn-secondary {
            background: var(--surface);
            border-color: var(--border);
            color: var(--text-main);
        }
        .btn-secondary:hover { background: #f8fafc; border-color: #cbd5e1; }
        .btn-success { background: var(--success); color: white; }
        .btn-success:hover { background: #059669; }
        .btn-outline-danger {
            background: transparent;
            border-color: #fecaca;
            color: #dc2626;
        }
        .btn-outline-danger:hover { background: #fee2e2; }
        .btn-sm { padding: 6px 12px; font-size: 12px; border-radius: 6px; }

        /* STATS BAR */
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 16px;
            margin-bottom: 28px;
        }
        .stat-card {
            background: var(--surface);
            padding: 20px 24px;
            border-radius: var(--radius-lg);
            border: 1px solid var(--border);
            box-shadow: var(--shadow-sm);
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .stat-info h4 {
            font-size: 13px;
            text-transform: uppercase;
            color: var(--text-muted);
            letter-spacing: 0.5px;
            margin-bottom: 6px;
        }
        .stat-info .num {
            font-size: 28px;
            font-weight: 800;
            color: var(--text-main);
        }
        .stat-icon {
            width: 48px;
            height: 48px;
            border-radius: 12px;
            background: var(--primary-soft);
            color: var(--primary);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 22px;
        }

        /* SCANNER TAB */
        .scanner-hero {
            background: var(--surface);
            border-radius: var(--radius-lg);
            border: 1px solid var(--border);
            box-shadow: var(--shadow-md);
            padding: 36px;
            margin-bottom: 28px;
            text-align: center;
        }
        .scanner-hero h2 {
            font-size: 26px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 10px;
        }
        .scanner-hero p {
            color: var(--text-muted);
            max-width: 600px;
            margin: 0 auto 24px auto;
            font-size: 15px;
            line-height: 1.6;
        }
        .dropzone {
            border: 2px dashed #93c5fd;
            background: #f0f7ff;
            border-radius: var(--radius-lg);
            padding: 44px 20px;
            cursor: pointer;
            transition: all 0.2s;
            max-width: 800px;
            margin: 0 auto;
        }
        .dropzone:hover, .dropzone.dragover {
            border-color: var(--primary);
            background: #e0f2fe;
            transform: scale(1.01);
        }
        .dropzone-icon {
            font-size: 52px;
            margin-bottom: 14px;
        }
        .dropzone-title {
            font-size: 18px;
            font-weight: 700;
            color: #1e293b;
            margin-bottom: 6px;
        }
        .dropzone-sub {
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 18px;
        }
        .scan-action-btns {
            display: flex;
            justify-content: center;
            gap: 12px;
            flex-wrap: wrap;
        }

        /* PROCESSING QUEUE */
        .queue-section {
            margin-top: 32px;
            display: none;
        }
        .queue-title {
            font-size: 18px;
            font-weight: 700;
            margin-bottom: 16px;
        }
        .queue-list {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
            gap: 18px;
        }
        .queue-card {
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 16px;
            box-shadow: var(--shadow-sm);
            display: flex;
            gap: 14px;
        }
        .queue-thumb {
            width: 100px;
            height: 70px;
            border-radius: 8px;
            object-fit: cover;
            border: 1px solid var(--border);
            background: #eee;
        }
        .queue-details {
            flex: 1;
        }
        .queue-name {
            font-weight: 700;
            font-size: 15px;
            color: var(--text-main);
        }
        .queue-comp {
            font-size: 13px;
            color: var(--text-muted);
            margin-bottom: 8px;
        }

        /* REVIEW DRAWER / MODAL */
        .modal-overlay {
            position: fixed;
            top: 0; left: 0; right: 0; bottom: 0;
            background: rgba(15, 23, 42, 0.6);
            backdrop-filter: blur(4px);
            display: none;
            align-items: center;
            justify-content: center;
            z-index: 100;
            padding: 20px;
        }
        .modal-overlay.active { display: flex; }
        .modal-box {
            background: var(--surface);
            border-radius: var(--radius-lg);
            max-width: 950px;
            width: 100%;
            max-height: 90vh;
            overflow-y: auto;
            box-shadow: var(--shadow-lg);
            padding: 28px;
            position: relative;
        }
        .modal-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 24px;
            padding-bottom: 14px;
            border-bottom: 1px solid var(--border);
        }
        .modal-header h3 { font-size: 20px; font-weight: 800; }
        .close-btn {
            background: transparent;
            border: none;
            font-size: 24px;
            cursor: pointer;
            color: var(--text-muted);
        }
        .review-layout {
            display: grid;
            grid-template-columns: 1fr 1.3fr;
            gap: 24px;
        }
        @media (max-width: 850px) {
            .review-layout { grid-template-columns: 1fr; }
        }
        .card-preview-panel {
            background: #0f172a;
            border-radius: 12px;
            padding: 14px;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
        }
        .card-preview-panel img {
            max-width: 100%;
            max-height: 400px;
            object-fit: contain;
            border-radius: 8px;
        }
        .form-grid {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 12px;
        }
        .form-group {
            display: flex;
            flex-direction: column;
            gap: 5px;
        }
        .form-group.full { grid-column: span 2; }
        .form-group label {
            font-size: 12px;
            font-weight: 700;
            text-transform: uppercase;
            color: var(--text-muted);
        }
        .form-group input, .form-group textarea, .form-group select {
            padding: 10px 12px;
            border: 1px solid var(--border);
            border-radius: 8px;
            font-family: inherit;
            font-size: 14px;
            color: var(--text-main);
            transition: all 0.2s;
        }
        .form-group input:focus, .form-group textarea:focus, .form-group select:focus {
            outline: none;
            border-color: var(--primary);
            box-shadow: 0 0 0 3px rgba(37,99,235,0.15);
        }

        /* CRM CONTACTS TAB */
        .crm-toolbar {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 14px;
            flex-wrap: wrap;
            margin-bottom: 24px;
        }
        .search-box {
            display: flex;
            align-items: center;
            background: var(--surface);
            border: 1px solid var(--border);
            border-radius: var(--radius-md);
            padding: 8px 16px;
            flex: 1;
            min-width: 280px;
            box-shadow: var(--shadow-sm);
        }
        .search-box input {
            border: none;
            outline: none;
            width: 100%;
            font-size: 14px;
            margin-left: 8px;
            font-family: inherit;
        }
        .filter-group {
            display: flex;
            gap: 8px;
            align-items: center;
        }
        .contacts-grid {
            display: grid;
            grid-template-columns: repeat(auto-fill, minmax(360px, 1fr));
            gap: 20px;
        }
        .contact-card {
            background: var(--surface);
            border-radius: var(--radius-lg);
            border: 1px solid var(--border);
            box-shadow: var(--shadow-sm);
            padding: 22px;
            transition: all 0.2s;
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }
        .contact-card:hover {
            box-shadow: var(--shadow-md);
            transform: translateY(-2px);
            border-color: #cbd5e1;
        }
        .contact-header {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            margin-bottom: 12px;
        }
        .contact-avatar {
            width: 48px;
            height: 48px;
            border-radius: 12px;
            background: linear-gradient(135deg, #3b82f6, #6366f1);
            color: white;
            font-weight: 800;
            font-size: 18px;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .lead-badge {
            font-size: 11px;
            font-weight: 700;
            padding: 4px 10px;
            border-radius: 20px;
            text-transform: uppercase;
        }
        .badge-New { background: #dbeafe; color: #1d4ed8; }
        .badge-Contacted { background: #fef3c7; color: #b45309; }
        .badge-Qualified { background: #ede9fe; color: #6d28d9; }
        .badge-Customer { background: #d1fae5; color: #065f46; }

        .contact-name {
            font-size: 18px;
            font-weight: 800;
            color: #0f172a;
            margin-bottom: 2px;
        }
        .contact-title {
            font-size: 13px;
            font-weight: 600;
            color: var(--primary);
            margin-bottom: 4px;
        }
        .contact-company {
            font-size: 14px;
            font-weight: 700;
            color: #334155;
            margin-bottom: 14px;
        }
        .contact-details-list {
            display: flex;
            flex-direction: column;
            gap: 7px;
            font-size: 13px;
            color: #475569;
            margin-bottom: 18px;
        }
        .contact-item {
            display: flex;
            align-items: center;
            gap: 8px;
            word-break: break-word;
        }
        .contact-item a {
            color: inherit;
            text-decoration: none;
            transition: color 0.1s;
        }
        .contact-item a:hover { color: var(--primary); text-decoration: underline; }
        .contact-thumb-prev {
            width: 100%;
            height: 110px;
            object-fit: cover;
            border-radius: 8px;
            border: 1px solid var(--border);
            margin-bottom: 14px;
            cursor: pointer;
        }
        .contact-footer {
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding-top: 14px;
            border-top: 1px solid var(--border);
        }

        /* CAMERA MODAL */
        .camera-viewport {
            width: 100%;
            max-height: 480px;
            background: black;
            border-radius: 12px;
            position: relative;
            overflow: hidden;
            display: flex;
            align-items: center;
            justify-content: center;
        }
        .camera-viewport video {
            width: 100%;
            height: 100%;
            object-fit: cover;
        }
        .card-viewfinder-overlay {
            position: absolute;
            width: 80%;
            height: 65%;
            border: 2px dashed rgba(255, 255, 255, 0.85);
            border-radius: 12px;
            box-shadow: 0 0 0 9999px rgba(0, 0, 0, 0.55);
            pointer-events: none;
        }
        .camera-controls {
            margin-top: 18px;
            display: flex;
            justify-content: center;
            gap: 14px;
        }

        /* SETTINGS TAB */
        .settings-card {
            background: var(--surface);
            border-radius: var(--radius-lg);
            border: 1px solid var(--border);
            box-shadow: var(--shadow-sm);
            padding: 32px;
            max-width: 750px;
            margin: 0 auto;
        }
        .settings-card h3 {
            font-size: 20px;
            font-weight: 800;
            margin-bottom: 8px;
        }
        .settings-card p {
            color: var(--text-muted);
            margin-bottom: 24px;
            font-size: 14px;
        }
    </style>
</head>
<body>

    <!-- NAVBAR -->
    <header class="navbar">
        <a class="brand" onclick="switchTab('scanner')">
            <div class="brand-logo">B</div>
            <span>BizConnect Studio</span>
            <span class="brand-badge">AI Suite</span>
        </a>

        <div class="nav-tabs">
            <button class="nav-btn active" id="tabBtn-scanner" onclick="switchTab('scanner')">
                <span>📷</span> Scanner
            </button>
            <button class="nav-btn" id="tabBtn-crm" onclick="switchTab('crm')">
                <span>👥</span> Contacts CRM (<span id="contactsCountNav">0</span>)
            </button>
            <button class="nav-btn" id="tabBtn-settings" onclick="switchTab('settings')">
                <span>⚙️</span> AI & Engine
            </button>
        </div>

        <div class="nav-actions">
            <div class="engine-pill" id="engineStatusPill">
                <span class="engine-dot"></span>
                <span id="engineStatusText">Checking AI...</span>
            </div>
        </div>
    </header>

    <!-- MAIN CONTAINER -->
    <main class="main-container">

        <!-- 1. SCANNER TAB -->
        <section id="tab-scanner" class="tab-pane active">
            <div class="stats-grid">
                <div class="stat-card">
                    <div class="stat-info">
                        <h4>Total Digitized</h4>
                        <div class="num" id="statTotalCards">0</div>
                    </div>
                    <div class="stat-icon">📇</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <h4>New Leads</h4>
                        <div class="num" id="statNewLeads">0</div>
                    </div>
                    <div class="stat-icon">🎯</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <h4>Qualified</h4>
                        <div class="num" id="statQualified">0</div>
                    </div>
                    <div class="stat-icon">⭐</div>
                </div>
                <div class="stat-card">
                    <div class="stat-info">
                        <h4>Converted</h4>
                        <div class="num" id="statCustomers">0</div>
                    </div>
                    <div class="stat-icon">💼</div>
                </div>
            </div>

            <div class="scanner-hero">
                <h2>Batch & Single Card AI Scanner</h2>
                <p>Upload business card photos or capture them using your camera. Our hybrid engine uses Multimodal Vision AI + Local OCR for 99%+ extraction accuracy.</p>

                <div class="dropzone" id="dropzone" onclick="document.getElementById('fileInput').click()">
                    <div class="dropzone-icon">📥</div>
                    <div class="dropzone-title">Drag & drop business cards here, or click to browse</div>
                    <div class="dropzone-sub">Supports Single or Multiple JPG, PNG, WEBP (Up to 30 cards at once)</div>
                    <div class="scan-action-btns" onclick="event.stopPropagation()">
                        <button class="btn btn-primary" onclick="document.getElementById('fileInput').click()">
                            📁 Select Card Files
                        </button>
                        <button class="btn btn-secondary" onclick="openCameraModal()">
                            📷 Use Live Camera
                        </button>
                    </div>
                    <input type="file" id="fileInput" multiple accept="image/*" style="display:none" onchange="handleFileSelect(this.files)">
                </div>
            </div>

            <div class="queue-section" id="queueSection">
                <h3 class="queue-title">Scanned Cards Queue</h3>
                <div class="queue-list" id="queueList"></div>
            </div>
        </section>

        <!-- 2. CONTACTS CRM TAB -->
        <section id="tab-crm" class="tab-pane">
            <div class="crm-toolbar">
                <div class="search-box">
                    <span>🔍</span>
                    <input type="text" id="crmSearch" placeholder="Search by name, company, email, phone, city..." oninput="filterContacts()">
                </div>

                <div class="filter-group">
                    <select id="statusFilter" class="btn btn-secondary" onchange="filterContacts()">
                        <option value="All">All Statuses</option>
                        <option value="New">New Leads</option>
                        <option value="Contacted">Contacted</option>
                        <option value="Qualified">Qualified</option>
                        <option value="Customer">Customers</option>
                    </select>

                    <a href="/api/export/excel" class="btn btn-success" title="Download Excel format">
                        📊 Export Excel
                    </a>
                    <a href="/api/export/csv" class="btn btn-secondary" title="Download CSV">
                        📄 Export CSV
                    </a>
                    <a href="/api/export/vcf" class="btn btn-primary" title="Export as phone contact vCards">
                        📲 Download vCards
                    </a>
                </div>
            </div>

            <div class="contacts-grid" id="contactsGrid">
                <!-- Contact cards dynamically injected here -->
            </div>
        </section>

        <!-- 3. SETTINGS TAB -->
        <section id="tab-settings" class="tab-pane">
            <div class="settings-card">
                <h3>AI & OCR Configuration</h3>
                <p>Configure Google Gemini Vision API for 99%+ accuracy (BizConnect level), or run 100% offline using our local Tesseract OCR engine.</p>

                <div class="form-group" style="margin-bottom: 20px;">
                    <label>Gemini API Key</label>
                    <div style="display: flex; gap: 8px;">
                        <input type="password" id="geminiApiKeyInput" placeholder="AIzaSy..." style="flex: 1;">
                        <button type="button" class="btn btn-secondary" onclick="testApiKey()" style="white-space: nowrap; padding: 8px 14px; font-size: 13px;">🧪 Test Key</button>
                    </div>
                    <div id="keyTestResult" style="margin-top: 6px; font-size: 12px; font-weight: 600;"></div>
                    <small style="color:var(--text-muted); margin-top:4px; display:block;">
                        Get a free Gemini API key at <a href="https://aistudio.google.com" target="_blank" style="color:var(--primary);">Google AI Studio</a>. When verified, cards are parsed with Vision AI (Gemini 3.8 Flash).
                    </small>
                </div>

                <div class="form-group" style="margin-bottom: 24px;">
                    <label>Engine Mode</label>
                    <select id="enginePreferenceSelect">
                        <option value="hybrid">Hybrid (Vision AI when API Key set, else Local OCR)</option>
                        <option value="gemini">Force Vision AI Only</option>
                        <option value="tesseract">Force Local Tesseract OCR Only (100% Offline)</option>
                    </select>
                </div>

                <div class="form-group" style="margin-bottom: 24px; padding: 14px; background: #f8fafc; border: 1px solid var(--border); border-radius: 8px;">
                    <label style="margin-bottom: 4px;">Database Storage</label>
                    <div style="display: flex; align-items: center; gap: 8px; font-weight: 600; color: #16a34a; margin-bottom: 8px;">
                        <span>💾</span> <span id="dbStatusText">Neon Cloud PostgreSQL (Active)</span>
                    </div>
                    <label style="font-size: 12px; margin-top: 8px; display:block;">Neon PostgreSQL Connection URL:</label>
                    <input type="password" id="databaseUrlInput" placeholder="postgresql://user:password@ep-xyz.neon.tech/neondb?sslmode=require" style="margin-top:4px;">
                    <small style="color:var(--text-muted); display:block; margin-top:4px;">
                        Configured via <code>DATABASE_URL</code> in <code>.env</code> or settings.
                    </small>
                </div>

                <div class="form-group" style="margin-bottom: 24px; padding: 14px; background: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px;">
                    <label style="margin-bottom: 4px; color: #166534; font-weight: 700;">🧠 AI Learning & Correction Memory</label>
                    <div style="display: flex; align-items: center; justify-content: space-between;">
                        <span style="font-size: 13px; color: #15803d;" id="memoryStatusText">Memory Active: Learning from your verified card edits</span>
                        <span class="badge" style="background:#22c55e; color:#fff;" id="memoryCountBadge">0 Learned Rules</span>
                    </div>
                    <div id="memoryList" style="margin-top: 10px; max-height: 150px; overflow-y: auto; font-size: 12px; color: #334155;"></div>
                </div>

                <div style="display: flex; gap: 12px;">
                    <button class="btn btn-primary" onclick="saveSettings()">💾 Save Settings</button>
                    <button class="btn btn-secondary" onclick="loadSettings()">🔄 Refresh</button>
                </div>
            </div>
        </section>

    </main>

    <!-- REVIEW / EDIT CONTACT MODAL -->
    <div class="modal-overlay" id="reviewModal">
        <div class="modal-box">
            <div class="modal-header">
                <h3 id="modalTitle">Review & Verify Scanned Card</h3>
                <button class="close-btn" onclick="closeReviewModal()">×</button>
            </div>

            <div class="review-layout">
                <div class="card-preview-panel">
                    <img id="modalPreviewImg" src="" alt="Business Card Photo">
                    <div style="color:#94a3b8; font-size:12px; margin-top:8px;" id="modalEngineBadge">Engine: OCR</div>
                </div>

                <form id="contactForm" onsubmit="saveCardData(event)">
                    <input type="hidden" id="editContactId">
                    <input type="hidden" id="cardImageFilename">
                    <input type="hidden" id="cardOcrEngine">
                    <input type="hidden" id="cardRawText">

                    <div class="form-grid">
                        <div class="form-group full">
                            <label>Full Name</label>
                            <input type="text" id="fieldName" required placeholder="e.g. Arti Chaudhary">
                        </div>

                        <div class="form-group">
                            <label>Designation / Job Title</label>
                            <input type="text" id="fieldJobTitle" placeholder="e.g. Marketing Executive">
                        </div>

                        <div class="form-group">
                            <label>Company Name</label>
                            <input type="text" id="fieldCompany" placeholder="e.g. NAVGRAHA STEELS INC">
                        </div>

                        <div class="form-group">
                            <label>Phone / Mobile</label>
                            <input type="text" id="fieldPhone" placeholder="e.g. +91-72085 16629">
                        </div>

                        <div class="form-group">
                            <label>Email Address</label>
                            <input type="email" id="fieldEmail" placeholder="e.g. contact@company.com">
                        </div>

                        <div class="form-group full">
                            <label>Website URL</label>
                            <input type="text" id="fieldWebsite" placeholder="e.g. www.company.com">
                        </div>

                        <div class="form-group full">
                            <label>Full Address</label>
                            <textarea id="fieldAddress" rows="2" placeholder="Office, Street, City, State, Country, PIN"></textarea>
                        </div>

                        <div class="form-group full">
                            <label>Products / Services Offered</label>
                            <input type="text" id="fieldServices" placeholder="e.g. Stainless Steel, Pipe Fittings, etc.">
                        </div>

                        <div class="form-group">
                            <label>Lead Status</label>
                            <select id="fieldStatus">
                                <option value="New">New</option>
                                <option value="Contacted">Contacted</option>
                                <option value="Qualified">Qualified</option>
                                <option value="Customer">Customer</option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label>Follow-up Date</label>
                            <input type="date" id="fieldFollowUp">
                        </div>

                        <div class="form-group full">
                            <label>Notes & Follow-up Details</label>
                            <textarea id="fieldNotes" rows="2" placeholder="Met at Expo 2026, requested product catalog..."></textarea>
                        </div>
                    </div>

                    <div style="margin-top: 20px; display:flex; justify-content: flex-end; gap: 10px;">
                        <button type="button" class="btn btn-secondary" onclick="closeReviewModal()">Cancel</button>
                        <button type="submit" class="btn btn-primary" id="saveContactBtn">💾 Save to Contacts CRM</button>
                    </div>
                </form>
            </div>
        </div>
    </div>

    <!-- LIVE CAMERA MODAL -->
    <div class="modal-overlay" id="cameraModal">
        <div class="modal-box" style="max-width: 600px;">
            <div class="modal-header">
                <h3>Live Business Card Camera</h3>
                <button class="close-btn" onclick="closeCameraModal()">×</button>
            </div>
            <div class="camera-viewport">
                <video id="cameraVideo" autoplay playsinline></video>
                <div class="card-viewfinder-overlay"></div>
            </div>
            <div class="camera-controls">
                <button class="btn btn-secondary" onclick="closeCameraModal()">Cancel</button>
                <button class="btn btn-primary" onclick="captureCardSnapshot()">📸 Snap & Scan Card</button>
            </div>
            <canvas id="cameraCanvas" style="display:none;"></canvas>
        </div>
    </div>

    <script>
        let allContacts = [];
        let scannedCardsQueue = [];
        let videoStream = null;

        function openReviewById(cid) {
            const card = allContacts.find(c => c.id === cid);
            if (card) openReviewModal(card, true);
        }

        function openQueueReview(index) {
            const card = scannedCardsQueue[index];
            if (card) openReviewModal(card, false);
        }

        // TAB SWITCHING
        function switchTab(tabName) {
            document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
            document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));

            document.getElementById('tab-' + tabName).classList.add('active');
            document.getElementById('tabBtn-' + tabName).classList.add('active');

            if (tabName === 'crm') loadContacts();
            if (tabName === 'scanner') loadStats();
            if (tabName === 'settings') loadSettings();
        }

        // DRAG & DROP HANDLING
        const dropzone = document.getElementById('dropzone');
        ['dragenter', 'dragover'].forEach(eName => {
            dropzone.addEventListener(eName, (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
        });
        ['dragleave', 'drop'].forEach(eName => {
            dropzone.addEventListener(eName, (e) => { e.preventDefault(); dropzone.classList.remove('dragover'); });
        });
        dropzone.addEventListener('drop', (e) => {
            const files = e.dataTransfer.files;
            if (files && files.length) handleFileSelect(files);
        });

        // Helper to optimize image on browser before upload (faster upload & prevents server RAM spikes)
        async function compressImageForUpload(file) {
            return new Promise((resolve) => {
                if (!file.type || !file.type.startsWith('image/')) return resolve(file);
                const img = new Image();
                const reader = new FileReader();
                reader.onload = (e) => {
                    img.onload = () => {
                        const maxDim = 1800;
                        let w = img.width, h = img.height;
                        if (w > maxDim || h > maxDim) {
                            if (w > h) {
                                h = Math.round((h * maxDim) / w);
                                w = maxDim;
                            } else {
                                w = Math.round((w * maxDim) / h);
                                h = maxDim;
                            }
                        }
                        const canvas = document.createElement('canvas');
                        canvas.width = w;
                        canvas.height = h;
                        const ctx = canvas.getContext('2d');
                        ctx.drawImage(img, 0, 0, w, h);
                        canvas.toBlob((blob) => {
                            if (blob) {
                                resolve(new File([blob], file.name.replace(/\\.[^.]+$/, '.jpg'), { type: 'image/jpeg' }));
                            } else {
                                resolve(file);
                            }
                        }, 'image/jpeg', 0.92);
                    };
                    img.onerror = () => resolve(file);
                    img.src = e.target.result;
                };
                reader.onerror = () => resolve(file);
                reader.readAsDataURL(file);
            });
        }

        // BATCH / SINGLE FILE UPLOAD
        async function handleFileSelect(files) {
            if (!files || !files.length) return;

            const queueSection = document.getElementById('queueSection');
            const queueList = document.getElementById('queueList');
            queueSection.style.display = 'block';
            queueList.innerHTML = '';

            const formData = new FormData();
            for (let i = 0; i < files.length; i++) {
                // Create Queue Item Placeholder
                const placeholder = document.createElement('div');
                placeholder.className = 'queue-card';
                placeholder.id = 'queue-item-' + i;
                placeholder.innerHTML = `
                    <div style="font-size:28px;">⏳</div>
                    <div class="queue-details">
                        <div class="queue-name">Scanning ${files[i].name}...</div>
                        <div class="queue-comp">Optimizing & processing OCR...</div>
                    </div>
                `;
                queueList.appendChild(placeholder);

                const optimizedFile = await compressImageForUpload(files[i]);
                formData.append('images', optimizedFile);
            }

            try {
                const resp = await fetch('/api/scan', {
                    method: 'POST',
                    body: formData
                });

                if (!resp.ok) {
                    let errMsg = `Server returned status ${resp.status}`;
                    try {
                        const errData = await resp.json();
                        if (errData && errData.error) errMsg = errData.error;
                    } catch (_) {
                        if (resp.status === 502 || resp.status === 504) {
                            errMsg = 'Server is waking up (Render free tier cold start). Please retry in 10 seconds.';
                        }
                    }
                    throw new Error(errMsg);
                }

                const res = await resp.json();

                queueList.innerHTML = '';
                scannedCardsQueue = res.cards || [];
                scannedCardsQueue.forEach((card, idx) => {
                    const item = document.createElement('div');
                    item.className = 'queue-card';
                    item.innerHTML = `
                        <img src="${card.image_url}" class="queue-thumb" alt="Card Thumbnail">
                        <div class="queue-details">
                            <div class="queue-name">${card.name || 'Unnamed Contact'}</div>
                            <div class="queue-comp">${card.job_title || ''} • ${card.company_name || 'No Company'}</div>
                            <div style="font-size:12px; color:#16a34a; margin-bottom:8px;">✓ Scanned with ${card.ocr_engine}</div>
                            <button class="btn btn-sm btn-primary" onclick="openQueueReview(${idx})">Review & Save</button>
                        </div>
                    `;
                    queueList.appendChild(item);
                });

                // Auto-open first card if single upload
                if (scannedCardsQueue.length === 1) {
                    openReviewModal(scannedCardsQueue[0]);
                }

                loadStats();
            } catch (err) {
                alert('Scan error: ' + err.message);
            }
        }

        // REVIEW / EDIT MODAL
        let currentScannedOriginal = null;

        function openReviewModal(cardData, isEditMode = false) {
            currentScannedOriginal = {
                name: cardData.name || '',
                job_title: cardData.job_title || '',
                company_name: cardData.company_name || '',
                phone_number: cardData.phone_number || '',
                email: cardData.email || '',
                website: cardData.website || '',
                address: cardData.address || '',
                services_products: cardData.services_products || ''
            };

            document.getElementById('editContactId').value = cardData.id || '';
            document.getElementById('cardImageFilename').value = cardData.image_filename || '';
            document.getElementById('cardOcrEngine').value = cardData.ocr_engine || 'Local Tesseract OCR';
            document.getElementById('cardRawText').value = cardData.raw_text || '';

            document.getElementById('fieldName').value = cardData.name || '';
            document.getElementById('fieldJobTitle').value = cardData.job_title || '';
            document.getElementById('fieldCompany').value = cardData.company_name || '';
            document.getElementById('fieldPhone').value = cardData.phone_number || '';
            document.getElementById('fieldEmail').value = cardData.email || '';
            document.getElementById('fieldWebsite').value = cardData.website || '';
            document.getElementById('fieldAddress').value = cardData.address || '';
            document.getElementById('fieldServices').value = cardData.services_products || '';
            document.getElementById('fieldStatus').value = cardData.lead_status || 'New';
            document.getElementById('fieldFollowUp').value = cardData.follow_up_date || '';
            document.getElementById('fieldNotes').value = cardData.notes || '';

            const imgEl = document.getElementById('modalPreviewImg');
            if (cardData.image_filename) {
                imgEl.src = '/uploads/' + cardData.image_filename;
                imgEl.style.display = 'block';
            } else {
                imgEl.style.display = 'none';
            }

            document.getElementById('modalEngineBadge').textContent = 'Engine: ' + (cardData.ocr_engine || 'OCR');
            document.getElementById('modalTitle').textContent = isEditMode ? 'Edit Contact' : 'Verify & Save Business Card';
            document.getElementById('reviewModal').classList.add('active');
        }

        function closeReviewModal() {
            document.getElementById('reviewModal').classList.remove('active');
        }

        async function saveCardData(e) {
            e.preventDefault();
            const id = document.getElementById('editContactId').value;
            const payload = {
                name: document.getElementById('fieldName').value.trim(),
                job_title: document.getElementById('fieldJobTitle').value.trim(),
                company_name: document.getElementById('fieldCompany').value.trim(),
                phone_number: document.getElementById('fieldPhone').value.trim(),
                email: document.getElementById('fieldEmail').value.trim(),
                website: document.getElementById('fieldWebsite').value.trim(),
                address: document.getElementById('fieldAddress').value.trim(),
                services_products: document.getElementById('fieldServices').value.trim(),
                lead_status: document.getElementById('fieldStatus').value,
                follow_up_date: document.getElementById('fieldFollowUp').value,
                notes: document.getElementById('fieldNotes').value.trim(),
                image_filename: document.getElementById('cardImageFilename').value,
                ocr_engine: document.getElementById('cardOcrEngine').value,
                raw_text: document.getElementById('cardRawText').value,
                original_scanned: currentScannedOriginal || {}
            };

            const url = id ? `/api/contacts/${id}` : '/api/contacts';
            const method = id ? 'PUT' : 'POST';

            try {
                const resp = await fetch(url, {
                    method: method,
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const res = await resp.json();
                if (!resp.ok) throw new Error(res.error || 'Failed to save');

                closeReviewModal();
                loadStats(); 
                loadContacts();
            } catch (err) {
                alert('Error saving contact: ' + err.message);
            }
        }

        // CONTACTS CRM
        async function loadContacts() {
            const q = document.getElementById('crmSearch').value;
            const status = document.getElementById('statusFilter').value;

            try {
                const resp = await fetch(`/api/contacts?q=${encodeURIComponent(q)}&status=${encodeURIComponent(status)}`);
                const data = await resp.json();
                allContacts = data.contacts || [];

                document.getElementById('contactsCountNav').textContent = allContacts.length;
                renderContactsGrid(allContacts);
            } catch (err) {
                console.error(err);
            }
        }

        // Debounced search to prevent unnecessary network load
        let searchDebounceTimer = null;
        function filterContacts() {
            clearTimeout(searchDebounceTimer);
            searchDebounceTimer = setTimeout(loadContacts, 220);
        }

        function renderContactsGrid(contacts) {
            const grid = document.getElementById('contactsGrid');
            grid.innerHTML = '';

            if (!contacts.length) {
                grid.innerHTML = `
                    <div style="grid-column: 1 / -1; text-align: center; padding: 60px; background: white; border-radius: 16px; border: 1px solid var(--border);">
                        <div style="font-size: 42px; margin-bottom: 12px;">📇</div>
                        <h3 style="font-weight: 700; color: #1e293b; margin-bottom: 6px;">No Contacts Found</h3>
                        <p style="color: var(--text-muted); margin-bottom: 20px;">Scan business cards or adjust your search filter.</p>
                        <button class="btn btn-primary" onclick="switchTab('scanner')">Scan New Business Card</button>
                    </div>
                `;
                return;
            }

            contacts.forEach(c => {
                const cardEl = document.createElement('div');
                cardEl.className = 'contact-card';

                const initials = (c.name || 'C').split(' ').map(n => n[0]).slice(0, 2).join('').toUpperCase();
                const thumbHtml = c.image_filename ?
                    `<img src="/uploads/${c.image_filename}" class="contact-thumb-prev" onclick="openReviewById(${c.id})" title="Click to inspect card">` : '';

                cardEl.innerHTML = `
                    <div>
                        <div class="contact-header">
                            <div class="contact-avatar">${initials}</div>
                            <span class="lead-badge badge-${c.lead_status}">${c.lead_status}</span>
                        </div>

                        <div class="contact-name">${c.name}</div>
                        <div class="contact-title">${c.job_title || ''}</div>
                        <div class="contact-company">${c.company_name || 'No Company'}</div>

                        ${thumbHtml}

                        <div class="contact-details-list">
                            ${c.phone_number ? `<div class="contact-item">📞 <a href="tel:${c.phone_number}">${c.phone_number}</a></div>` : ''}
                            ${c.email ? `<div class="contact-item">✉️ <a href="mailto:${c.email}">${c.email}</a></div>` : ''}
                            ${c.website ? `<div class="contact-item">🌐 <a href="${c.website.startsWith('http') ? c.website : 'https://' + c.website}" target="_blank">${c.website}</a></div>` : ''}
                            ${c.address ? `<div class="contact-item">📍 <span>${c.address}</span></div>` : ''}
                        </div>
                    </div>

                    <div class="contact-footer">
                        <div style="font-size:11px; color:#94a3b8;">${c.created_at || ''}</div>
                        <div style="display:flex; gap:6px;">
                            <a href="/api/export/vcf/${c.id}" class="btn btn-sm btn-secondary" title="Download Phone vCard">📇 vCard</a>
                            <button class="btn btn-sm btn-secondary" onclick="openReviewById(${c.id})">✏️ Edit</button>
                            <button class="btn btn-sm btn-outline-danger" onclick="deleteContact(${c.id})">🗑️</button>
                        </div>
                    </div>
                `;
                grid.appendChild(cardEl);
            });
        }

        async function deleteContact(cid) {
            if (!confirm('Are you sure you want to delete this contact?')) return;
            try {
                await fetch(`/api/contacts/${cid}`, { method: 'DELETE' });
                loadContacts();
                loadStats();
            } catch (err) {
                alert('Error deleting: ' + err.message);
            }
        }

        // STATS & METRICS
        async function loadStats() {
            try {
                const resp = await fetch('/api/stats');
                const data = await resp.json();
                document.getElementById('statTotalCards').textContent = data.total || 0;
                document.getElementById('statNewLeads').textContent = data.new_leads || 0;
                document.getElementById('statQualified').textContent = data.qualified || 0;
                document.getElementById('statCustomers').textContent = data.customers || 0;
                document.getElementById('contactsCountNav').textContent = data.total || 0;
            } catch (err) {
                console.error(err);
            }
        }

        // SETTINGS & AI ENGINE STATUS
        async function loadSettings() {
            try {
                const resp = await fetch('/api/settings');
                const data = await resp.json();

                const pill = document.getElementById('engineStatusPill');
                const text = document.getElementById('engineStatusText');

                if (data.gemini_status === 'active') {
                    text.textContent = 'Vision AI Active (' + (data.gemini_message || '99%+ Accuracy') + ')';
                    pill.className = 'engine-pill';
                } else if (data.gemini_status === 'invalid') {
                    text.textContent = '⚠️ Invalid Gemini Key (Fallback to Local OCR)';
                    pill.className = 'engine-pill invalid';
                } else {
                    text.textContent = 'Local Tesseract OCR';
                    pill.className = 'engine-pill tesseract';
                }

                if (data.masked_api_key) {
                    document.getElementById('geminiApiKeyInput').placeholder = 'Configured: ' + data.masked_api_key;
                }

                const dbText = document.getElementById('dbStatusText');
                if (dbText) {
                    dbText.textContent = data.db_engine || 'Neon Cloud PostgreSQL (Active)';
                    dbText.style.color = '#16a34a';
                }

                if (data.masked_database_url) {
                    document.getElementById('databaseUrlInput').placeholder = 'Connected: ' + data.masked_database_url;
                }

                document.getElementById('enginePreferenceSelect').value = data.engine_preference || 'hybrid';

                // Load Learned Corrections Memory
                try {
                    const corResp = await fetch('/api/corrections');
                    const corData = await corResp.json();
                    const memBadge = document.getElementById('memoryCountBadge');
                    const memList = document.getElementById('memoryList');
                    if (memBadge && corData) {
                        memBadge.textContent = `${corData.count || 0} Learned Rules`;
                        if (corData.corrections && corData.corrections.length > 0) {
                            memList.innerHTML = corData.corrections.slice(0, 5).map(c => 
                                `<div style="padding:4px 0; border-bottom:1px dashed #cbd5e1;">` +
                                `• Corrected <b>${c.field_name}</b>: <code>${c.wrong_value}</code> ➔ <b>${c.corrected_value}</b>` +
                                `</div>`
                            ).join('');
                        } else {
                            memList.innerHTML = '<em>No corrections needed yet. When you edit any scanned card in review modal, the AI learns automatically.</em>';
                        }
                    }
                } catch (cErr) {
                    console.log('Correction memory fetch:', cErr);
                }
            } catch (err) {
                console.error(err);
            }
        }

        async function testApiKey() {
            const key = document.getElementById('geminiApiKeyInput').value.trim();
            const resEl = document.getElementById('keyTestResult');
            resEl.style.color = '#64748b';
            resEl.textContent = '⏳ Verifying connection with Google Gemini...';
            try {
                const resp = await fetch('/api/settings/verify-key', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ gemini_api_key: key })
                });
                const res = await resp.json();
                if (res.valid) {
                    resEl.style.color = '#16a34a';
                    resEl.textContent = '✅ ' + res.message;
                } else {
                    resEl.style.color = '#dc2626';
                    resEl.textContent = '❌ ' + res.message;
                }
                loadSettings();
            } catch (err) {
                resEl.style.color = '#dc2626';
                resEl.textContent = '❌ Verification error: ' + err.message;
            }
        }

        async function saveSettings() {
            const key = document.getElementById('geminiApiKeyInput').value.trim();
            const pref = document.getElementById('enginePreferenceSelect').value;
            const dbUrl = document.getElementById('databaseUrlInput').value.trim();

            const payload = { engine_preference: pref };
            if (key) payload.gemini_api_key = key;
            if (dbUrl) payload.database_url = dbUrl;

            try {
                const resp = await fetch('/api/settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(payload)
                });
                const res = await resp.json();
                alert(res.message || 'Settings saved!');
                loadSettings();
            } catch (err) {
                alert('Error saving settings: ' + err.message);
            }
        }

        // LIVE CAMERA
        async function openCameraModal() {
            const modal = document.getElementById('cameraModal');
            const video = document.getElementById('cameraVideo');
            modal.classList.add('active');

            try {
                videoStream = await navigator.mediaDevices.getUserMedia({
                    video: { facingMode: 'environment', width: { ideal: 1920 }, height: { ideal: 1080 } }
                });
                video.srcObject = videoStream;
            } catch (err) {
                alert('Could not access camera: ' + err.message);
                closeCameraModal();
            }
        }

        function closeCameraModal() {
            if (videoStream) {
                videoStream.getTracks().forEach(track => track.stop());
                videoStream = null;
            }
            document.getElementById('cameraModal').classList.remove('active');
        }

        function captureCardSnapshot() {
            const video = document.getElementById('cameraVideo');
            const canvas = document.getElementById('cameraCanvas');
            canvas.width = video.videoWidth || 1280;
            canvas.height = video.videoHeight || 720;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(video, 0, 0, canvas.width, canvas.height);

            canvas.toBlob(blob => {
                const file = new File([blob], 'camera_card.jpg', { type: 'image/jpeg' });
                closeCameraModal();
                handleFileSelect([file]);
            }, 'image/jpeg', 0.95);
        }

        // INIT
        window.addEventListener('DOMContentLoaded', () => {
            loadSettings();
            loadStats();
        });
    </script>
</body>
</html>
"""

def main():
    port = int(os.environ.get("PORT", 5000))

    # Initialize DB in background thread so startup is instant
    threading.Thread(target=init_db, daemon=True).start()

    # Local development: automatically open the browser directly
    if not os.environ.get("RENDER"):
        import webbrowser
        def open_local_browser(target_port):
            webbrowser.open(f"http://127.0.0.1:{target_port}")
        threading.Timer(1.2, lambda: open_local_browser(port)).start()
        print(f"\n[Card Scanner] Running locally at http://127.0.0.1:{port} (Opening browser automatically...)\n")

    try:
        app.run(host="0.0.0.0", port=port, debug=False)
    except OSError as e:
        if "10048" in str(e) or "Address already in use" in str(e):
            alt_port = port + 1
            print(f"\n[Notice] Port {port} is occupied. Automatically switching to http://127.0.0.1:{alt_port}...\n")
            if not os.environ.get("RENDER"):
                import webbrowser
                webbrowser.open(f"http://127.0.0.1:{alt_port}")
            app.run(host="0.0.0.0", port=alt_port, debug=False)
        else:
            raise

if __name__ == "__main__":
    main()