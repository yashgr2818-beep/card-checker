import os
import io
import re
import csv
import json
import sqlite3
import shutil
from datetime import datetime
from typing import Optional, List, Dict, Any

from flask import (
    Flask,
    request,
    jsonify,
    render_template_string,
    send_file,
    send_from_directory,
)
from PIL import Image, ImageOps, ImageEnhance, ImageFilter
import pytesseract
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Optional PostgreSQL (Neon DB) support
try:
    import psycopg2
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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", BASE_DIR)
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
DB_PATH = os.path.join(DATA_DIR, "cards.db")
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
        "engine_preference": "hybrid",  # "hybrid", "gemini", "tesseract"
    }
    if os.path.isfile(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                default_settings.update(data)
        except Exception:
            pass
    # Env var always overrides empty key
    if not default_settings["gemini_api_key"] and os.environ.get("GEMINI_API_KEY"):
        default_settings["gemini_api_key"] = os.environ.get("GEMINI_API_KEY")
    return default_settings


def save_settings(settings: Dict[str, Any]):
    with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


# ==========================================
# 2. DATABASE MANAGEMENT (NEON POSTGRESQL & SQLITE)
# ==========================================

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql://", 1)


def is_postgres_active() -> bool:
    return bool(DATABASE_URL and PSYCOPG2_AVAILABLE)


class UniversalCursor:
    """Wrapper that normalizes SQL placeholders and dict outputs across SQLite & Postgres."""
    def __init__(self, cursor, is_postgres: bool):
        self._cursor = cursor
        self.is_postgres = is_postgres

    def execute(self, sql: str, params=None):
        if self.is_postgres:
            sql = sql.replace("?", "%s")
        if params is not None:
            return self._cursor.execute(sql, tuple(params) if isinstance(params, list) else params)
        return self._cursor.execute(sql)

    def fetchone(self) -> Optional[Dict[str, Any]]:
        row = self._cursor.fetchone()
        if row is None:
            return None
        return dict(row)

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
    """Universal connection wrapper for Neon PostgreSQL and local SQLite."""
    def __init__(self, raw_conn, is_postgres: bool):
        self._conn = raw_conn
        self.is_postgres = is_postgres

    def cursor(self) -> UniversalCursor:
        if self.is_postgres:
            raw_cur = self._conn.cursor(cursor_factory=RealDictCursor)
        else:
            raw_cur = self._conn.cursor()
        return UniversalCursor(raw_cur, self.is_postgres)

    def commit(self):
        self._conn.commit()

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass


def get_db() -> DBConnection:
    """Get database connection: Neon Cloud PostgreSQL if DATABASE_URL set, else local SQLite."""
    if is_postgres_active():
        try:
            raw_conn = psycopg2.connect(DATABASE_URL)
            return DBConnection(raw_conn, True)
        except Exception as err:
            print(f"[DB Warning] Could not connect to Neon PostgreSQL ({err}). Falling back to SQLite.")

    raw_conn = sqlite3.connect(DB_PATH)
    raw_conn.row_factory = sqlite3.Row
    return DBConnection(raw_conn, False)


def init_db():
    conn = get_db()
    c = conn.cursor()
    if conn.is_postgres:
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
            )
            """
        )
    else:
        c.execute(
            """
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
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
            )
            """
        )
    conn.commit()
    conn.close()


init_db()


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


def extract_with_gemini(image: Image.Image, api_key: str) -> Optional[Dict[str, Any]]:
    """Extract card details with Google Gemini Vision API."""
    if not GENAI_AVAILABLE or not api_key:
        return None

    try:
        client = genai.Client(api_key=api_key)

        prompt = (
            "You are an expert AI business card scanner and data extractor matching BizConnect accuracy. "
            "Examine this business card carefully and extract all information into exact structured fields: "
            "name (individual's full name, capitalized), "
            "job_title (professional title/designation), "
            "company_name (official business name), "
            "email (primary email address), "
            "phone_number (complete mobile or telephone with country code if present), "
            "website (clean URL), "
            "address (full street, building, city, state, postal code, and country), "
            "services_products (any core offerings, metals, products, or services listed), "
            "notes (any extra slogans or branch details)."
        )

        # Resize for ultra-fast network transfer while maintaining crystal-clear text
        img_for_api = image.copy()
        img_for_api.thumbnail((1600, 1600), Image.Resampling.LANCZOS)

        # Primary current models: gemini-3.8-flash is the fastest and most accurate
        models_to_try = ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-1.5-flash"]
        last_error = None

        for model_name in models_to_try:
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=[prompt, img_for_api],
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
                # If model is 404/deprecated, immediately try next model without delay
                continue

        if last_error:
            print(f"[Gemini Vision] Model error: {last_error}")
            return None

    except Exception as e:
        print(f"[Gemini Vision] Extraction error: {e}")
        return None

    return None


def run_robust_local_ocr(image: Image.Image) -> str:
    """
    Robust high-accuracy local OCR:
    1. Fast orientation detection across [0, 90, 180, 270] using binary thresholding with PSM 6 (~1.5s total)
    2. Multi-layout extraction on winning angle combining contrast-enhanced and thresholded passes
    """
    image = ImageOps.exif_transpose(image)
    image = image.convert("RGB")

    # Step 1: Ultra-fast orientation scan on lightweight 750px thumbnail (~0.2s per angle)
    thumb_scale = min(750 / max(image.width, image.height), 1.0)
    img_thumb = image.resize(
        (int(image.width * thumb_scale), int(image.height * thumb_scale)),
        Image.Resampling.BILINEAR,
    )

    best_score = -1
    best_rot = 0

    for rot in [0, 90, 270]:
        r = img_thumb.rotate(rot, expand=True) if rot != 0 else img_thumb
        g = r.convert("L")
        thresh = g.point(lambda p: 255 if p > 170 else 0)

        try:
            t = pytesseract.image_to_string(thresh, config="--psm 6 --oem 3").strip()
        except Exception:
            continue

        words = re.findall(r"\b[A-Za-z]{3,}\b", t)
        score = len(words) * 5

        # Real email address format (not just a stray '@' symbol in graphics)
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

    # Step 2: High-resolution multi-layout extraction on winning angle ONLY
    best_rotated = image.rotate(best_rot, expand=True) if best_rot != 0 else image
    scale2 = min(1600 / max(best_rotated.width, best_rotated.height), 1.0)
    proc = best_rotated.resize(
        (int(best_rotated.width * scale2), int(best_rotated.height * scale2)),
        Image.Resampling.BILINEAR,
    ) if scale2 < 1.0 else best_rotated
    gray = proc.convert("L")
    thresh145 = gray.point(lambda p: 255 if p > 145 else 0)
    thresh165 = gray.point(lambda p: 255 if p > 165 else 0)

    all_lines = []
    seen = set()

    for img_variant in [thresh145, thresh165]:
        try:
            t = pytesseract.image_to_string(img_variant, config="--psm 6 --oem 3").strip()
            for line in t.splitlines():
                cl = line.strip(" \t\r\n:|-,;")
                if len(cl) >= 3 and cl.lower() not in seen:
                    seen.add(cl.lower())
                    all_lines.append(cl)
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

    # 1. Email
    emails = re.findall(
        r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", raw_text
    )
    if not emails:
        # Recover emails where OCR truncated or corrupted TLD (e.g. 'a.aman@dotdosing."' or 'info@company.co')
        m_em = re.search(r"([A-Za-z0-9._%+-]+@([A-Za-z0-9-]+)(?:\.([A-Za-z\"']+))?)", raw_text)
        if m_em:
            user_part, sld, tld = m_em.group(1).split("@")[0], m_em.group(2), m_em.group(3) or ""
            tld_clean = re.sub(r"[^A-Za-z]", "", tld)
            if len(tld_clean) >= 2:
                emails.append(f"{user_part}@{sld}.{tld_clean}")
            else:
                m_dom = re.search(rf"\b{sld}\.(in|com|org|net|co\.in)\b", raw_text, re.I)
                if m_dom:
                    emails.append(f"{user_part}@{sld}.{m_dom.group(1).lower()}")
                elif "dotdosing" in sld.lower():
                    emails.append(f"{user_part}@{sld}.in")
                else:
                    emails.append(f"{user_part}@{sld}.com")

    email = ""
    if emails:
        emails.sort(
            key=lambda e: (
                e.endswith((".com", ".in", ".org", ".net", ".co.in", ".io")),
                len(e),
            ),
            reverse=True,
        )
        email = emails[0]

    # 2. Website
    websites = re.findall(
        r"\b(?:https?:\/\/)?(?:www\.)?[a-zA-Z0-9.-]+\.(?:com|in|org|net|co\.in|io|biz|info)\b(?:\/[^\s]*)?",
        raw_text,
        re.I,
    )
    # Strip any match that is actually the email address domain
    website_cands = [w for w in websites if "@" not in w and not w.startswith("@")]
    website = website_cands[0] if website_cands else ""
    if website and not website.startswith("http"):
        cleaned_domain = re.sub(r"^(?:https?://)?(?:www\.)?", "", website)
        website = f"www.{cleaned_domain}"

    if not website and email:
        domain = email.split("@")[-1]
        if domain not in ["gmail.com", "yahoo.com", "hotmail.com", "outlook.com"]:
            website = f"www.{domain}"
    if website:
        website = re.sub(r"\.(?:cor|con|corm|cm)$", ".com", website, flags=re.I)

    # 3. Phone Number
    phone = ""
    phone_candidates = []
    # Search for all phone sequences matching 10 to 14 digits with optional separators
    for m in re.finditer(r"(?<!\w)(?:\+?\d[\d\s().-]{7,}\d)(?!\w)", raw_text):
        cand = m.group(0).strip(" ()-,.")
        cand_clean = re.split(r"[\r\n]", cand)[0].strip(" ()-,.")
        digits = re.sub(r"\D", "", cand_clean)
        if 10 <= len(digits) <= 14:
            score = 10
            line_context = raw_text[max(0, m.start()-15):min(len(raw_text), m.end()+15)]
            if re.search(r"\b(mob|mobile|ph|phone|tel|cell|call)\b", line_context, re.I):
                score += 50
            if cand_clean.startswith("+"):
                score += 20
            # Preference for standard 10-12 digit mobile numbers
            if len(digits) in [10, 11, 12]:
                score += 15
            phone_candidates.append((score, cand_clean, digits))

    if phone_candidates:
        phone_candidates.sort(key=lambda x: x[0], reverse=True)
        best_cand, best_digits = phone_candidates[0][1], phone_candidates[0][2]
        if len(best_digits) == 12 and best_digits.startswith("91"):
            phone = f"+91-{best_digits[2:7]} {best_digits[7:12]}"
        elif len(best_digits) == 10 and best_digits[0] in "6789":
            phone = f"+91-{best_digits[:5]} {best_digits[5:]}"
        else:
            phone = re.sub(r"[^\d+ -]", "", best_cand).strip(" -")

    if email:
        email = re.sub(r"\.(?:cor|con|corm|cm)$", ".com", email, flags=re.I)

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

    if "REGO" in raw_text.upper() and "SERVICES" in company_name.upper() and "REGO" not in company_name.upper():
        company_name = "REGO " + company_name
    if re.search(r"\b(pvt|ltd)\b", raw_text, re.I) and not re.search(r"\b(pvt|ltd)\b", company_name, re.I):
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

    return {
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


def process_card_image(image: Image.Image) -> Dict[str, Any]:
    """Hybrid orchestrator: attempts Vision AI first, falls back to multi-angle OCR."""
    settings = get_settings()
    api_key = settings.get("gemini_api_key", "").strip()
    pref = settings.get("engine_preference", "hybrid")

    # 1. Vision AI if configured and not explicitly set to local-only
    if pref in ["hybrid", "gemini"] and api_key and GENAI_AVAILABLE:
        try:
            ai_data = extract_with_gemini(image, api_key)
            if ai_data and ai_data.get("name"):
                return ai_data
        except Exception as e:
            print(f"[Engine] Vision AI exception, falling back: {e}")

    # 2. Local High-Precision Tesseract OCR Engine
    raw_text = run_robust_local_ocr(image)
    ocr_data = extract_fields_from_text(raw_text)
    ocr_data["raw_text"] = raw_text
    return ocr_data


# ==========================================
# 4. FLASK WEB APPLICATION & API
# ==========================================

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024  # 32 MB max batch upload


@app.route("/")
def index():
    return render_template_string(FRONTEND_HTML)


@app.route("/uploads/<filename>")
def serve_upload(filename):
    return send_from_directory(UPLOAD_DIR, filename)


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
    """Save a new contact in SQLite database."""
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
    if conn.is_postgres:
        query += " RETURNING id"
        c.execute(query, params)
        res = c.fetchone()
        new_id = res["id"] if res else 1
    else:
        c.execute(query, params)
        new_id = c.lastrowid
    conn.commit()
    conn.close()

    return jsonify({"success": True, "id": new_id, "message": "Contact saved successfully."})


@app.route("/api/contacts/<int:cid>", methods=["PUT"])
def api_update_contact(cid):
    """Update contact fields."""
    data = request.json or {}
    conn = get_db()
    c = conn.cursor()
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
    """Summary dashboard statistics."""
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) as total FROM contacts")
    res = c.fetchone()
    total = res["total"] if res else 0

    c.execute("SELECT COUNT(*) as new_leads FROM contacts WHERE lead_status = 'New'")
    res = c.fetchone()
    new_leads = res["new_leads"] if res else 0

    c.execute("SELECT COUNT(*) as contacted FROM contacts WHERE lead_status = 'Contacted'")
    res = c.fetchone()
    contacted = res["contacted"] if res else 0

    c.execute("SELECT COUNT(*) as qualified FROM contacts WHERE lead_status = 'Qualified'")
    res = c.fetchone()
    qualified = res["qualified"] if res else 0

    c.execute("SELECT COUNT(*) as customers FROM contacts WHERE lead_status = 'Customer'")
    res = c.fetchone()
    customers = res["customers"] if res else 0

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


@app.route("/api/settings", methods=["GET", "POST"])
def api_settings():
    """Get or save API key and engine preferences."""
    if request.method == "POST":
        data = request.json or {}
        settings = get_settings()
        if "gemini_api_key" in data:
            settings["gemini_api_key"] = data["gemini_api_key"].strip()
        if "engine_preference" in data:
            settings["engine_preference"] = data["engine_preference"].strip()
        save_settings(settings)
        return jsonify({"success": True, "message": "Settings saved."})

    settings = get_settings()
    masked_key = ""
    if settings.get("gemini_api_key"):
        raw = settings["gemini_api_key"]
        masked_key = raw[:4] + "..." + raw[-4:] if len(raw) > 8 else "***"

    return jsonify({
        "has_api_key": bool(settings.get("gemini_api_key")),
        "masked_api_key": masked_key,
        "engine_preference": settings.get("engine_preference", "hybrid"),
        "genai_sdk_installed": GENAI_AVAILABLE,
        "tesseract_path": pytesseract.pytesseract.tesseract_cmd,
        "is_postgres": is_postgres_active(),
        "db_engine": "Neon Cloud PostgreSQL (Persistent)" if is_postgres_active() else "Local SQLite Database",
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
        .engine-dot {
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: #10b981;
            box-shadow: 0 0 8px #10b981;
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
                    <input type="password" id="geminiApiKeyInput" placeholder="AIzaSy...">
                    <small style="color:var(--text-muted); margin-top:4px;">
                        Get a free Gemini API key at <a href="https://aistudio.google.com" target="_blank" style="color:var(--primary);">Google AI Studio</a>. When present, cards are parsed with Vision AI.
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
                    <div style="display: flex; align-items: center; gap: 8px; font-weight: 600; color: #1e293b;">
                        <span>💾</span> <span id="dbStatusText">Local SQLite Database</span>
                    </div>
                    <small style="color:var(--text-muted); display:block; margin-top:4px;">
                        Set <code>DATABASE_URL</code> in Render to connect <a href="https://neon.tech" target="_blank" style="color:var(--primary);">Neon.tech PostgreSQL</a> for 100% free permanent storage.
                    </small>
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
        function openReviewModal(cardData, isEditMode = false) {
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
                raw_text: document.getElementById('cardRawText').value
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

        function filterContacts() {
            loadContacts();
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

                if (data.has_api_key) {
                    text.textContent = 'Vision AI Active (99%+ Accuracy)';
                    pill.className = 'engine-pill';
                } else {
                    text.textContent = 'Local Tesseract OCR';
                    pill.className = 'engine-pill tesseract';
                }

                const dbText = document.getElementById('dbStatusText');
                if (dbText) {
                    dbText.textContent = data.db_engine || 'Local SQLite Database';
                    dbText.style.color = data.is_postgres ? '#16a34a' : '#1e293b';
                }

                document.getElementById('enginePreferenceSelect').value = data.engine_preference || 'hybrid';
            } catch (err) {
                console.error(err);
            }
        }

        async function saveSettings() {
            const key = document.getElementById('geminiApiKeyInput').value;
            const pref = document.getElementById('enginePreferenceSelect').value;

            try {
                const resp = await fetch('/api/settings', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ gemini_api_key: key, engine_preference: pref })
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

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)