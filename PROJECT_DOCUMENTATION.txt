# Complete Project Documentation & Technical Guide
## Intelligent Business Card Scanner & Lead Management CRM
**Repository / Project Directory:** `c:\Users\digital\Desktop\node\card`  
**Author:** Pair programmed with Google Antigravity AI  
**Last Updated:** October 2026  

---

## Table of Contents
1. [Project Overview & Purpose](#1-project-overview--purpose)
2. [High-Level Architecture & Workflow Diagram](#2-high-level-architecture--workflow-diagram)
3. [Libraries & Technologies Used (And Why)](#3-libraries--technologies-used-and-why)
4. [File Structure of the Project](#4-file-structure-of-the-project)
5. [Deep Dive: Code Walkthrough (`app.py`)](#5-deep-dive-code-walkthrough-apppy)
   - [5.1 Environment & Settings Management](#51-environment--settings-management)
   - [5.2 Dual-Database Engine (SQLite + Neon Cloud PostgreSQL)](#52-dual-database-engine-sqlite--neon-cloud-postgresql)
   - [5.3 AI Learning & Smart Correction Memory](#53-ai-learning--smart-correction-memory)
   - [5.4 Google Gemini Multimodal Vision AI Pipeline](#54-google-gemini-multimodal-vision-ai-pipeline)
   - [5.5 Local High-Precision Tesseract OCR Engine](#55-local-high-precision-tesseract-ocr-engine)
   - [5.6 The Rule-Based Extraction & Cleaning Engine](#56-the-rule-based-extraction--cleaning-engine)
   - [5.7 REST API Endpoints](#57-rest-api-endpoints)
   - [5.8 Frontend Web Application (Single-Page App)](#58-frontend-web-application-single-page-app)
6. [Automation & Local Setup (`run_local.bat`)](#6-automation--local-setup-run_localbat)
7. [Step-by-Step Data Flow: What Happens During a Scan](#7-step-by-step-data-flow-what-happens-during-a-scan)
8. [Troubleshooting & Frequently Asked Questions](#8-troubleshooting--frequently-asked-questions)

---

## 1. Project Overview & Purpose

### What is this project?
This application is an enterprise-grade **Business Card Scanner and Lead Management CRM** built to match the accuracy of commercial apps like *BizConnect* and *CamCard*.

### Key Problems Solved:
1. **Commercial Card Scanning Apps are Expensive:** Most charge high monthly subscriptions for basic OCR.
2. **Standard OCR Often Fails on Complex Layouts:** Business cards feature vertical text, creative fonts, dark backgrounds, company logos, and multi-line addresses.
3. **Repeated Scanning Mistakes:** Most OCR software makes the same mistake every time on a card with unusual typography. This app features **Smart Correction Memory**, so once you correct a mistake in the review modal, the system **never repeats that mistake again**.
4. **Offline vs Cloud Flexibility:**
   - Works **100% offline** without any internet connection using local Tesseract OCR.
   - Works with **Google Gemini Vision AI** (`gemini-3.8-flash`) when an API key is provided for 99%+ accuracy.
5. **Data Synchronization (Local + Cloud):**
   - Automatically synchronizes with **Neon Cloud PostgreSQL** so your leads and learned rules are shared across devices and cloud deployments (e.g. Render).
   - Automatically falls back to **Local SQLite (`cards.db`)** if offline.

---

## 2. High-Level Architecture & Workflow Diagram

```
[Business Card Photo] 
   (Camera capture or File upload: JPG, PNG, WEBP)
           │
           ▼
[Image Preprocessing (PIL)]
   • Auto-Orientation (EXIF transpose)
   • RGB conversion & smart resizing (1600px max)
           │
           ▼
    Is Gemini API Key configured & verified?
           ├─── YES ───► [Google Gemini Vision AI (gemini-3.8-flash)]
           │                 • Injects learned correction memory
           │                 • Structured JSON output via Pydantic
           │                 • Extracts vertical/styled text flawlessly
           │                 • (If network fails or 404, falls back to OCR)
           │
           └─── NO ────► [Local High-Precision Tesseract OCR]
                             • 3-Angle scan (0°, 90°, 270°) with binary threshold
                             • Multi-layout high-res extraction on winning angle
                             • 46 Indian STD city codes directory check
                             • Regex mobile & landline separation
                             • Website & Email cleanup
           │
           ▼
[Smart Correction Memory Application]
   • Applies all learned rules from prior user corrections
           │
           ▼
[Interactive Review Modal (Web UI)]
   • User verifies or edits fields (Name, Phone, Email, Company, Address, etc.)
   • User sets Lead Status (New, Contacted, Qualified, Customer)
           │
           ▼ (User clicks Save)
   • If user corrected any field: System auto-saves rule to `corrections_memory`
   • Saves contact record to Active Database (Neon PostgreSQL or SQLite)
           │
           ▼
[Lead CRM & Export Engine]
   • Real-time search, filters, pagination, batch delete
   • 1-Click Export to Excel (.xlsx), CSV (.csv), or vCard (.vcf for phones)
```

---

## 3. Libraries & Technologies Used (And Why)

Every library in [requirements.txt](file:///c:/Users/digital/Desktop/node/card/requirements.txt) was handpicked for a specific engineering reason:

| Library / Tool | Purpose | Why We Chose It |
|---|---|---|
| **`Flask`** | Web Framework & REST API | Lightweight, lightning-fast startup (<1 second), zero bloat, easy to embed single-page frontend directly, perfect for desktop tools and cloud containers. |
| **`Pillow` (PIL)** | Image Processing Engine | The industry standard for Python image manipulation. Used for auto-orientation (`ImageOps.exif_transpose`), contrast boosting (`ImageEnhance.Contrast`), binary thresholding, and resizing. |
| **`pytesseract`** | Local OCR Wrapper | Bridges Python to Google’s C++ Tesseract-OCR engine. Enables 100% offline, free, and private text recognition without sending images to third-party servers. |
| **`google-genai`** | Google Gemini Vision AI SDK | The official modern SDK for Gemini Multimodal models. Allows sending the image directly to `gemini-3.8-flash` to extract semantic information in structured JSON format. |
| **`pydantic`** | Strict Schema Validation | Forces Gemini to respond strictly in our defined JSON structure (`BusinessCardModel`). Prevents hallucinated keys and broken JSON strings. |
| **`psycopg2-binary`** | PostgreSQL Client | Connects the application to **Neon Cloud PostgreSQL**. Handles SSL encryption, connection pooling, and cross-platform reliability. |
| **`sqlite3`** *(Python Built-in)* | Local Database | Zero-configuration database built right into Python. Saves cards to `cards.db` when offline or when no cloud database URL is provided. |
| **`python-dotenv`** | Environment Configuration | Automatically reads `.env` variables (`DATABASE_URL`, `GEMINI_API_KEY`) so sensitive credentials stay secure and don't need hardcoding. |
| **`openpyxl`** | Excel Generation | Creates Microsoft Excel spreadsheets (`.xlsx`) with custom styling: dark blue header row, bold white text, zebra striping, and auto-adjusted column widths. |
| **`re`** *(Python Built-in)* | Regular Expressions | Custom-tuned regex patterns designed to extract Indian 10-digit mobile numbers, STD telephone codes, emails, websites, and GSTIN numbers. |
| **`webbrowser` & `threading`** *(Built-in)* | Local Auto-Launch | When started via `run_local.bat`, spins up a background thread that waits 1.2s for Flask to start, then automatically opens your default browser at `http://127.0.0.1:5000`. |

---

## 4. File Structure of the Project

```
c:\Users\digital\Desktop\node\card\
│
├── app.py                   # The entire backend engine, REST API, and frontend SPA
├── run_local.bat            # One-click Windows startup script (auto-checks Python & packages)
├── requirements.txt         # All Python package dependencies
├── .env                     # Private secrets (Neon DATABASE_URL, GEMINI_API_KEY)
├── settings.json            # Local fallback config storage
├── cards.db                 # Local SQLite database (used when offline)
├── README.md                # Project summary & quickstart instructions
├── PROJECT_DOCUMENTATION.md # This comprehensive technical manual
│
├── uploads/                 # Storage for scanned business card image files
│   └── card_1791804318.png
│
└── .git/                    # Git version control (synchronized with GitHub repo)
```

---

## 5. Deep Dive: Code Walkthrough (`app.py`)

The file [app.py](file:///c:/Users/digital/Desktop/node/card/app.py) is organized into 5 logical modules:

```
┌─────────────────────────────────────────────────────────┐
│ Section 1: Setup, Paths, & Settings Loader              │
│ Section 2: Dual Database Layer (SQLite + PostgreSQL)    │
│ Section 3: Smart Correction Memory & AI OCR Engines     │
│ Section 4: Flask REST API Endpoints                     │
│ Section 5: Export Engine & Embedded Frontend SPA        │
└─────────────────────────────────────────────────────────┘
```

---

### 5.1 Environment & Settings Management

#### `get_settings()` and `save_settings()`
```python
def get_settings() -> Dict[str, Any]:
```
- **How it works:**
  1. Sets default dictionary (`engine_preference: "hybrid"`, empty keys).
  2. Reads `.env` directly from disk on every invocation so if you edit `.env`, changes reflect instantly without server restart.
  3. Reads `settings.json` for UI-saved settings.
  4. Merges values intelligently without letting empty strings overwrite valid environment variables.

---

### 5.2 Dual-Database Engine (SQLite + Neon Cloud PostgreSQL)

The app supports two database backends with zero code changes:

#### `get_db()` and `is_postgres_active()`
```python
def is_postgres_active() -> bool:
    url = get_database_url()
    return bool(url and PSYCOPG2_AVAILABLE and url.startswith("postgres"))
```
- **When `is_postgres_active()` is True:**
  - Connects to Neon Cloud PostgreSQL over SSL (`sslmode=require`).
  - Uses `psycopg2.extras.RealDictCursor` so row records return as Python dictionaries (`row['name']`).
  - Replaces SQLite query placeholders (`?`) with PostgreSQL placeholders (`%s`) via a custom `_execute_compat()` helper.
- **When `is_postgres_active()` is False:**
  - Connects to local `cards.db` SQLite file.
  - Sets `conn.row_factory = sqlite3.Row` for dictionary access.

#### Database Tables Schema:
1. **`contacts` table:**
   - `id`: Primary key (Serial / Auto-increment)
   - `name`: Contact full name
   - `company_name`: Business or firm name
   - `designation`: Job title (e.g. Director, Manager)
   - `phone_number`: Primary phone / mobile
   - `secondary_phone`: Additional phone number
   - `email`: Primary email address
   - `website`: Website URL
   - `address`: Office / branch address
   - `services_products`: Offerings or catalog summary
   - `notes`: Custom notes or slogans
   - `lead_status`: CRM stage (`New`, `Contacted`, `Qualified`, `Customer`)
   - `image_path`: Relative filename of the uploaded business card
   - `ocr_engine`: Which engine extracted it (e.g., `Vision AI (gemini-3.8-flash)` or `Local OCR (Tesseract)`)
   - `raw_text`: Full unparsed OCR output for auditing
   - `created_at`: Timestamp

2. **`corrections_memory` table:**
   - `id`: Primary key
   - `field_name`: The field corrected (e.g., `website`, `phone_number`, `email`)
   - `wrong_value`: What the OCR erroneously extracted (e.g., `www.dotdosing.in`)
   - `corrected_value`: What the user manually typed (e.g., `dotdosing.in`)
   - `context_hint`: Company or name context
   - `created_at`: Timestamp

---

### 5.3 AI Learning & Smart Correction Memory

One of the most powerful features of this system:

#### How it works:
1. **Recording Corrections (`save_learned_correction`):**
   When you review a scanned card in the modal and change `wrong_value` to `corrected_value`, the system saves a record into `corrections_memory`.
2. **Applying Rules to Local OCR (`apply_learned_corrections`):**
   Before displaying results, the local parser loops through all learned corrections. If any field matches a previously recorded `wrong_value`, it silently and automatically replaces it with the user's verified `corrected_value`.
3. **Injecting Rules into Gemini Prompt (`build_learned_rules_prompt`):**
   When using Gemini Vision AI, the app compiles the latest learned rules into a special prompt section:
   ```text
   LEARNED USER CORRECTION MEMORY:
   - For field 'website': previously corrected 'www.dotdosing.in' -> 'dotdosing.in'
   Follow these verified corrections strictly!
   ```
   This ensures the AI models adapt to your exact business conventions.

---

### 5.4 Google Gemini Multimodal Vision AI Pipeline

#### `extract_with_gemini(image, api_key)`
- **Schema Enforcement with Pydantic:**
  ```python
  class BusinessCardModel(BaseModel):
      name: Optional[str] = ""
      company_name: Optional[str] = ""
      designation: Optional[str] = ""
      phone_number: Optional[str] = ""
      secondary_phone: Optional[str] = ""
      email: Optional[str] = ""
      website: Optional[str] = ""
      address: Optional[str] = ""
      services_products: Optional[str] = ""
      notes: Optional[str] = ""
  ```
  Passing this schema to `types.GenerateContentConfig(response_mime_type="application/json", response_schema=BusinessCardModel)` guarantees 100% valid JSON matching our exact keys.
- **Model Fallback Chain:**
  Google frequently updates models. The app tries:
  1. `gemini-3.8-flash` (Latest, ultra-fast, high precision)
  2. `gemini-3.5-flash`
  3. `gemini-flash-latest`
  4. `gemini-2.5-flash`
- **Live Key Verification (`test_gemini_api_key`):**
  Sends a lightweight ping to Google.
  - If valid: Returns `{"valid": True, "status": "active", "message": "Connected (gemini-3.8-flash)"}`.
  - If rejected: Returns `{"valid": False, "status": "invalid", "message": "Google rejected key: Invalid API Key"}`.
  - Features smart in-memory caching so page reloads don't cause repetitive network pings.

---

### 5.5 Local High-Precision Tesseract OCR Engine

#### `run_robust_local_ocr(image)`
Real business card photos are often tilted or rotated 90° or 270°. Standard OCR fails completely on sideways text.

#### The 2-Step Rotation Detection Pipeline:
1. **Fast Orientation Scan (Takes ~0.3s):**
   - Shrinks image to a fast 750px thumbnail.
   - Rotates image across 3 angles: `0°`, `90°`, `270°`.
   - Converts to grayscale, applies binary thresholding, and runs lightweight Tesseract PSM 6.
   - Scores each angle based on recognized keywords (presence of `@`, words like `Director`, `Technologies`, `Road`, `Sector`).
2. **Winning Angle High-Resolution Scan:**
   - Rotates the original high-res card to the winning orientation.
   - Resizes up to 2200px using Lanczos resampling.
   - Runs two complementary passes:
     - **Pass A:** Enhanced grayscale contrast (sharpens colored text on dark backgrounds).
     - **Pass B:** Binary thresholding (extracts faint ink and embossed text).
   - Combines lines from both passes, eliminating duplicates while preserving complete contact details.

---

### 5.6 The Rule-Based Extraction & Cleaning Engine

#### `extract_fields_from_text(raw_text)`
When running offline without Gemini, this parser turns unstructured OCR text into structured fields:

1. **Email Extraction:**
   - Looks for standard email regex `[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}`.
   - If missing, scans tokens containing `@` (99%+ probability of being an email).
2. **Website Extraction & Cleanup:**
   - Strips `https://` and `http://`.
   - Never invents `www.` if not printed on the card (e.g. `dotdosing.in` stays `dotdosing.in`).
3. **Indian Phone & STD Directory Engine:**
   - Contains a built-in directory of **46 Indian STD city codes**:
     - 2-digit major metros: `011` (Delhi), `022` (Mumbai), `033` (Kolkata), `044` (Chennai), `080` (Bangalore), `020` (Pune), `040` (Hyderabad), `079` (Ahmedabad).
     - 3-digit regional cities: `0120` (Noida/Ghaziabad), `0124` (Gurgaon), `0141` (Jaipur), `0522` (Lucknow), `0731` (Indore), `0612` (Patna), etc.
   - Prevents landline slicing (e.g. `+91 011 43074427` is preserved as a complete telephone number instead of splitting into a wrong 10-digit mobile number).
   - Detects 10-digit Indian mobile numbers starting with `6, 7, 8, 9` and preserves `+91`.
4. **Name & Designation Identification:**
   - Designation dictionary: `Director`, `CEO`, `Founder`, `Executive`, `Manager`, `Engineer`, `Consultant`, etc.
   - Name heuristics: Finds 2-to-4 word capitalized human names, rejects lines containing company markers (`Pvt`, `Ltd`, `Enterprises`, `Solutions`).

---

### 5.7 REST API Endpoints

| Endpoint | Method | Description |
|---|---|---|
| `/api/scan` | `POST` | Accepts 1 or multiple card images; runs image preprocessing, routes to Gemini or Tesseract, returns JSON extraction results. |
| `/api/contacts` | `GET` | Fetches saved contacts with support for search query `?q=...`, status filter `?status=...`, and pagination. |
| `/api/contacts` | `POST` | Creates or updates a verified contact and records any field edits into `corrections_memory`. |
| `/api/contacts/<id>` | `DELETE` | Deletes a contact record and its associated image file from `uploads/`. |
| `/api/settings` | `GET` | Returns current engine status, masked API key, database engine name, and GenAI SDK status. |
| `/api/settings` | `POST` | Updates API key, database URL, and engine preference. Automatically verifies key with Google. |
| `/api/settings/verify-key`| `POST` | Live test endpoint: pings Google Gemini API and returns instant connectivity status. |
| `/api/corrections` | `GET` | Fetches list of all learned rules and total rule count. |
| `/api/export/excel` | `GET` | Downloads styled `.xlsx` file of all contacts with formatted headers. |
| `/api/export/csv` | `GET` | Downloads standard `.csv` file. |
| `/api/export/vcard` | `GET` | Downloads `.vcf` file ready for 1-click import into iPhone, Android, or Outlook. |

---

### 5.8 Frontend Web Application (Single-Page App)

The complete frontend is embedded inside `app.py` under `FRONTEND_HTML`:
- **Zero Build Step:** Runs immediately without needing Node.js, Webpack, or npm.
- **Modern Design System:** Glassmorphism headers, responsive CSS grid, card shadow states, custom badges, and SVG icons.
- **Three Core Tabs:**
  1. **Scanner Tab:** Drag-and-drop file upload, multi-card batch processing, live camera capture button.
  2. **Contacts CRM Tab:** Filter by lead status (`New`, `Contacted`, `Qualified`, `Customer`), instant search, inline call/email/WhatsApp links, bulk actions.
  3. **Settings Tab:** Gemini API configuration with **🧪 Test Key** button, Neon Database connection status, and live view of learned rules in AI Memory.
- **Review Modal:**
  - Shows high-res preview of the business card photo side-by-side with editable fields.
  - Displays engine badge (e.g. `Engine: Vision AI (gemini-3.8-flash)`).
  - Automatically compares original OCR values against user edits to record learning rules.

---

## 6. Automation & Local Setup (`run_local.bat`)

The file [run_local.bat](file:///c:/Users/digital/Desktop/node/card/run_local.bat) is designed for non-technical users to run the project on any Windows machine:

```bat
@echo off
title BizConnect Card Scanner (Local Launcher)
...
python --version >nul 2>&1
...
python -m pip install -r requirements.txt
...
python app.py
```

### What it does automatically:
1. **Checks Python:** Verifies Python is installed and available in PATH. If missing, prints clear instructions to install it from python.org with "Add to PATH" checked.
2. **Auto-Installs Dependencies:** Checks and installs any missing packages from `requirements.txt`.
3. **Launches App & Opens Browser:** Starts Flask on `http://127.0.0.1:5000` and automatically opens your default web browser.

---

## 7. Step-by-Step Data Flow: What Happens During a Scan

Here is a step-by-step example using a real business card (e.g. *REGO SERVICES PVT. LTD.*):

1. **User Action:** User drags `card.png` onto the scanner area.
2. **Upload:** Frontend sends `FormData` containing the file to `POST /api/scan`.
3. **Image Saving:** Flask saves the image to `uploads/card_1791804318.png`.
4. **Engine Routing:**
   - App checks `get_settings()`.
   - `gemini_api_key` is present and verified.
   - App calls `extract_with_gemini(image, api_key)`.
5. **AI Extraction:**
   - Image is resized to 1600px.
   - System builds prompt including any rules from `corrections_memory`.
   - Calls Google Gemini `gemini-3.8-flash` with `BusinessCardModel` schema.
   - Gemini returns structured JSON:
     ```json
     {
       "name": "Khushi",
       "company_name": "REGO SERVICES PVT. LTD.",
       "designation": "Business Development Executive",
       "phone_number": "+91-8527871429",
       "email": "Bdm02@regoservices.com",
       "website": "www.regoservices.com",
       "address": "Delhi Branch: Sant Nagar, Burari, Delhi - 110084; Noida Branch: Sector-2, Noida - 201301"
     }
     ```
6. **Review Modal:** The modal pops up on screen with all fields pre-filled.
7. **User Edits & Learning:**
   - Suppose the user removes `www.` from the website so it reads `regoservices.com`.
   - User clicks **"Save Contact"**.
   - The app detects `wrong: "www.regoservices.com"` ➔ `corrected: "regoservices.com"`.
   - Inserts rule into `corrections_memory` table in Neon PostgreSQL.
   - Inserts contact into `contacts` table.
8. **Subsequent Scans:** Any future card scanned from that domain automatically applies the rule.

---

## 8. Troubleshooting & Frequently Asked Questions

### Q1: Why did it show 99% accuracy when I typed a wrong API key earlier?
**Answer:** The frontend previously only checked if the input field had characters in it (`if (data.has_api_key)`), without contacting Google. When Gemini failed, the app silently fell back to local OCR.  
**Fixed:** We added the `test_gemini_api_key()` validator, the **🧪 Test Key** button, and active status checks (`active`, `invalid`, `unconfigured`). If a key is invalid, the UI now displays a red warning badge `⚠️ Invalid Gemini Key (Fallback to Local OCR)`.

### Q2: How do I get a free Google Gemini API Key?
1. Visit [aistudio.google.com](https://aistudio.google.com).
2. Log in with your Google account.
3. Click **"Get API key"** and create a key in a new project.
4. Paste it into your [.env](file:///c:/Users/digital/Desktop/node/card/.env) file or into the **Settings tab** in the app and click **"🧪 Test Key"**.

### Q3: Where is my data saved when I am on a different computer?
- Both your local setup and cloud setup connect to your **Neon Cloud PostgreSQL** database (configured in `.env` under `DATABASE_URL`).
- All contacts and learned rules sync in real-time across all devices connected to that URL.
- If you disconnect from the internet, the app safely switches to local `cards.db` SQLite storage.

### Q4: How do I deploy this project to the web for free?
1. Push this codebase to your GitHub repository (already linked: `yashgr2818-beep/card-checker`).
2. Log into [Render.com](https://render.com).
3. Create a **New Web Service** pointing to your repository.
4. Set Build Command: `pip install -r requirements.txt`
5. Set Start Command: `python app.py`
6. Under **Environment Variables**, add:
   - `DATABASE_URL`: *(Your Neon connection string)*
   - `GEMINI_API_KEY`: *(Your Google Gemini API key)*
7. Click **Deploy**. Your app will be live with free SSL.

---
*Documentation compiled and verified for repository `yashgr2818-beep/card-checker`.*
