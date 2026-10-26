# 📇 BizConnect AI Studio - Business Card Scanner & Lead CRM

A modern, high-accuracy AI business card scanner and CRM suite matching the features and workflow of BizConnect.

![BizConnect AI Studio](https://img.shields.io/badge/AI-Gemini%20Vision%20%2B%20Tesseract-blue?style=for-the-badge)
![Database](https://img.shields.io/badge/Database-Neon%20PostgreSQL-green?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.11+-brightgreen?style=for-the-badge)
![Flask](https://img.shields.io/badge/Framework-Flask-black?style=for-the-badge)
![Executable](https://img.shields.io/badge/Windows-Standalone%20EXE-orange?style=for-the-badge)
![Docker](https://img.shields.io/badge/Deploy-Docker%20%2F%20Render-blueviolet?style=for-the-badge)

---

## ✨ Features

- **Hybrid AI & Local OCR Engine**:
  - **Google Gemini Multimodal Vision AI**: High-accuracy cloud AI extraction with continuous self-correcting learning memory.
  - **100% Offline Local Tesseract OCR**: Built-in 4-way auto-orientation detection and multi-pass extraction when running offline or without an API key.
- **Real-Time Internet Connectivity Monitor**:
  - Continuous background network probe measuring live connection latency (ms).
  - Instant offline warning banner and automatic fallback to Local OCR when Wi-Fi/Internet drops.
- **Neon Cloud PostgreSQL Database**:
  - Enterprise connection pooling (`ThreadedConnectionPool`) for fast query execution and zero local file locks.
  - Automatic schema initialization and indexing.
- **Self-Learning AI Memory**:
  - Learns from every manual edit or correction made in the review modal, ensuring typos or uncommon domains are never misread twice.
- **Single & Batch Scanning**: Drag-and-drop multiple business cards at once.
- **Live Viewfinder Camera**: Scan cards directly from mobile or desktop webcams with a card boundary viewfinder.
- **Mini-CRM Lead Dashboard**:
  - Real-time search across Name, Company, Email, Phone, and City.
  - Lead pipeline tracking (`New Leads`, `Contacted`, `Qualified`, `Converted / Customer`).
  - Follow-up dates and notes.
- **1-Click Export**:
  - 📥 **CSV (.csv)** with UTF-8 BOM encoding for seamless opening in Excel, Google Sheets, or CRM systems.
- **Standalone Windows Executable**:
  - Pre-compiled standalone `CardScanner.exe` bundled inside `dist/CardScanner/`—runs on any Windows computer **without installing Python**.

---

## 🚀 Quick Start (Local Setup)

### Option A: 1-Click Launchers (Easiest for Windows)
- **Zero Python Installation Needed:** Double-click [`run_exe.bat`](run_exe.bat) to launch the standalone `CardScanner.exe`.
- **Python Development Mode:** Double-click [`run_local.bat`](run_local.bat) to auto-verify dependencies and start `app.py`.

### Option B: Manual Terminal Setup

1. **Clone & Install Dependencies**
```bash
git clone https://github.com/yashgr2818-beep/card-checker.git
cd card-checker
pip install -r requirements.txt
```

2. **(Optional) Install Tesseract OCR**
- **Windows**: Install [Tesseract OCR Windows](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux / Ubuntu**:
  ```bash
  sudo apt-get update && sudo apt-get install -y tesseract-ocr
  ```
- **macOS**:
  ```bash
  brew install tesseract
  ```

3. **Run the Application**
```bash
python app.py
```
Open [http://127.0.0.1:5000](http://127.0.0.1:5000) in your browser.

---

## 📦 Building the Standalone Executable

To re-compile the standalone `.exe` package after code changes:
1. Simply double-click [`build_exe.bat`](build_exe.bat).
2. It compiles `app.py` into `dist/CardScanner/CardScanner.exe` and automatically cleans up temporary build files.
3. Zip and share the `dist/CardScanner` folder with any user.

---

## ☁️ Deploy to Render with Docker

1. Create a new **Web Service** on [Render](https://render.com).
2. Connect this GitHub repository.
3. Select **Docker** as the Runtime environment (Render will automatically use the included `Dockerfile`).
4. Add your environment variables:
   - `DATABASE_URL`: `your_neon_postgres_url`
   - `GEMINI_API_KEY`: `your_google_gemini_api_key`
5. Click **Deploy**!

---

## 🔒 Privacy & Security

Card uploads, environment secrets (`.env`), database URLs, and API keys are excluded via `.gitignore` to protect sensitive contact data and credentials from being published.
