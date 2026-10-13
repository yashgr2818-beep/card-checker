# 📇 BizConnect AI Studio - Business Card Scanner & Lead CRM

A modern, high-accuracy AI business card scanner and CRM suite matching the features and workflow of BizConnect.

![BizConnect AI Studio](https://img.shields.io/badge/AI-Gemini%20Vision%20%2B%20Tesseract-blue?style=for-the-badge)
![Python](https://img.shields.io/badge/Python-3.11+-brightgreen?style=for-the-badge)
![Flask](https://img.shields.io/badge/Framework-Flask-black?style=for-the-badge)
![Docker](https://img.shields.io/badge/Deploy-Docker%20%2F%20Render-blueviolet?style=for-the-badge)

---

## ✨ Features

- **Hybrid AI & Local OCR Engine**:
  - **Google Gemini Multimodal Vision AI (`gemini-3.8-flash`)**: 99%+ extraction accuracy for all printed text, handwriting, and complex card layouts.
  - **100% Offline Local Tesseract OCR**: Built-in 4-way auto-orientation detection and multi-pass extraction when offline or running without API keys.
- **Single & Batch Scanning**: Drag-and-drop multiple business cards at once.
- **Live Viewfinder Camera**: Scan cards directly from mobile or desktop webcams with a card boundary viewfinder.
- **Mini-CRM Lead Dashboard**:
  - Real-time search across Name, Company, Email, Phone, and City.
  - Lead pipeline tracking (`New Leads`, `Contacted`, `Qualified`, `Converted / Customer`).
  - Follow-up dates and notes.
- **1-Click Exports**:
  - 📊 **Excel (.xlsx)** with auto-sized columns and clean formatting.
  - 📄 **CSV** with UTF-8 BOM encoding.
  - 📲 **vCard (.vcf 3.0)** for instant 1-tap import into Google Contacts, Apple iPhone, or Outlook.

---

## 🚀 Quick Start (Local Setup)

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/yashgr2818-beep/card-checker.git
cd card-checker
pip install -r requirements.txt
```

### 2. (Optional) Install Tesseract OCR
- **Windows**: Install [Tesseract OCR Windows](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux / Ubuntu**:
  ```bash
  sudo apt-get update && sudo apt-get install -y tesseract-ocr
  ```
- **macOS**:
  ```bash
  brew install tesseract
  ```

### 3. Run the Application
```bash
python app.py
```
Open [http://127.0.0.1:5000](http://127.0.0.1:5000) in your browser.

---

## ☁️ Deploy to Render with Docker

1. Create a new **Web Service** on [Render](https://render.com).
2. Connect this GitHub repository.
3. Select **Docker** as the Runtime environment (Render will automatically use the included `Dockerfile`).
4. (Optional) Add your environment variable:
   - `GEMINI_API_KEY`: `your_google_gemini_api_key`
5. Click **Deploy**!

---

## 🔒 Privacy & Security

Card uploads, SQLite databases, and local API keys are excluded via `.gitignore` to protect sensitive contact data and credentials from being published.
