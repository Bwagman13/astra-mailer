# ✉ Astra Mailer

A cross-platform desktop application for generating personalized one-on-one meeting emails for students.

Upload a spreadsheet of student data, write a boilerplate email template, and let AI personalize each message — then export or draft them in Outlook.

Works on **Windows**, **macOS**, and **Linux**. Outlook draft creation is a Windows-only bonus feature; all other features (generation, export, review) work everywhere.

---

## Quick Start

### 1. Install Python

Download Python 3.10+ from [python.org](https://www.python.org/downloads/).

- **Windows**: Check "Add Python to PATH" during installation
- **Mac**: You can also install via `brew install python`

### 2. Install Dependencies

Open a terminal in the `astra-mailer` folder and run:

```
pip install -r requirements.txt
```

For **Outlook draft creation** (optional, Windows only):
```
pip install pywin32
```

### 3. Set Up Your API Key

Copy `.env.example` to `.env` and add your Anthropic API key:

```
cp .env.example .env
# Then edit .env with your key:
ANTHROPIC_API_KEY=sk-ant-your-actual-key
```

Get a key at [console.anthropic.com](https://console.anthropic.com/).

Alternatively, you can paste the key directly in the app's Setup tab.

### 4. Run the App

```
python astra_mailer.py
```

---

## How to Use

### Tab ① — Setup
- Enter or confirm your API key
- Choose the email tone (Professional / Warm / Concise)
- Write your boilerplate email template
- Set the subject line
- Placeholders like `{name}` and `{meeting_time}` are supported but optional — the AI will personalize regardless

### Tab ② — Data
- Click **Upload Spreadsheet** and select a `.csv` or `.xlsx` file
- The app auto-detects common column names (Name, Email, Meeting Time, etc.)
- Review and adjust the column mapping if needed
- Preview your data in the table

### Tab ③ — Generate
- Review the summary to make sure everything is configured
- Click **Generate All Emails**
- Watch progress in real-time
- Cancel anytime if needed

### Tab ④ — Review & Export
- Click any row to see the full generated email
- **Copy to Clipboard** — copy a single email
- **Export CSV** — all emails in a spreadsheet
- **Export HTML Review** — a printable review page
- **Export Individual TXT** — one file per student
- **Create Outlook Drafts** — (if Outlook is available) creates unsent drafts you can review before sending

---

## Spreadsheet Format

Your spreadsheet should have a header row. The app tries to auto-detect these columns:

| Role | Recognized Headers |
|------|--------------------|
| Name | `Name`, `Student Name`, `Full Name`, `First Name` |
| Email | `Email`, `E-mail`, `Email Address` |
| Meeting Time | `Meeting Time`, `Date`, `Scheduled`, `Appointment` |
| Class | `Class`, `Course`, `Section`, `Subject` |
| Topic | `Topic`, `Agenda`, `Discussion`, `Notes` |
| Advisor | `Advisor`, `Teacher`, `Instructor` |

Additional columns are automatically included as context for personalization.

A sample file (`sample_students.csv`) is included for testing.

---

## Outlook Integration

The app can create **draft emails** in Microsoft Outlook (desktop version, **Windows only**). Drafts are NOT sent automatically — you review and send them yourself.

**Requirements:**
- Microsoft Outlook desktop app must be installed and configured
- Install `pywin32`: `pip install pywin32`

**On Mac/Linux:** The Outlook draft button is hidden. Use the CSV, HTML, or TXT export options instead — these work on all platforms. When Astra runs the app on her Windows PC with Outlook installed, the draft feature will be available automatically.

---

## What Data Is Sent to the AI?

For each student, the following is sent to Anthropic's API:
- Your email template
- The student's mapped fields (name, email, meeting time, class, etc.)
- The selected tone

No data is stored by the API beyond the request. No API key is saved to disk — it stays in memory or your `.env` file.

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "No module named PySide6" | Run `pip install PySide6` |
| "No module named anthropic" | Run `pip install anthropic` |
| API key errors | Check your key at console.anthropic.com |
| Outlook not detected | Install Outlook desktop + `pip install pywin32` |
| CSV encoding issues | Re-save your CSV as UTF-8 in Excel |
| XLSX won't load | Make sure `openpyxl` is installed |

---

## File Structure

```
astra-mailer/
├── astra_mailer.py      # Main application
├── requirements.txt     # Python dependencies
├── .env.example         # API key template
├── .env                 # Your actual API key (create this)
├── sample_students.csv  # Test data
└── README.md            # This file
```

---

## License

Built for Astra. Free to use and modify.
