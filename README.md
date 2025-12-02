MailETL — Gmail Attachment Data Pipeline

A fully automated workflow that extracts Excel attachments directly from Gmail, converts them into structured tables, and loads them into Power Query for analysis — no manual downloads, no dragging, no cleaning.

This project demonstrates real-world ETL design: API creation, Base64 decoding, binary reconstruction, and automated data transformation.

🚀 Why This Project Exists

Gmail does not integrate directly with Power Query or Excel.
Companies receive daily emails with sales files, reports, and logs — and manually downloading them is slow and error-prone.

MailETL solves that.

It turns your Gmail inbox into a data source, builds a custom API around it, and automates the entire pipeline:

Gmail → Google Apps Script API → Power Query → Clean Structured Data

A real, production-style workflow.

📌 Features

Reads email attachments automatically

Converts Base64 strings back into binary files

Parses Excel/CSV content inside Power Query

Fully refreshable pipeline

Zero manual steps once set up

Uses only Google Apps Script + Excel (no servers, no paid tools)

🏗️ Architecture
Incoming Emails (Gmail)
        ↓
Google Apps Script (custom API)
        ↓
JSON Response (metadata + Base64 attachment)
        ↓
Power Query (decode → binary → CSV → clean)
        ↓
Excel / BI-ready dataset

📸 Screenshots
1. Gmail Inbox (Raw Source)

Attachments arriving as data files.
(screenshot placeholder)

2. Google Apps Script API

Converting Gmail attachments into JSON.
(screenshot placeholder)

3. Power Query Transformation

Decoding Base64 → reconstructing binary → expanding.
(screenshot placeholder)

4. Final Output

Clean structured dataset generated automatically.
(screenshot placeholder)

📂 Repository Structure
MailETL/
│
├── README.md
├── HOW_TO_USE.md
│
├── src/
│   └── gmail_script.js
│
├── images/
│   ├── gmail_inbox.png
│   ├── apps_script.png
│   ├── pq_binary.png
│   └── final_output.png
│
└── sample_output/
    └── cleaned_data.xlsx

🧠 What You Learn From This Project

API creation using Google Apps Script

Handling Base64 encoding

Reconstructing files in Power Query

Building a real ETL pipeline

Structuring a professional GitHub project

Turning inbox chaos into clean analytical data

🙌 Contributions & Feedback

This is a personal project built for learning automation and real-world ETL patterns.
Feel free to fork, improve, or reach out with suggestions.
