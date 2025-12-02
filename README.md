# MailETL – Gmail Attachment Data Pipeline

This project automates a very common boring task:

> Take Excel/CSV attachments from Gmail and turn them into a clean, refreshable Excel table using Google Apps Script + Power Query.

You:
- Send / receive sales files in Gmail  
- Apps Script exposes those attachments as JSON via a Web App URL  
- Power Query calls that URL, converts the Base64 **Binary** back to file content, and combines everything into a single, cleaned table in Excel  

Perfect for recurring reports where the data always lands in your inbox.

---

## 1. What this project does

- Reads recent emails from Gmail (e.g. “sales data chennai”, “sales data delhi” etc.)
- Extracts attachment metadata:
  - Email date  
  - From  
  - Subject  
  - File name  
  - MIME type  
  - Attachment as Base64 string (`Binary`)
- Exposes all of this as a JSON endpoint via Google Apps Script
- Uses Power Query in Excel to:
  - Call the Web App URL  
  - Convert Base64 text back to real binary using  
    `Binary.FromText([Binary], BinaryEncoding.Base64)`
  - Turn the binary into a table using `Csv.Document` (or Excel.Workbook)
  - Expand and clean the data into a final fact table

---

## 2. Tech stack

- **Gmail** – source of emails + attachments  
- **Google Apps Script** – bridges Gmail → JSON API  
- **Excel Power Query** – pulls data from the API, decodes, transforms and cleans  

---

## 3. Setup

### 3.1 Enable IMAP in Gmail

1. Open **Gmail**  
2. Go to **Settings → See all settings**  
3. Open the **Forwarding and POP/IMAP** tab  
4. Under **IMAP access**, select **Enable IMAP**  
5. Click **Save changes**

*(This doesn’t send anything out; it just allows programmatic access.)*

---

### 3.2 Create the Google Apps Script Web App

1. Go to **script.google.com** and click **New project**  
2. Replace the default code with the script from this repo (`Code.gs`)  
3. Click **Deploy → New deployment**  
4. Under **Select type**, choose **Web app**  
5. Set **Who has access?** to something appropriate:
   - For testing: **Anyone**  
   - For real usage: at least **Anyone with Google account** or restricted to your account
6. Click **Deploy**  
7. Approve the permissions  
8. Copy the **Web App URL** – you’ll use this inside Power Query

> ⚠️ When you’re done testing, you can **disable or restrict** this deployment so the endpoint isn’t public.

---

### 3.3 Power Query – connect to Gmail data

1. Open **Excel → Data → From Web**  
2. Paste your **Web App URL** from Apps Script  
3. In the Navigator, select the JSON list and:
   1. Turn it into a **Table**  
   2. **Expand** the record so you see columns like `EmailDate`, `From`, `Subject`, `FileName`, `Binary`, etc.  
4. **Filter** rows:
   - Use `Subject` (e.g. contains “sales data”)  
   - Remove unnecessary columns if needed  
5. Add a **Custom Column**:

   ```m
   = Binary.FromText([Binary], BinaryEncoding.Base64)
This converts the Base64 text back into binary file content.

Add another Custom Column using Csv.Document (or Excel.Workbook if your files are .xlsx), for example:

m
Copy code
= Csv.Document([Custom])
Expand that new table column to get your actual data columns

Rename columns, set data types, and do your usual cleaning steps

Load the final query into a worksheet as a table

At this point, the pipeline is wired up.

4. How to Use (after setup)
Once everything above is configured, day-to-day usage is simple:

Send / receive new data files

Keep using the same Gmail address

Use similar subjects (e.g. “sales data nagpur”, “sales data delhi”)

Attach CSV/Excel files in the same structure you designed the query for

Refresh the pipeline

Open the Excel file that contains this query

Go to Data → Refresh All

Power Query will:

Call the Apps Script Web App

Pick up the latest email attachments

Decode them back to binary

Rebuild and clean the combined data table

Use the refreshed table

Build pivot tables, charts, dashboards, or whatever analysis you want on top of the refreshed data

Every time new files land in Gmail, just Refresh All again

5. Files in this repo
Code.gs – Google Apps Script that reads Gmail and returns JSON

MailETL-Gmail-Attachment-Data-Pipeline.xlsx (or similar) – Excel file with Power Query steps

screenshots/ – Pipeline screenshots (Gmail, Apps Script, Power Query, final Excel table)

README.md – This documentation

6. Privacy & safety checklist (before uploading)
If you publish this project publicly:

Remove real email addresses, names, or any sensitive data from screenshots

Use dummy sample data in example files instead of live production files

Don’t commit your real Web App URL if the deployment is still open to “Anyone”

Consider restricting or disabling the Web App deployment after testing

7. Ideas for future improvements
Add label-based filtering (only a specific Gmail label is read)

Support multiple attachment types (CSV + XLSX) automatically

Add a date range filter in the Apps Script API

Build a Power BI / Excel dashboard on top of the final table

Built to turn “messy Gmail attachments” into a repeatable data pipeline with almost zero manual work.
