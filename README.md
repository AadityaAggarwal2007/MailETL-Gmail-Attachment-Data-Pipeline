# Gmail to Power Query Automation (Using Google Apps Script)

This project creates a workflow to fetch Gmail data (including attachments) using Google Apps Script and load it into Excel through Power Query.

---

## 1. Enable IMAP in Gmail

1. Open Gmail  
2. Go to **Settings → See all settings**  
3. Open **Forwarding and POP/IMAP**  
4. Under **IMAP access**, select **Enable IMAP**  
5. Click **Save Changes**

---

## 2. Create Google Apps Script (Gmail → JSON API)

1. Open Google Apps Script  
2. Click **New Project**  
3. Paste your script  
4. Click **Deploy → New deployment**  
5. Under **Select type**, choose **Web app**  
6. Set **Who has access?** → **Anyone**  
7. Click **Deploy**, authorize, and allow  
8. Copy the generated **Web App URL**

Power Query will use this URL to fetch Gmail data.

---

## 3. Load Gmail Data into Power Query

### A. Import the API data  
Excel → **Data → From Web**  
Paste your Web App URL.

### B. Convert to table  
Power Query → **To Table**

### C. Expand the list  
Use the expand icon to bring all message fields (Date, From, Subject, etc.).

### D. Filter the required emails  
Use Subject, From, or MIME type to narrow down the data.  
Remove unnecessary columns.

### E. (Optional) Extract attachment content  
If your script already returns structured data in JSON format  
(e.g., attachment name, type, or decoded content),  
Power Query will automatically display it after expansion.  
No binary conversion or CSV decoding steps are included here.

### F. Clean and shape the data  
Rename columns, remove null records, set correct data types.

---

## Result

A refreshable Gmail → Excel pipeline using Google Apps Script as a bridge.  
Each refresh pulls your latest Gmail data directly into Power Query — ready for analysis, filtering, or reporting.

---
