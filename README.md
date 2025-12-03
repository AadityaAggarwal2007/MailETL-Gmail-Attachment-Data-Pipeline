 # MailETL – Gmail Attachment Data Pipeline

This project automates a very common boring task:

> Take Excel/CSV attachments from Gmail and turn them into a clean, refreshable Excel table using Google Apps Script + Power Query.

---

## 1. Enable IMAP in Gmail

Before connecting anything, IMAP must be enabled.

1. Open Gmail  
2. Go to Settings → See all settings  
3. Open the “Forwarding and POP/IMAP” tab  
4. Under “IMAP access”, select **Enable IMAP**  
5. Save the changes

---

## 2. Create Google Apps Script (Gmail → JSON API)

1. Open Google Apps Script  
2. Click **New Project**  
3. Paste the script code  
4. Click **Deploy** → **New deployment**  
5. Under “Select type”, choose **Web app**  
6. Set “Who has access?” → **Anyone**  
7. Click **Deploy**, then authorize the permissions  
8. Copy the **Web App URL**  
   (Power Query will use this URL as an API endpoint)

---

## 3. Connect Power Query to Gmail Data

### Step A: Import the Web App URL  
Excel → **Data** → **From Web**  
Paste the Web App URL.

### Step B: Convert the result into a table  
Power Query → **To Table**

### Step C: Expand the data  
Use the expand icon to bring all fields into the table.

### Step D: Filter relevant emails  
Filter using Subject, From, or MIME type.  
Remove unnecessary columns.

### Step E: Decode Base64 attachments  
Add a custom column:

Binary.FromText([Binary], BinaryEncoding.Base64)

sql
Copy code

### Step F: Convert the binary into a CSV table  
Add another custom column:

Csv.Document([Content])

yaml
Copy code

### Step G: Expand the CSV  
Expand the generated table to view actual CSV data.

### Step H: Clean and transform  
Rename columns, change data types, remove nulls, and complete your transformations.

---

## Result

You now have a refreshable pipeline that pulls emails and CSV attachments straight from Gmail into Excel via Power Query.

Refreshing the workbook updates everything automatically whenever new emails arrive.

---
<img width="1536" height="1024" alt="image" src="https://github.com/user-attachments/assets/86b66d8e-d2be-4788-ae53-f36b9fa69fd6" />

<img width="2940" height="956" alt="image" src="https://github.com/user-attachments/assets/8e80cae5-1b71-4882-83ee-380cba14befb" />

<img width="2940" height="1782" alt="image" src="https://github.com/user-attachments/assets/92e8c8b4-6920-41f9-8200-a262b637def6" />

<img width="1470" height="891" alt="Screenshot 2025-12-02 at 6 36 34 PM" src="https://github.com/user-attachments/assets/cc337453-08d2-4eb6-b8f7-4dde28ae8efb" />
<img width="1676" height="944" alt="image" src="https://github.com/user-attachments/assets/3c626a18-8130-4729-96ce-37edd639fa25" />

