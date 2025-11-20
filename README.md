# Mass-Communication-Macros
A VBA-based automation system that sends mass communication emails through Outlook using batches, and scans for undeliverable/bounce messages to update Excel statuses automatically.

# Mass Communication Email Automation (Excel + Outlook VBA)

Automates large-scale email communication using Excel and Outlook, including batching recipients, drafting bulk messages using `.msg` templates, and automatically detecting undelivered (bounce) emails to update status back in Excel.

This is a sanitized demo version—no internal or sensitive data included.

---

## 🚀 Purpose

This tool was designed to streamline bulk outreach workflows by reducing manual effort in email drafting and status tracking. It is useful for:

- Vendor / stakeholder communication
- Bulk announcements and policy updates
- Large distribution lists where bounce tracking is required

---

## ✨ Key Features

### **📩 1. Bulk Email Sender (`SendBulkEmailOnly`)**
- Pulls recipient emails from Excel (Column A)
- Sends in configurable BCC batches (e.g., 50 / 100 / 200)
- Uses a chosen Outlook `.msg` template
- Creates drafts for manual review before sending
- Updates status column:  
  `Drafted (timestamp)`

### **📬 2. Bounce Detection (`CheckUndeliveredForSentBatch`)**
- Scans Outlook Inbox for bounce / NDR messages
- Detects common failure subjects:
  - *"Undeliverable"*
  - *"Delivery has failed"*
  - *"Returned mail"*
- Extracts failing address using Regex
- Marks matching rows in Excel as:  
  `UNDELIVERED (timestamp)`
- Handles duplicates and missing records gracefully

---

## 📂 Folder Structure

📁 Mass-Communication-Macros
│── SendBulkEmailOnly.bas
│── CheckUndeliveredForSentBatch.bas
│── README.md

---

## 🛠 Setup & Requirements

### **Prerequisites**

| Component | Required |
|----------|----------|
| Microsoft Excel | ✔ |
| Microsoft Outlook (Desktop) | ✔ |
| `.msg` email template | ✔ |
| Macros enabled | ✔ |

### **Excel Format**

| Column | Purpose |
|--------|---------|
| A | Recipient email ID |
| B | Status (auto-updated) |

Example:

| Email | Status |
|--------|--------|
| vendor1@example.com | |
| vendor2@example.com | |

---

## ▶️ How to Use

### **🔹 Step 1: Run Bulk Sender**

1. Open Excel
2. Press **`ALT + F8`**
3. Run: **`SendBulkEmailOnly`**
4. Enter batch size (e.g., `100`)
5. Choose `.msg` template
6. Draft emails appear in Outlook → You send manually

---

### **🔹 Step 2: Run Bounce Checker**

1. Press **`ALT + F8`**
2. Run: **`CheckUndeliveredForSentBatch`**
3. Script scans Inbox
4. Status column updates automatically

---

## 🧠 Example `.msg` Template

Subject: Important Update – Please Review

Hi,

This is a general informational update. No immediate action is required unless specified.

Regards,
Communications Team

---

## 🔧 Skills Demonstrated

- Outlook Object Model Automation
- Excel VBA (Late Binding / Dictionaries)
- Regex-based parsing
- Batch processing logic
- Email workflow automation
- Data-cleaning + logging logic

---



