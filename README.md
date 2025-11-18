# Mass-Communication-Macros
A VBA-based automation system that sends mass communication emails through Outlook using batches, and scans for undeliverable/bounce messages to update Excel statuses automatically.

Excel + Outlook Mass Communication Automation
=================================================

A VBA macro system to send bulk emails and detect undeliverable addresses automatically.

-------------------------------------------------
🚀 Overview
-------------------------------------------------
This project provides two fully automated VBA macros that streamline mass communication workflows using Excel + Outlook.

The automation supports:
- Sending bulk emails to large recipient lists
- Splitting recipients into BCC batches (e.g., 50, 100, 200)
- Using a pre-saved Outlook .msg email template
- Auto-updating Excel status for each recipient
- Scanning Outlook Inbox for undeliverable / bounce messages
- Matching failed deliveries back to Excel
- Updating the recipient list with "UNDELIVERED" status

This is a sanitized demo version with no internal or confidential information.

-------------------------------------------------
✨ Features
-------------------------------------------------

1. Bulk Email Sender (SendBulkEmailOnly)
----------------------------------------
- Reads email IDs from Column A in Excel
- Groups emails into user-defined BCC batches
- Uses a selected .msg Outlook template
- Drafts emails in Outlook for manual review
- Updates Excel status: 'Drafted <date/time>'

2. Bounce Detection (CheckUndeliveredLater)
-------------------------------------------
- Scans Outlook Inbox for bounce messages
- Detects subjects like:
  * 'Undeliverable'
  * 'Delivery has failed'
  * 'Returned mail'
- Extracts the failed email address using Regex
- Matches the address in Excel
- Updates status to: 'UNDELIVERED <date/time>'
- Handles duplicates correctly

-------------------------------------------------
🛠 How It Works
-------------------------------------------------

1. Prepare Excel Sheet
----------------------
Sheet1:
Column A → Email IDs
Column B → Status (auto-filled)

Example:
Email ID            | Status
--------------------|--------
test1@example.com   |
test2@example.com   |
test3@example.com   |

2. Run Bulk Email Sender
------------------------
- ALT + F8
- Select: SendBulkEmailOnly
- Enter batch size (e.g., 100)
- Choose .msg template
- Outlook drafts created
- Column B updated with timestamp

3. Run Undelivered Checker
--------------------------
- ALT + F8
- Select: CheckUndeliveredLater
- Macro scans Inbox
- Extracts email addresses
- Marks matching rows as UNDELIVERED

-------------------------------------------------
📨 Demo Email Template
-------------------------------------------------

Subject: Important Update – Please Review

Hi,

We are sharing an important update that may require your review.

Please take a moment to go through the information and reach out if any clarification is needed.
This is a general communication sent to a large set of recipients, and no immediate action is required unless specified.

Thank you for your time.

Best regards,
Demo Team
Communications Unit

-------------------------------------------------

🧩 Skills Demonstrated
-------------------------------------------------
- Excel VBA automation  
- Outlook object model  
- Regex-based text parsing  
- Batch processing logic  
- Email workflow optimization  



