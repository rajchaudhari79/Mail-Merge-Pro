# Mail Merge Pro 🚀

A premium, open-source Google Sheets Add-on that lets you send personalized mass emails directly from your spreadsheet using the Gmail API. Skip the expensive third-party subscriptions and send emails natively!

## ✨ Features

- **Native Gmail Integration:** Uses your existing Gmail Drafts as beautiful HTML email templates.
- **Dynamic Variables:** Automatically detects variables like `{{Name}}` or `{{Company}}` in your drafts and replaces them with data from your sheet.
- **Dynamic CC & BCC:** Add `CC` or `BCC` columns to your sheet to send dynamic carbon copies.
- **Dynamic Attachments:** Add an `Attachment` column with Google Drive links to automatically attach unique files for each recipient!
- **Scheduled Sending (Send Later):** Pick a date and time, and the script will automatically send your emails in the background.
- **Analytics Dashboard:** Built-in donut charts to see your exact Sent, Opened, Replied, and Bounced percentages in real-time.
- **Open Tracking:** Built-in invisible tracking pixel to see exactly when recipients open your emails.
- **Bounce & Reply Sync:** Fast inbox scanner that detects if an email bounced or was replied to, syncing the status right back to your spreadsheet.
- **Privacy First:** The script runs entirely inside your own Google Account. No third-party servers, no data sharing.

## 🛠️ Installation

Because this is a Google Apps Script project, installation is incredibly simple and doesn't require any downloads.

1. Open a new or existing Google Spreadsheet.
2. Click on **Extensions > Apps Script** in the top menu.
3. In the editor, delete any default code in `Code.gs` and replace it with the code from [`Code.gs`](./Code.gs) in this repository.
4. Click the **`+`** icon next to Files, select **HTML**, and name it exactly `sidebar`.
5. Replace the default HTML with the code from [`sidebar.html`](./sidebar.html).
6. Click the **Save** icon.

## 🚀 How to Use

1. Go back to your Google Spreadsheet and refresh the page.
2. A new custom menu called **Mail Merge Pro** will appear next to the "Help" menu at the top.
3. Add your data to the sheet. Make sure you have a column named **Email**. 
4. Go to Gmail and create a Draft email. Use variables like `{{Name}}` in the subject and body.
5. In your spreadsheet, click **Mail Merge Pro -> Start Mail Merge**.
6. The sidebar will appear. Select your draft, enter your Sender Name, and click **Start Mail Merge**!

## 📈 Setting up Open Tracking (Optional)

To track when people open your emails, you need to deploy the script as a Web App:
1. In the Apps Script editor, click **Deploy > New deployment**.
2. Select **Web app** from the gear icon.
3. Set **Execute as** to `Me` and **Who has access** to `Anyone`.
4. Click **Deploy** and copy the resulting Web App URL.
5. Paste this URL into the tracking field in the Mail Merge Pro sidebar.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
