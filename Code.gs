/**
 * Adds a custom menu to the active spreadsheet.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge Pro')
      .addItem('Start Mail Merge', 'showSidebar')
      .addToUi();
}

/**
 * Opens a sidebar in the document containing the HTML UI.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('Mail Merge Pro')
      .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Handles GET requests to the Web App for Open Tracking.
 */
function doGet(e) {
  if (e.parameter.action === 'open') {
    try {
      const sheetId = e.parameter.sheetId;
      const row = parseInt(e.parameter.row);
      
      const ss = SpreadsheetApp.openById(sheetId);
      const sheet = ss.getSheets()[0];
      
      const headers = sheet.getDataRange().getValues()[0];
      const statusColIndex = headers.findIndex(h => h.toLowerCase() === 'merge status');
      
      if (statusColIndex !== -1) {
        const currentStatus = sheet.getRange(row, statusColIndex + 1).getValue();
        // Only update if it's currently "Sent". If it's already "Opened", "Replied", etc, we keep it.
        if (String(currentStatus).startsWith('Sent')) {
           const timestamp = new Date().toLocaleString();
           sheet.getRange(row, statusColIndex + 1).setValue(`Opened: ${timestamp}`);
        }
      }
    } catch (err) {
      console.error('Tracking Error:', err);
    }
  }
  
  // Return a 1x1 transparent SVG image
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"></svg>';
  return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.SVG);
}

/**
 * Gets a list of Gmail drafts.
 */
function getGmailDrafts() {
  try {
    const drafts = GmailApp.getDrafts();
    return drafts.map(draft => {
      return {
        id: draft.getId(),
        subject: draft.getMessage().getSubject() || "(No Subject)"
      };
    });
  } catch (e) {
    console.error(e);
    return [];
  }
}

/**
 * Gets headers from the active sheet.
 */
function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return [];
  return data[0].filter(String);
}

/**
 * Gets the Web App URL from User Properties (saved from Sidebar).
 */
function getWebAppUrl() {
  return PropertiesService.getUserProperties().getProperty('WEB_APP_URL') || '';
}

/**
 * Saves the Web App URL from the sidebar.
 */
function saveWebAppUrl(url) {
  PropertiesService.getUserProperties().setProperty('WEB_APP_URL', url.trim());
  return true;
}

/**
 * Main function to start the mail merge.
 */
function sendMailMerge(draftId, enableTracking, senderName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("Sheet requires at least a header row and one data row.");

  // Save the Spreadsheet ID so background scanners can find it later
  PropertiesService.getScriptProperties().setProperty('ACTIVE_SPREADSHEET_ID', ss.getId());

  const webAppUrl = getWebAppUrl();
  if (enableTracking && !webAppUrl) {
    throw new Error("You must Deploy as a Web App to use Open Tracking. Please uncheck Tracking or deploy the script.");
  }

  const headers = data[0];
  const draft = GmailApp.getDraft(draftId);
  if (!draft) throw new Error("Selected draft not found.");
  
  const draftMessage = draft.getMessage();
  const subjectTemplate = draftMessage.getSubject();
  const bodyTemplate = draftMessage.getBody(); 
  const plainBodyTemplate = draftMessage.getPlainBody(); 
  const attachments = draftMessage.getAttachments();
  
  let emailColIndex = headers.findIndex(h => h.toLowerCase() === 'email' || h.toLowerCase() === 'to' || h.toLowerCase() === 'email address');
  if (emailColIndex === -1) throw new Error("Could not find a column named 'Email' or 'To'.");

  let statusColIndex = headers.findIndex(h => h.toLowerCase() === 'merge status');
  if (statusColIndex === -1) {
    statusColIndex = headers.length;
    sheet.getRange(1, statusColIndex + 1).setValue('Merge Status');
  }

  let sentCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const emailAddress = row[emailColIndex];
    const currentStatus = row[statusColIndex];

    // Skip if already processed
    if (!emailAddress || String(currentStatus).startsWith('Sent:') || String(currentStatus).startsWith('Opened:') || String(currentStatus).startsWith('Replied:')) continue;

    let finalSubject = subjectTemplate;
    let finalBody = bodyTemplate;
    let finalPlainBody = plainBodyTemplate;

    // Replace variables
    for (let j = 0; j < headers.length; j++) {
      if (!headers[j]) continue;
      const escapedHeader = headers[j].replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
      const regex = new RegExp(`{{${escapedHeader}}}`, 'gi');
      finalSubject = finalSubject.replace(regex, row[j] || '');
      finalBody = finalBody.replace(regex, row[j] || '');
      finalPlainBody = finalPlainBody.replace(regex, row[j] || '');
    }

    // Inject Tracking Pixel safely before </body> to prevent malformed HTML (Spam trigger)
    if (enableTracking && webAppUrl) {
      const trackingUrl = `${webAppUrl}?action=open&sheetId=${ss.getId()}&row=${i + 1}`;
      const imgTag = `<img src="${trackingUrl}" width="1" height="1" alt="" style="display:none;"/>`;
      if (finalBody.toLowerCase().includes('</body>')) {
        finalBody = finalBody.replace(/<\/body>/i, imgTag + '</body>');
      } else {
        finalBody += imgTag;
      }
    }

    try {
      const options = {
        htmlBody: finalBody,
        attachments: attachments,
        cc: draftMessage.getCc() || undefined,
        bcc: draftMessage.getBcc() || undefined,
        replyTo: draftMessage.getReplyTo() || undefined
      };
      
      if (senderName) {
        options.name = senderName;
      }
      
      GmailApp.sendEmail(emailAddress, finalSubject, finalPlainBody, options);
      
      const timestamp = new Date().toLocaleString();
      sheet.getRange(i + 1, statusColIndex + 1).setValue(`Sent: ${timestamp}`);
      sentCount++;
      Utilities.sleep(1000); 
    } catch (e) {
      sheet.getRange(i + 1, statusColIndex + 1).setValue(`Error: ${e.message}`);
    }
  }

  return { success: true, count: sentCount };
}

/**
 * Scans the inbox for bounces and replies.
 * Can be run manually from the UI or via a Time-Driven Trigger.
 */
function scanInboxForRepliesAndBounces() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('ACTIVE_SPREADSHEET_ID');
  if (!sheetId) throw new Error("No active mail merge found. Run a merge first.");
  
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { updatedCount: 0 };
  
  const headers = data[0];
  const emailColIndex = headers.findIndex(h => h.toLowerCase() === 'email' || h.toLowerCase() === 'to');
  const statusColIndex = headers.findIndex(h => h.toLowerCase() === 'merge status');
  
  if (emailColIndex === -1 || statusColIndex === -1) return { updatedCount: 0 };

  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusColIndex]);
    const email = data[i][emailColIndex];
    if (!email) continue;

    // Only check rows that are Sent or Opened
    if (status.startsWith('Sent') || status.startsWith('Opened')) {
      
      // 1. Check for Bounces
      // mailer-daemon@googlemail.com containing the recipient's email address
      const bounceThreads = GmailApp.search(`from:mailer-daemon@googlemail.com "${email}" newer_than:14d`);
      if (bounceThreads.length > 0) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue('Bounced: ' + new Date().toLocaleString());
        updatedCount++;
        continue;
      }
      
      // 2. Check for Replies
      const replyThreads = GmailApp.search(`from:${email} newer_than:14d`);
      if (replyThreads.length > 0) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue('Replied: ' + new Date().toLocaleString());
        updatedCount++;
      }
    }
  }

  return { success: true, updatedCount: updatedCount };
}
