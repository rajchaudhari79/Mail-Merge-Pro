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
        // Only update if it's currently "Sent"
        if (String(currentStatus).startsWith('Sent')) {
           const timestamp = new Date().toLocaleString();
           sheet.getRange(row, statusColIndex + 1).setValue(`Opened: ${timestamp}`);
        }
      }
    } catch (err) {
      console.error('Tracking Error:', err);
    }
  }
  
  const svg = '<svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"></svg>';
  return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.SVG);
}

/**
 * Gets a list of Gmail drafts.
 */
function getGmailDrafts() {
  try {
    const drafts = GmailApp.getDrafts();
    return drafts.map(draft => ({
      id: draft.getId(),
      subject: draft.getMessage().getSubject() || "(No Subject)"
    }));
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
 * Analyzes the 'Merge Status' column to return counts for the dashboard.
 */
function getAnalytics() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return { sent: 0, opened: 0, replied: 0, bounced: 0, error: 0, pending: 0 };

  const headers = data[0];
  const statusColIndex = headers.findIndex(h => h.toLowerCase() === 'merge status');
  
  let stats = { sent: 0, opened: 0, replied: 0, bounced: 0, error: 0, pending: 0 };
  
  if (statusColIndex === -1) {
    stats.pending = data.length - 1;
    return stats;
  }

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][statusColIndex] || '');
    if (status.startsWith('Sent')) stats.sent++;
    else if (status.startsWith('Opened')) stats.opened++;
    else if (status.startsWith('Replied')) stats.replied++;
    else if (status.startsWith('Bounced')) stats.bounced++;
    else if (status.startsWith('Error')) stats.error++;
    else stats.pending++;
  }
  return stats;
}

/**
 * Extracts a Drive file ID from a URL
 */
function extractDriveId(url) {
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

/**
 * Main function to start the mail merge.
 * Now expects a config object: { draftId, enableTracking, senderName, sheetId }
 */
function sendMailMerge(config) {
  let ss;
  if (config.sheetId) {
    ss = SpreadsheetApp.openById(config.sheetId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    config.sheetId = ss.getId();
  }
  
  const sheet = ss.getSheets()[0]; // Use first sheet
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) throw new Error("Sheet requires at least a header row and one data row.");

  PropertiesService.getScriptProperties().setProperty('ACTIVE_SPREADSHEET_ID', config.sheetId);

  const webAppUrl = getWebAppUrl();
  if (config.enableTracking && !webAppUrl) {
    throw new Error("You must set the Web App URL for Open Tracking.");
  }

  const headers = data[0];
  const draft = GmailApp.getDraft(config.draftId);
  if (!draft) throw new Error("Selected draft not found.");
  
  const draftMessage = draft.getMessage();
  const subjectTemplate = draftMessage.getSubject();
  const bodyTemplate = draftMessage.getBody(); 
  const plainBodyTemplate = draftMessage.getPlainBody(); 
  const baseAttachments = draftMessage.getAttachments() || [];
  
  let emailColIndex = headers.findIndex(h => h.toLowerCase() === 'email' || h.toLowerCase() === 'to' || h.toLowerCase() === 'email address');
  if (emailColIndex === -1) throw new Error("Could not find a column named 'Email' or 'To'.");

  // Optional Columns
  const ccColIndex = headers.findIndex(h => h.toLowerCase() === 'cc');
  const bccColIndex = headers.findIndex(h => h.toLowerCase() === 'bcc');
  const attachmentColIndex = headers.findIndex(h => h.toLowerCase() === 'attachment' || h.toLowerCase() === 'file');

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

    // Inject Tracking Pixel safely
    if (config.enableTracking && webAppUrl) {
      const trackingUrl = `${webAppUrl}?action=open&sheetId=${config.sheetId}&row=${i + 1}`;
      const imgTag = `<img src="${trackingUrl}" width="1" height="1" alt="" style="display:none;"/>`;
      if (finalBody.toLowerCase().includes('</body>')) {
        finalBody = finalBody.replace(/<\/body>/i, imgTag + '</body>');
      } else {
        finalBody += imgTag;
      }
    }

    // Handle Dynamic CC/BCC
    let finalCc = draftMessage.getCc() || undefined;
    let finalBcc = draftMessage.getBcc() || undefined;
    if (ccColIndex !== -1 && row[ccColIndex]) {
       finalCc = finalCc ? `${finalCc}, ${row[ccColIndex]}` : row[ccColIndex];
    }
    if (bccColIndex !== -1 && row[bccColIndex]) {
       finalBcc = finalBcc ? `${finalBcc}, ${row[bccColIndex]}` : row[bccColIndex];
    }

    // Handle Dynamic Attachments
    let finalAttachments = [...baseAttachments];
    if (attachmentColIndex !== -1 && row[attachmentColIndex]) {
       try {
         const fileId = extractDriveId(row[attachmentColIndex]);
         if (fileId) {
            const blob = DriveApp.getFileById(fileId).getBlob();
            finalAttachments.push(blob);
         }
       } catch(e) {
         console.error("Could not fetch attachment for row " + (i+1), e);
       }
    }

    try {
      const options = {
        htmlBody: finalBody,
        attachments: finalAttachments,
        cc: finalCc,
        bcc: finalBcc,
        replyTo: draftMessage.getReplyTo() || undefined
      };
      
      if (config.senderName) {
        options.name = config.senderName;
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
 * Schedules a mail merge for later.
 */
function scheduleMerge(config, dateStr) {
  config.sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  // Save config to properties
  PropertiesService.getScriptProperties().setProperty('SCHEDULED_CONFIG', JSON.stringify(config));
  
  const date = new Date(dateStr);
  if (date.getTime() <= new Date().getTime()) {
    throw new Error("Scheduled time must be in the future.");
  }
  
  ScriptApp.newTrigger('executeScheduledMerge')
    .timeBased()
    .at(date)
    .create();
    
  return { success: true, date: date.toLocaleString() };
}

/**
 * Background trigger function for Scheduled Merge
 */
function executeScheduledMerge(e) {
  try {
    const configStr = PropertiesService.getScriptProperties().getProperty('SCHEDULED_CONFIG');
    if (!configStr) return;
    
    const config = JSON.parse(configStr);
    sendMailMerge(config);
    
    PropertiesService.getScriptProperties().deleteProperty('SCHEDULED_CONFIG');
    
    // Delete the trigger that just ran
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      if (t.getHandlerFunction() === 'executeScheduledMerge') {
         ScriptApp.deleteTrigger(t);
      }
    });
  } catch (err) {
    console.error("Scheduled Merge Error:", err);
  }
}

/**
 * Scans the inbox for bounces and replies.
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

    if (status.startsWith('Sent') || status.startsWith('Opened')) {
      const bounceThreads = GmailApp.search(`from:mailer-daemon@googlemail.com "${email}" newer_than:14d`);
      if (bounceThreads.length > 0) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue('Bounced: ' + new Date().toLocaleString());
        updatedCount++;
        continue;
      }
      
      const replyThreads = GmailApp.search(`from:${email} newer_than:14d`);
      if (replyThreads.length > 0) {
        sheet.getRange(i + 1, statusColIndex + 1).setValue('Replied: ' + new Date().toLocaleString());
        updatedCount++;
      }
    }
  }

  return { success: true, updatedCount: updatedCount };
}
