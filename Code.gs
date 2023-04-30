// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/mail-merge

/**
 * @OnlyCurrentDoc
 */

/**
 * Change these to match the column names you are using for email 
 * recipient addresses and email sent column.
 */
const EMAIL_COL = "Email";
const EMAIL_SENT_COL = "Email Sent";
const EMAIL_SUBJECT_COL = "Email Subject"

/** 
 * Creates the menu item "Mail Merge" for user to run scripts on drop-down.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mail Merge')
    .addItem('Send Emails', 'sendEmails') // Call the sendEmails function defined below
    .addToUi();
}

/**
 * Sends emails from sheet data.
 * @param {Sheet} sheet to read data from
*/
function sendEmails(sheet = SpreadsheetApp.getActiveSheet()) {
  const rows = sheet.getDataRange().getDisplayValues();

  // Convert the 2d array into an object array
  // See https://stackoverflow.com/a/22917499/1027723
  // See https://mashe.hawksey.info/?p=17869/#comment-184945 for more details
  const header = rows.shift();
  const rowsObj = rows.map(r => (header.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const recipients = []; // list of recipients to email
  rowsObj.forEach(function (row, rowIdx) {
    rowIdx = rowIdx + 2; // Account for skipped header and rows starting at 1
    // Only send email if the row is not hidden by a filter or by the user and the email_sent cell is blank
    if (sheet.isRowHiddenByFilter(rowIdx)) {
      console.log(`Row ${rowIdx} is hidden by filter ${JSON.stringify(row)}`);
    } else if (sheet.isRowHiddenByUser(rowIdx)) {
      console.log(`Row ${rowIdx} is hidden by user ${JSON.stringify(row)}`);
    } else if (row[EMAIL_SENT_COL] !== '') {
      console.log(`Row ${rowIdx} already received an email ${JSON.stringify(row)}`);
    } else {
      console.log(`Will send email to row ${rowIdx}: ${JSON.stringify(row)}`);
      recipients.push(row[EMAIL_COL] + ` (${row[EMAIL_SUBJECT_COL]})`);
    }
  });

  // Display a message box with a preview of the emails to send
  const okCancelRes = Browser.msgBox("Mail Merge", `Sending ${recipients.length} emails to\\n• ` +
    recipients.slice(0, 20).join("\\n• ") +
    (recipients.length > 20 ? "\\n..." : ""),
    Browser.Buttons.OK_CANCEL);
  // If the user clicked the Cancel button
  if (okCancelRes === "cancel") return;

  // Create an array to record sent emails
  const emailSentResults = [];

  rowsObj.forEach(function (row, rowIdx) {
    rowIdx = rowIdx + 2; // Account for skipped header and rows starting at 1
    // Only send email if the row is not hidden by a filter or by the user and the email_sent cell is blank
    if (sheet.isRowHiddenByFilter(rowIdx) || sheet.isRowHiddenByUser(rowIdx) || row[EMAIL_SENT_COL] !== '') {
      emailSentResults.push([row[EMAIL_SENT_COL]]);
      return;
    }
    try {
      // Try to get the Gmail template, fill it in, and send the email
      const emailTemplate = getGmailTemplateFromDrafts_(row[EMAIL_SUBJECT_COL]);
      const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);
      sendEmail_(row[EMAIL_COL], msgObj, emailTemplate);
      emailSentResults.push([new Date()]); // Record email sent date
    } catch (e) {
      emailSentResults.push([e.message]); // Record the error
    }
  });

  // Update the sheet with new data
  sheet.getRange(2, header.indexOf(EMAIL_SENT_COL) + 1, emailSentResults.length)
    .setValues(emailSentResults);

  /**
   * Send an email to the specified recipient.
   * See https://developers.google.com/apps-script/reference/gmail/gmail-app#sendEmail(String,String,String,Object).
   * @param {string} recipient to send the email to
   * @param {object} msgObj
   * @param {object} emailTemplate
   */
  function sendEmail_(recipient, msgObj, emailTemplate) {
    // If you need to send emails with unicode/emoji characters change GmailApp for MailApp
    // Uncomment advanced parameters as needed (see docs for limitations)
    GmailApp.sendEmail(recipient, msgObj.subject, msgObj.text, {
      htmlBody: msgObj.html,
      // bcc: 'a.bbc@email.com',
      // cc: 'a.cc@email.com',
      // from: 'an.alias@email.com',
      name: "Parents' Club of Stanford - Membership",
      // replyTo: 'a.reply@email.com',
      // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
      attachments: emailTemplate.attachments,
      inlineImages: emailTemplate.inlineImages
    });
  }

  /**
   * Get a Gmail draft message by matching the subject line.
   * @param {string} subjectLine to search for draft message
   * @return {object} containing the subject, plain and html message body and attachments
  */
  function getGmailTemplateFromDrafts_(subjectLine) {
    try {
      // Get drafts
      const drafts = GmailApp.getDrafts();
      // Filter the drafts that match subject line
      const draft = drafts.filter(subjectFilter_(subjectLine))[0];
      // Get the message object
      const msg = draft.getMessage();

      // Handle inline images and attachments so they can be included in the merge
      // Based on https://stackoverflow.com/a/65813881/1027723
      const allInlineImages = draft.getMessage().getAttachments({ includeInlineImages: true, includeAttachments: false });
      const attachments = draft.getMessage().getAttachments({ includeInlineImages: false });
      const htmlBody = msg.getBody();

      // Create an inline image object with the image name as key 
      // (can't rely on image index as array based on insert order)
      const imgObj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

      //Regexp searches for all img string positions with cid
      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];

      //Initiates the allInlineImages object
      const inlineImagesObj = {};
      // built an inlineImagesObj from inline image matches
      matches.forEach(match => inlineImagesObj[match[1]] = imgObj[match[2]]);

      return {
        message: { subject: subjectLine, text: msg.getPlainBody(), html: htmlBody },
        attachments: attachments, inlineImages: inlineImagesObj
      };
    } catch (e) {
      throw new Error("Oops - can't find Gmail draft");
    }

    /**
     * Filter draft objects with the matching subject linemessage by matching the subject line.
     * @param {string} subjectLine to search for draft message
     * @return {object} GmailDraft object
    */
    function subjectFilter_(subjectLine) {
      return function (element) {
        if (element.getMessage().getSubject() === subjectLine) {
          return element;
        }
      }
    }
  }

  /**
   * Fill template string with data object
   * @see https://stackoverflow.com/a/378000/1027723
   * @param {string} template string containing {{}} markers which are replaced with data
   * @param {object} data object used to replace {{}} markers
   * @return {object} message replaced with data
  */
  function fillInTemplateFromObject_(template, data) {
    // We have two templates: one for plain text and one for the html body.
    // Stringifing the object means we can do a global replace.
    let templateString = JSON.stringify(template);

    // Perform token replacement
    templateString = templateString.replace(/{{[^{}]+}}/g, key => {
      return escapeData_(data[key.replace(/[{}]+/g, "")] || "");
    });
    return JSON.parse(templateString);
  }

  /**
   * Escape cell data to make JSON safe
   * @see https://stackoverflow.com/a/9204218/1027723
   * @param {string} str to escape JSON special characters from
   * @return {string} escaped string
  */
  function escapeData_(str) {
    return str
      .replace(/[\\]/g, '\\\\')
      .replace(/[\"]/g, '\\\"')
      .replace(/[\/]/g, '\\/')
      .replace(/[\b]/g, '\\b')
      .replace(/[\f]/g, '\\f')
      .replace(/[\n]/g, '\\n')
      .replace(/[\r]/g, '\\r')
      .replace(/[\t]/g, '\\t');
  };
}
