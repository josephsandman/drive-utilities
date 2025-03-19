// Copyright Martin Hawksey 2020
//
// Licensed under the Apache License, Version 2.0 (the "License"); you may not
// use this file except in compliance with the License.  You may obtain a copy
// of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
// WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
// License for the specific language governing permissions and limitations under
// the License.
 
/**
 * @OnlyCurrentDoc
*/
 
/**
 * Sends emails to recipients based on data from a specified Google Sheet.
 * 
 * Retrieves recipient email addresses and the associated email sent status and 
 * uses a Gmail draft message as a template. Logs the status of sent emails directly
 * into the specified column. Prompts the user for missing header names if not provided.
 * 
 * @param {string} [subjectLine] - Subject line for the email draft message. Optional; prompts if not provided.
 * @param {string} thisSheet - The Google Sheet file URL or ID with mail merge data.
 * @param {string} thisTab - The name of the tab within the Google Sheet with mail merge data.
 * @param {string} [emailRecipients] - Header of the column containing recipient email addresses. Optional; prompts if not provided.
 * @param {string} [emailSent] - Header of the column where email sent dates are logged. Optional; prompts if not provided.
 * 
 * @return {void} This function does not return a value.
 * 
 * @throws {Error} If the specified sheet, tab, or required headers do not exist.
 * 
 * @example
 * sendEmails("Weekly Update", "1B2c...xyz", "Mail Merge");
 */
function sendEmails(subjectLine, thisSheet, thisTab, emailRecipients, emailSent) {
  console.log(`DEBUG: Parameters received by sendEmails(): \rsubjectLine: ${subjectLine}\rthisSheet: ${thisSheet}\rthisTab: ${thisTab}\remailRecipients: ${emailRecipients}\remailSent: ${emailSent}`);
  console.time('sendEmails() processing time');
  
  let RECIPIENT_COL = emailRecipients || getUserInput("Enter the header name for recipient email addresses:");
  let EMAIL_SENT_COL = emailSent || getUserInput("Enter the header name for email sent status:");

  if (!subjectLine) {
    subjectLine = Browser.inputBox( "Mail Merge",
                                    "Type or copy/paste the subject line of the Gmail " +
                                    "draft message you would like to mail merge with:",
                                    Browser.Buttons.OK_CANCEL);
    if (subjectLine === "cancel" || subjectLine === "") {
      console.error(`ERROR: Abort script due to prompt response: '${subjectLine}'`);
      return;
    }
  }

  if (!thisSheet || !thisTab) {
    sheet = SpreadsheetApp.getActiveSheet();
  } else {
    sheet = SpreadsheetApp.openById(getIdFromUrl(thisSheet)).getSheetByName(String(thisTab));
    if (!sheet) {
      console.warn(`WARNING: Sheet named '${thisTab}' was not found in '${thisSheet}'.`);
      sheet = SpreadsheetApp.getActiveSheet();
      console.log(`DEBUG: Proceeding with the current active sheet as default: '${sheet}'`);
    }
  }
  
  console.info(`INFO: Sending mail merge from '${sheet.getName()}' with subject: '${subjectLine}'`);
  
  const emailTemplate = getGmailTemplateFromDrafts_(subjectLine);
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();

  const heads = data.shift();

  if (!isValidHeaderRow(heads)) {
    console.error(`ERROR: Header row must contain at least two unique headers. Invalid headers: '${heads}'`);
    return;
  }

  console.log(`DEBUG: Headers array: '${heads}'`);

  if (!heads.includes(RECIPIENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${RECIPIENT_COL}'`);
    return;
  }

  if (!heads.includes(EMAIL_SENT_COL)) {
    console.error(`ERROR: Abort script due to missing column header: '${EMAIL_SENT_COL}'`);
    return;
  }
  
  console.log(`DEBUG: Ready to define emailSentColIdx: '${heads.indexOf(EMAIL_SENT_COL)}'`);
  
  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);

  console.log(`DEBUG: Email sent column: '${emailSentColIdx}'`);
  
  const obj = data.map(r => (heads.reduce((o, k, i) => (o[k] = r[i] || '', o), {})));

  const out = [];

  console.time("Total row processing time");
  obj.forEach(function(row, rowIdx) {
    if (row[RECIPIENT_COL] === '' && row[EMAIL_SENT_COL] === '' && !sheet.isRowHiddenByFilter(rowIdx + 2)) {
      console.time(`Row '${rowIdx + 2}' processing time `);
      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        MailApp.sendEmail(
          row[RECIPIENT_COL], 
          msgObj.subject, 
          msgObj.text, 
          {
            htmlBody: msgObj.html,
            // bcc: 'a.bbc@email.com',
            // cc: 'a.cc@email.com',
            // name: 'name of the sender',
            // replyTo: 'a.reply@email.com',
            // noReply: true, // if the email should be sent from a generic no-reply email address (not available to gmail.com users)
            attachments: emailTemplate.attachments,
            inlineImages: emailTemplate.inlineImages
          }
        );
        out.push([new Date()]);
        console.info(`INFO: Email sent to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2})`);
      } catch (e) {
        out.push([e.message]);
        console.error(`ERROR: Failed to send email to '${row[RECIPIENT_COL]}' (Row ${rowIdx + 2}). Error: ${e.message}`);
      } finally {
        console.timeEnd(`Row '${rowIdx + 2}' processing time `);
      }
    } else {
      if (row[EMAIL_SENT_COL] !== '') {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Email already sent.`);
      } 
      if (sheet.isRowHiddenByFilter(rowIdx + 2)) {
        console.log(`DEBUG: Skipping Row ${rowIdx + 2} - Row hidden by filter.`);
      }
      out.push([row[EMAIL_SENT_COL]]);
    }
  });
  console.timeEnd("Total row processing time");
  
  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);
  console.log(`DEBUG: Finished writing outputs to rows.`);
  
  /**
   * Get a Gmail draft message by matching the subject line.
   * 
   * @param {string} subject_line - Subject line to search for the draft message.
   * @return {object} Contains the subject, plain and HTML message body, and attachments.
   */
  function getGmailTemplateFromDrafts_(subject_line) {
    console.info(`INFO: Searching for draft with subject line: '${subject_line}'`);
    try {
      const drafts = GmailApp.getDrafts();
      console.debug(`DEBUG: Total drafts retrieved: ${drafts.length}`);
      const draft = drafts.filter(subjectFilter_(subject_line))[0];
      if (!draft) {
        console.warn(`WARNING: No draft found matching the subject line: '${subject_line}'`);
        return;
      } else {
        console.info(`INFO: Draft found with subject: '${draft.getMessage().getSubject()}'`);
      }
      const msg = draft.getMessage();

      const allInlineImages = draft.getMessage().getAttachments({includeInlineImages: true, includeAttachments: false});
      console.debug(`DEBUG: Total inline images found: ${allInlineImages.length}`);
      const attachments = draft.getMessage().getAttachments({includeInlineImages: false});
      const htmlBody = msg.getBody();
      console.debug(`DEBUG: Draft HTML body length: ${htmlBody.length}`);

      const img_obj = allInlineImages.reduce((obj, i) => (obj[i.getName()] = i, obj), {});

      const imgexp = RegExp('<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>', 'g');
      const matches = [...htmlBody.matchAll(imgexp)];
      console.debug(`DEBUG: Total inline image matches found in HTML body: ${matches.length}`);

      const inlineImagesObj = {};
      matches.forEach(match => inlineImagesObj[match[1]] = img_obj[match[2]]);
      console.info(`INFO: Returning message details for subject: '${subject_line}'`);

      return {
        message: {
          subject: subject_line,
          text: msg.getPlainBody(),
          html: htmlBody
        },
        attachments: attachments,
        inlineImages: inlineImagesObj
      };
    } catch (e) {
      console.error(`ERROR: No Gmail draft found: '${e.message}'`);
      return;
    }

    /**
     * Filter draft objects by the matching subject line.
     * 
     * @param {string} subject_line - Subject line to search for the draft message.
     * @return {function} GmailDraft object filter function.
     */
    function subjectFilter_(subject_line) {
      return function(element) {
        if (element.getMessage().getSubject() === subject_line) {
          return element;
        }
      }
    }
  }
  
  /**
   * Fills template string with data object values.
   * 
   * @param {string} template - Template string containing {{}} markers for replacement.
   * @param {object} data - Object with values to replace the {{}} markers.
   * @return {object} Message replaced with data.
   */
  function fillInTemplateFromObject_(template, data) {
    let template_string = JSON.stringify(template);

    template_string = template_string.replace(/{{[^{}]+}}/g, key => {
      const text = data[key.replace(/[{}]+/g, "")] || "";
      return escapeData_(text.replace(/\n/g, '<br>'));
    });
    return JSON.parse(template_string);
  }

  /**
   * Escapes cell data to make it JSON safe.
   * 
   * @param {string} str - String to escape special characters.
   * @return {string} Escaped string.
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

  function isValidHeaderRow(headers) {
    const uniqueHeaders = new Set(headers.filter(header => header.trim() !== ""));
    return uniqueHeaders.size >= 2;
  };

  console.timeEnd('sendEmails() processing time');
}