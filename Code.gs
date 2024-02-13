/**
 * Gmail to PDF Script v2.1.2
 *
 * @copyright 2023-2024 Andy Briggs <andy.briggs@vmfh.org>
 * @license MIT
 * @link https://github.com/pharmot/gmail-to-pdf/blob/main/Code.gs
 * Created        : 2023-12-29
 * Last modified  : 2024-02-12
 * 
 * If you change any of these variables, run the 'setup' function again to
 * apply the changes. 
 */

/** Name of the label you use to mark emails as needing to be exported */
const exportLabel  = "Export";

/** Label to add to threads after they've been exported */
const doneLabel = "Saved PDF";

/** Label to add to threads if they encounter an error while exporting */
const errorLabel = "Export Error";

/** Label to add if threads will be deleted soon and need review */
const reviewLabel = "Review for Export"

/** Name of your email archive folder in your Drive */
const driveFolder  = "Gmail Archive";

/** Labels to exclude from review (besides the above) */
const excludeLabels = [
  "Save Not Needed",
  "Patient Specific",
];

/** Set to false if you do not want to check for emails to review weekly */
const checkWeekly = true;

/** Set to false if you do not want an email when the weekly check is complete */
const emailWeekly = true;

/** Set to false if you do not want to export daily */
const exportDaily = true;

/** Set to false to disable email notification if export fails */
const emailOnError = true;

/** Set to false if you don't want to export attachments as separate files */
const exportAttachments = true;

/**
 * List the file names of attachments you don't want to export, such as logos
 * and calendar invites. Don't include "image.png" or "image001.png",
 * "image002.png", etc. as these are the generic names assigned to unnamed
 * images. File names are not case-sensitive, but must include the file extension.
 */
const excludeAttachments = [
  "vmfh.png",
  "logo.png",
  "vmfh logo.png",
  "service management.jpg",
  "invite.ics",
];


/**
 * =============================================================================
 * Do not modify below this point
 * =============================================================================
 */
const setup = () => {
  ScriptStatus.set(exportLabel, '_exportLabel');
  ScriptStatus.set(doneLabel, '_doneLabel');
  ScriptStatus.set(errorLabel, '_errorLabel');
  ScriptStatus.set(driveFolder, '_driveFolder');
  ScriptStatus.set(reviewLabel, '_reviewLabel');
  ScriptStatus.set(exportAttachments, '_exportAttachments');
  ScriptStatus.set(emailWeekly, '_emailWeekly');
  ScriptStatus.set(emailOnError, '_emailOnError');
  ScriptStatus.set(JSON.stringify(excludeAttachments), '_excludeAttachments');
  ScriptStatus.set(JSON.stringify(excludeLabels), '_excludeLabels' );
  Logger.log('[SCRIPT SETUP] - Variables updated')
  const myExportLabel = getOrCreateLabel(exportLabel);
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if ( trigger.getHandlerFunction() !== 'resumeExport') {
      ScriptApp.deleteTrigger(trigger)
    }    
  }); 

  if ( exportDaily ) {
    ScriptApp.newTrigger('main').timeBased().everyDays(1).atHour(3).create();
    Logger.log('- Created daily trigger to export emails labeled "%s"', exportLabel);
  }  
  if ( checkWeekly ) {
    const now = new Date();
    const runAt = new Date(now.getFullYear(), now.getMonth(), now.getDate(), now.getHours(), now.getMinutes() + 1,0,0);
    ScriptApp.newTrigger('reviewEmails').timeBased().everyWeeks(1).onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(12).create();
    ScriptApp.newTrigger('reviewEmails').timeBased().at(runAt).create();
    Logger.log('- Created weekly trigger to add "%s" label to emails that need review. Will run now, then every Monday at 12:00 thereafter.', reviewLabel);
    Logger.log('OK to close this tab')
  }
}

const main = () => {
  if ( PropertiesService.getScriptProperties().getProperty('_exportLabel') === null ) {
    throw new Error("Must run 'setup' script first");
  }
  ScriptStatus.set('running');
  const unfinished = exportEmails();
  if ( unfinished === true) {
    ScriptStatus.set('not running');
    console.log('Batch complete, will run again');
    return new Trigger('resumeExport', 1);
  }
  ScriptStatus.set('done');
  console.log('Export finished, beginning clean up')
  cleanUp();
}
const reviewEmails = (initial=true) => {
  const _reviewLabel = ScriptStatus.get('_reviewLabel');
  const _exportLabel = ScriptStatus.get('_exportLabel');
  const _doneLabel = ScriptStatus.get('_doneLabel');
  const _errorLabel = ScriptStatus.get('_errorLabel');
  const _driveFolder = ScriptStatus.get('_driveFolder');
  const _emailWeekly = ScriptStatus.get('_emailWeekly');
  
  /** @type Array */
  let _excludeLabels = JSON.parse(ScriptStatus.get('_excludeLabels'));
  _excludeLabels.push(_exportLabel, _doneLabel, _errorLabel);

  let d = new Date();
  d.setFullYear(d.getFullYear()-1);
  d.setMonth(d.getMonth() + 1);
  const archiveDate = Utilities.formatDate(d, "America/Los_Angeles", "yyyy-MM-dd");

  let str = `before:${archiveDate} -label:${cleanLabel(_reviewLabel)}`;
  _excludeLabels.forEach(lbl => str += ` -label:${cleanLabel(lbl)}`);
  
  const lblReview = getOrCreateLabel(_reviewLabel);

  Logger.log("SEARCH: %s", str);
  const threads = GmailApp.search(str);
  const numThreads = threads.length;
  Logger.log("Found %s threads", numThreads);
  threads.forEach( thread => thread.addLabel(lblReview) );
  if ( numThreads === 500 ) {
    reviewEmails(false);
  } else {
    if ( numThreads > 0 || !initial ) {
      const exported = ScriptStatus.count();

      if ( _emailWeekly ) {
        const folders = DriveApp.getFoldersByName(_driveFolder);
        const folder = folders.next();
        const folderUrl = folder.getUrl();

        const emailOpts = {
          htmlBody: `<h3 style="font-family: Montserrat,Roboto,Arial;color: #55a63a">GmailToPdf Script &ndash; Weekly Summary</h3>
          <p style="font-family:Roboto,Arial;font-size:11pt;max-width:500px;">
          <strong>${parseInt(exported)}</strong> emails with the ${getLabelSpan(_exportLabel)} label have been saved as PDFs to your 
          <a style="color:#1155cc" href="${folderUrl}">${_driveFolder} folder</a> and relabeled as ${getLabelSpan(_doneLabel)}.</p>
          <p style="font-family:Roboto,Arial;font-size:11pt;max-width:500px;">
          Emails that will be deleted within the next month have been labeled ${getLabelSpan(_reviewLabel)}.
          Review <a style="color:#1155cc" href="https://mail.google.com/mail/u/0/#label/${_reviewLabel.replace(/ /,'+')}">
          all threads with this label</a> and add the ${getLabelSpan(_exportLabel)} label if they need to be exported.</p>
          <p style="font-family:Roboto,Arial;font-size:11pt;max-width:500px;">
          <em style="color:#54565a">Labels excluded from search: ${_excludeLabels.join(", ")}</em></p>`,
          to: Session.getActiveUser().getEmail(),
          subject: `Weekly Summary: Email export and review for deletion`,
          noReply: true,
          name: "Gmail Export Script",
        }
        MailApp.sendEmail(emailOpts);
      }
    }    
  }
}
const resumeExport = e => {
  if ( ScriptStatus.get() === 'done' ) {
    console.log('Was completed. Cleaning up.');
    cleanUp();
    return;
  }
  if ( ScriptStatus.get() === 'running' ) {
    console.log('Already running, exiting');
    return;
  }
  ScriptStatus.set('running')
  const unfinished = exportEmails();
  if ( unfinished === true) {
    ScriptStatus.set('not running')
    console.log('Batch complete, will run again');
    return;
  } 
  ScriptStatus.set('done');
  console.log('Export finished, beginning clean up')
  cleanUp();
};

const exportEmails = () => {
  let currentThread;
  const _exportLabel = ScriptStatus.get('_exportLabel');
  const _doneLabel = ScriptStatus.get('_doneLabel');
  const _errorLabel = ScriptStatus.get('_errorLabel');
  const _driveFolder = ScriptStatus.get('_driveFolder');
  const _reviewLabel = ScriptStatus.get('_reviewLabel');
  const _exportAttachments = ScriptStatus.get('_exportAttachments');
  const _emailOnError = ScriptStatus.get('_emailOnError');
  const _excludeAttachments = JSON.parse(ScriptStatus.get('_excludeAttachments')).map(n => n.toUpperCase() );
  try {
    const tz = Session.getScriptTimeZone()
    const maxThreads = 5;
    
    const threads = GmailApp.search(
      `label:${cleanLabel(_exportLabel)} -label:${cleanLabel(_errorLabel)}`,
      0,
      maxThreads
    );
    if (threads.length > 0) {
      console.log(`Found ${threads.length} emails to export`);
      const folders = DriveApp.getFoldersByName(_driveFolder);
      const folder = folders.hasNext() ?
        folders.next() :
        DriveApp.createFolder(_driveFolder);

      threads.forEach(thread => {

        currentThread = thread.getId();

        let html = `<html><style type="text/css">body{font-size: 12px;padding:0 10px;min-width:700px;-webkit-print-color-adjust: exact;}body>dl.email-meta{font-family:"Helvetica Neue",Helvetica,Arial,sans-serif;font-size: 12px;padding: 10px 0;margin: 5px 0;border: 1px solid #aaa;page-break-before:always}body>dl.email-meta:first-child {page-break-before:auto}body > dl.email-meta dt{color:#808080;float:left;width:60px;clear:left;text-align:right;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-style:normal;font-weight:700;line-height:1.2}body>dl.email-meta dd{margin-left:70px;line-height:1.2}body>dl.email-meta dd a {color:#808080;font-size:0.85em;text-decoration:none;font-weight:normal}body>dl.email-meta dd.strong {font-weight:bold}body>dl.email-meta dt.subject,body > dl.email-meta dd.subject,{font-size: 14px;}body>div.email-attachments {font-size:0.85em;color:#999}</style><body>`

        const messages = thread.getMessages();
        const attachments = [];
        const subject = thread.getFirstMessageSubject();
        const firstDate = messages[0].getDate();
        const fileDate = Utilities.formatDate(firstDate, tz, "yyyy-MM-dd HHmm");
        let baseFilename = `Email: ${fileDate}`;
        console.log(`Exporting thread '${subject}'`);

        messages.forEach( message => {   

          let body = message.getBody();

          html += `<dl class="email-meta">
          <dt class="subject">Subject:</dt> <dd class="subject">${message.getSubject()}</dd>
          <dt>From:</dt><dd class="strong">${formatEmails(message.getFrom())}</dd>
          <dt>Date:</dt><dd>${Utilities.formatDate(message.getDate(), tz, "MMMMM dd, yyyy 'at' h:mm a ")}</dd>
          <dt>To:</dt><dd>${formatEmails(message.getTo())}</dd>`;
          
          if ( message.getCc().length > 0 ) {
            html += `<dt>cc:</dt> <dd>${formatEmails(message.getCc())}</dd>`;
          }

          if ( message.getBcc().length > 0 ) {
            html += `<dt>bcc:</dt> <dd>${formatEmails(message.getBcc())}</dd>`;
          }

          html += '</dl>';
          body = embedHtmlImages(body);          
          body = embedInlineImages(body, message.getRawContent());
          if ( _exportAttachments ) {
            const atts = message.getAttachments();
            if ( atts.length > 0 ) {
              body += `<br />\n<strong>Attachments:</strong>
                        <div class="email-attachments">`;
              atts.forEach( att => {
                const attFilename = att.getName();
                let imageData;
                if ( imageData = renderDataUri(att) ) {
                  body += `<img src="${imageData}" alt="&lt;${attFilename}&gt;" /><br />`;
                } else {
                  body += `&lt;${attFilename}&gt;<br />`;
                }
                attachments.push(att);
              });
              body += `</div>`;
            }
          }
          html += body;        
        });

        html += '</body></html>';
        const filename = `${baseFilename}  ${subject}`;
        const threadHtmlBlob = Utilities.newBlob(html, 'text/html', filename);
        if ( _exportAttachments ) {
          if ( attachments.length > 0 ) {
            attachments.forEach( att => {
              const attname = att.getName();
              if ( ! _excludeAttachments.includes(attname.toUpperCase() ) ) {
                const newName = `${baseFilename} ATTACHMENT: ${attname}`;
                const files = folder.getFilesByName(newName);
                const file = files.hasNext() ?
                            files.next() :
                            folder.createFile(att);
                file.setName(newName);
              }              
            });
          }
        }

        const existing = folder.getFilesByName(`${filename}.pdf`);

        if ( existing.hasNext() ) existing.next().setTrashed(true);

        folder.createFile(threadHtmlBlob.getAs('application/pdf'))
        .setName(`${filename}.pdf`);

        console.log(`└─File exported successfully: ${filename}`);
        thread.addLabel(getOrCreateLabel(_doneLabel))
        .removeLabel(getOrCreateLabel(_exportLabel))
        .removeLabel(getOrCreateLabel(_reviewLabel));
        ScriptStatus.increment();
        currentThread = null;


      });
      
      return maxThreads === threads.length;
    }
    console.log("No emails found with label '%s'", _exportLabel);
    return false;

  } catch (err) {
    console.log(`ERROR: ${err.message}`);
    if ( _emailOnError ) {
      const emailOpts = {
        htmlBody: `<p style="font-family:Roboto,Arial;font-size:11pt;font-weight:bold;">
        An error occurred when attempting to export an email to PDF.</p>
        <p style="font-family:'Roboto Mono',Consolas,'Courier New';font-size:10pt;max-width:80em;background-color:#fce8e6;color:#b80e22;padding:1em;border-radius:0.5em;">${err.message}</p>`,
        to: Session.getActiveUser().getEmail(),
        subject: `Export error`,
        noReply: true,
        name: "Gmail Export Script",
      }
      if ( currentThread !== null ) {
        const errThread = GmailApp.getThreadById(currentThread);
        errThread.addLabel(getOrCreateLabel(_errorLabel));
        const threadDate = Utilities.formatDate(errThread.getLastMessageDate(), Session.getScriptTimeZone(), "MM/dd/yyyy");
        emailOpts.htmlBody += `<p style="font-family:Roboto,Arial;font-size:11pt;">
        ${threadDate} &mdash; <a style="color:#1155cc"
        href="https://mail.google.com/mail/u/0/#inbox/${currentThread}">
        ${errThread.getFirstMessageSubject()}</a></p><br><br>
        <p style="font-family:Roboto,Arial;font-size:11pt;">To save the email thread,
        you can print the message to the "Save to Google Drive" printer and save individual
        attachments if needed.</p>
        <p style="font-family:Roboto,Arial;font-size:11pt;"><i>The </i>${getLabelSpan(_errorLabel)}
        <i> label has been added to this thread.</i></p>`;
      }
      MailApp.sendEmail(emailOpts);
    }
    return true;    
  }
}

/**
 * Get a user label, creating it if it doesn't exist
 * @param {String} labelName - Name of the label
 * @returns {GmailApp.GmailLabel}
 */
function getOrCreateLabel(labelName) {
  if ( labelName.length === 0 ) return;
  let label = GmailApp.getUserLabelByName(labelName);
  if ( !label ) {
    label = GmailApp.createLabel(labelName);
  }
  return label

}

/**
 * Download and embed all images referenced within an html document as data uris
 *
 * @param   {String} html
 * @returns {String}        Html with embedded images
 */
function embedHtmlImages(html) {
  html = processImageTags(html);
  html = processStyleAttributes(html);
  html = processStyleTags(html);
  return html;
}
/**
 * Download and embed all img tags
 *
 * @param   {String} html
 * @returns {String}        Html with embedded images
 */
function processImageTags(html){
  return html.replace(
    /(<img[^>]+src=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi,
    function(m, tag, q, src) {
      return tag + q + (renderDataUri(src) || src) + q;
    }
  );
}

/**
 * Download and embed all HTML Style Attributes
 *
 * @param   {String} html
 * @returns {String}        Html with embedded style attributes
 */
function processStyleAttributes(html){
  return html.replace(
    /(<[^>]+style=)(["'])((?:(?!\2)[^\\]|\\.)*)\2/gi,
    function(m, tag, q, style) {
      style = style.replace(
        /url\((\\?["']?)([^\)]*)\1\)/gi,
        function(m, q, url) {
          return 'url(' + q + (renderDataUri(url) || url) + q + ')';
        }
      );
      return tag + q + style + q;
    }
  );
}

/**
 * Download and embed all HTML Style Tags
 *
 * @param   {String} html
 * @returns {String}        Html with embedded style tags
 */
function processStyleTags(html){
  return html.replace(
    /(<style[^>]*>)(.*?)(?:<\/style>)/gi,
    function(m, tag, style, end) {
      style = style.replace(
        /url\((["']?)([^\)]*)\1\)/gi,
        function(m, q, url) {
          return 'url(' + q + (renderDataUri(url) || url) + q + ')';
        }
      );
      return tag + style + end;
    }
  );
}

/**
 * Extract and embed all base64-encoded inline images
 *
 * @param   {String} html - Message body
 * @param   {String} raw  - Unformatted message contents
 * @returns {String}      - Html with embedded images
 */
function embedInlineImages(html, raw) {
  const images = [];

  // locate all inline content ids
  raw.replace(
    /<img[^>]+src=(?:3D)?(["'])cid:((?:(?!\1)[^\\]|\\.)*)\1/gi,
    function(m, q, cid) {
      images.push(cid);
      return m;
    }
  );

  // extract all inline images
  images.forEach( cid => {
    let cidIndex = raw.search(new RegExp("Content-ID ?:.*?" + cid, 'i'));
    if (cidIndex === -1) return null;
    let prevBoundaryIndex = raw.lastIndexOf("\r\n--", cidIndex);
    let nextBoundaryIndex = raw.indexOf("\r\n--", prevBoundaryIndex+1);
    let part = raw.substring(prevBoundaryIndex, nextBoundaryIndex);
    let encodingLine = part.match(/Content-Transfer-Encoding:.*?\r\n/i)[0];
    let encoding = encodingLine.split(":")[1].trim();
    if (encoding != "base64") return null;
    let contentTypeLine = part.match(/Content-Type:.*?\r\n/i)[0];
    let contentType = contentTypeLine.split(":")[1].split(";")[0].trim();
    let startOfBlob = part.indexOf("\r\n\r\n");
    let blobText = part.substring(startOfBlob).replace("\r\n","");
    const myRegex = new RegExp(
      '<img[^>]+src=(?:3D)?(")cid:' + cid + '\\1',
      "gi"
    );
    html = html.replace(
      myRegex,
      `<img src="data:${contentType};base64, ${blobText}"`
    );
  })
  // process all img tags which reference "attachments"
  return processImgAttachments(html);
}

/**
 * Download and embed all HTML Inline Image Attachments
 *
 * @param   {String} html
 * @returns {String}        Html with inline image attachments
 */
function processImgAttachments(html){
  return html.replace(
    /(<img[^>]+src=)(["'])(\?view=att(?:(?!\2)[^\\]|\\.)*)\2/gi,
    function(m, tag, q, src) {
      return tag + q + (renderDataUri(images.shift()) || src) + q;
    }
  );
}

/**
 * Convert an image into a base64-encoded data uri.
 *
 * @param   {Blob|string} Blob - object containing an image file or
 *                               a remote url string
 * @returns {string}           - Data uri
 */
function renderDataUri(image) {
  if ( typeof image == 'string' &&
       !(isValidUrl(image) &&
         (image = fetchRemoteFile(image))
        ) ) {
    return null;
  }
  if (isa_(image, 'Blob') || isa_(image, 'GmailAttachment')) {
    if (image.getContentType() != null) {
      var type = image.getContentType().toLowerCase();
      var data = Utilities.base64Encode(image.getBytes());
      if (type.indexOf('image') == 0) {
        return 'data:' + type + ';base64,' + data;
      }
    } 
  }
  return null;
}

/**
 * Fetch a remote file and return as a Blob object on success
 *
 * @param   {String} url
 * @returns {Blob}
 */
function fetchRemoteFile(url) {
  try {
    var response = UrlFetchApp.fetch(url, {'muteHttpExceptions': true});
    return response.getResponseCode() == 200 ? response.getBlob() : null;
  } catch (e) {
    return null;
  }
}

/**
 * Validate a url string (taken from jQuery)
 *
 * @param   {String}  url
 * @returns {Boolean}
 */
function isValidUrl(url) {
  return /^(https?|ftp):\/\/(((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:)*@)?(((\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5])\.(\d|[1-9]\d|1\d\d|2[0-4]\d|25[0-5]))|((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)(:\d*)?)(\/((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)+(\/(([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)*)*)?)?(\?((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|[\uE000-\uF8FF]|\/|\?)*)?(\#((([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(%[\da-f]{2})|[!\$&'\(\)\*\+,;=]|:|@)|\/|\?)*)?$/i.test(url);
}

/**
 * Turn emails of the form "<handle@domain.tld>" into 'mailto:' links.
 *
 * @param   {String} emails
 * @returns {String}
 */
function formatEmails(emails) {
  var pattern = new RegExp(/<(((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.?)>/i);
  return emails.replace(pattern, function(match, handle) {
    return '<a href="mailto:' + handle + '">' + handle + '</a>';
  });
}

const cleanUp = () => {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if ( trigger.getHandlerFunction() === 'resumeExport') {
      ScriptApp.deleteTrigger(trigger)
    }    
  });
}

const Trigger = (function () {
  class Trigger {
    constructor(functionName, everyMinutes) {
      return ScriptApp.newTrigger(functionName)
        .timeBased()
        .everyMinutes(everyMinutes)
        .create();
    }

    static deleteTrigger(e) {
      if (typeof e !== 'object')
        return console.log(`${e} is not an event object`);
      if (!e.triggerUid)
        return console.log(`${JSON.stringify(e)} doesn't have a triggerUid`);
      ScriptApp.getProjectTriggers().forEach(trigger => {
        if (trigger.getUniqueId() === e.triggerUid) {
          console.log('deleting trigger', e.triggerUid);
          return ScriptApp.deleteTrigger(trigger);
        }
      });
    }
  }
  return Trigger;
})();

class ScriptStatus {
  /**
   * Set a script property
   * @param {String|Number}  val    - Value to set
   * @param {String}         [prop] - Name of property. Defaults to 'status'
   */
  static set(val, prop='status') {
    return PropertiesService.getScriptProperties().setProperty(prop, val);
  }
  /**
   * Get a script property.
   * @param {String} [prop] - Name of property. Defaults to 'status'
   */
  static get(prop='status') {
    return PropertiesService.getScriptProperties().getProperty(prop);
  }

  /** Add 1 to the '_count' script property */
  static increment() {
    let val = PropertiesService.getScriptProperties().getProperty('_count');
    if ( !val ) val = 0;
    val = parseFloat(val);
    val += 1;
    PropertiesService.getScriptProperties().setProperty('_count', val);
    return val;
  }

  static count() {
    let val = PropertiesService.getScriptProperties().getProperty('_count');
    if ( !val ) val = 0;
    PropertiesService.getScriptProperties().setProperty('_count', 0);
    return val;
  }
  
  /**
   * Delete script properties
   * @param {String[]} props - Property names to delete
   */
  static remove(props){
    props.forEach(prop => {
      PropertiesService.getScriptProperties().deleteProperty(prop);
    })
  }
}

const getLabelSpan = content => {
  return `<span style="padding:0.1em 0.3em;background:#dedede;border-radius:0.5em;white-space:nowrap;">${content}</span>`;
}


/**
 * Test class name for Google Apps Script objects. They have no constructors
 * so we must test them with toString.
 *
 * @param   {Object}   obj
 * @param   {String}   class
 * @returns {Boolean}
 */
function isa_(obj, myClass) {
  return typeof obj == 'object' && typeof obj.constructor == 'undefined' && obj.toString() == myClass;
}

/**
 * Replace characters in a label name to use for search
 * @param {String} lbl  - Label display name
 * @returns {String}    - Label name ready to use in search criteria
 */
const cleanLabel = lbl => lbl.toLowerCase().replace(/[\/ ]/g,'-');
