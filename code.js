/**
* The following script retrieves attachments from emails under a certain label.
* it will store the attachment in Drive under a path derived from the email subject
* in addition, it will log the transaction on the current sheet
*
* To use, you must first set a filter in the gmail account to create the label based on your conditions
* currently, the filter should be something like this,
*  (submit hw) {any text}  @studentName#studentID
* if the subject is formated like above, it will capture the hw assignment as the string before @ 
* the studentName, and the studentId
*
*
* modified code found : Andreas Gohr and splitbrain.org.  I added the ability to log to a spreadsheet
*/

/**
* Const needed.  Label to match gmail setting, desired path. 
*/
var GMAIL_LABEL = 'submitted-hw';
var GDRIVE_FILE = 'studentHw/$y/$id/$sublabel/$y-$m-$d_$student_$hw_$mc_$ac.$ext';


/**
 * Get all the starred threads within our label and process their attachments
 */

function main() {
  var labels = getSubLabels(GMAIL_LABEL);
  for(var i=0; i<labels.length; i++) {
    var threads = getUnprocessedThreads(labels[i]);
    for(j=0; j<threads.length; j++) {
      processThread(threads[j], labels[i]);
    }
  }
}

/**
 * Returns the Google Drive folder object matching the given path
 *
 * Creates the path if it doesn't exist, yet.
 *
 * @param {string} path
 * @return {Folder}
 */
function getOrMakeFolder(path) {
  var folder = DriveApp.getRootFolder();
  var names = path.split('/');
  while(names.length) {
    var name = names.shift();
    if(name === '') continue;
    
    var folders = folder.getFoldersByName(name);
    if(folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = folder.createFolder(name);
    }
  }
  
  return folder;
}

/**
 * Get all the given label and all its sub labels
 *
 * @param {string} name
 * @return {GmailLabel[]}
 */
function getSubLabels(name) {
  var labels = GmailApp.getUserLabels();
  var matches = [];
  for(var i=0; i<labels.length; i++){
    var label = labels[i];
    if(
      label.getName() === name ||
      label.getName().substr(0, name.length+1) === name+'/'
    ) {
      matches.push(label);
    }
  }
  
  return matches;
}

/**
 * Get all starred threads in the given label
 *
 * @param {GmailLabel} label
 * @return {GmailThread[]}
 */
function getUnprocessedThreads(label) {
  var from = 0;
  var perrun = 50; //maximum is 500
  var threads;
  var result = [];
  
  do {
    threads = label.getThreads(from, perrun);
    from += perrun;
    
    for(var i=0; i<threads.length; i++) {
      if(!threads[i].hasStarredMessages()) continue;
      result.push(threads[i]);
    }
  } while (threads.length === perrun);
  
  Logger.log(result.length + ' threads to process in ' + label.getName());
  return result;
}

/**
 * Get the extension of a file
 *
 * @param  {string} name
 * @return {string}
 */
function getExtension(name) {
  var re = /(?:\.([^.]+))?$/;
  return re.exec(name)[1].toLowerCase();
}

/**
 * Apply template vars
 *
 * @param {string} filename with template placeholders
 * @param {info} values to fill in
 * @param {string}
 */
function createFilename(filename, info) {
  var keys = Object.keys(info);
  keys.sort(function(a,b) {
    return b.length - a.length;
  });
  
  for(var i=0; i<keys.length; i++) {
    filename = filename.replace(new RegExp('\\$'+keys[i], 'g'), info[keys[i]]);
  }
  return filename;
}

/**
* Saves the attachment to drive
* @param {attachment} attachment
* @param {str}  path
*/
function saveAttachment(attachment, path) {
  var parts = path.split('/');
  var file = parts.pop();
  var path = parts.join('/');
  
  var folder = getOrMakeFolder(path);
  var check = folder.getFilesByName(file);
  if(check.hasNext()) {
    Logger.log(path + '/' + file + ' already exists. File not overwritten.');
    return;
  }
  folder.createFile(attachment).setName(file);
  Logger.log(path + '/' + file + ' saved.');
}

/**
 * @param {GmailThread} thread
 * @param {GmailLabel} label where this thread was found
 */
function processThread(thread, label) {
  var messages = thread.getMessages();
  for(var j=0; j<messages.length; j++) {
    var message = messages[j];
    if(!message.isStarred()) continue;
    Logger.log('processing message from '+message.getDate());
    // get the hw and student details from the email subject line
    Logger.log('subject ' + message.getSubject());
    var obj = splitNameAndId(message.getSubject());
    
    Logger.log(obj);
    // skip over if email subject not formatted correctly 
    if(!obj  || !obj.hw || !obj.id || !obj.name) continue;
    
    Logger.log('continue....');
    var attachments = message.getAttachments();
    for(var i=0; i<attachments.length; i++) {
      var attachment = attachments[i];
      var info = {
        'name': attachment.getName(),
        'ext': getExtension(attachment.getName()),
        'from': message.getFrom(), // domain part of email
        'sublabel': label.getName().substr(GMAIL_LABEL.length+1),
        'student': obj.name,
        'id': obj.id,
        'hw': obj.hw,
        'y': ('0000' + (message.getDate().getFullYear())).slice(-4),
        'm': ('00' + (message.getDate().getMonth()+1)).slice(-2),
        'd': ('00' + (message.getDate().getDate())).slice(-2),
        'h': ('00' + (message.getDate().getHours())).slice(-2),
        'i': ('00' + (message.getDate().getMinutes())).slice(-2),
        's': ('00' + (message.getDate().getSeconds())).slice(-2),
        'mc': j,
        'ac': i,
      }
      var file = createFilename(GDRIVE_FILE, info);
      var hwArray = [obj.name, obj.id, obj.hw, message.getDate(),  message.getFrom().replace(/^.+<([^>]+)>$/, "$1"), file];
      appendRow(hwArray);
      saveAttachment(attachment, file);
    }
    Logger.log('unstar...');
    message.unstar();
  }
  Logger.log('archive thread ');
  thread.markRead().moveToArchive();
}


/**
 * @param {GmailThread} thread
 * @return {Object}
 */
function splitNameAndId(str){
  try {
    var after = str.split("@");
    if (!after[1] ) return false;
    var nameId = after[1].split("#");
    return {'name': nameId[0], 'id': nameId[1], 'hw': after[0] };
    
  } catch(err) {
    Logger.log('splitNameAndId error '+ err);
  }
}

/**
* appends row to current sheet
*
* @param {array} hw obj 
*/
function appendRow(array){
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
// Appends a new row with 3 columns to the bottom of the
// spreadsheet containing the values in the array
    sheet.appendRow(array);
  } catch(err) {
    Logger.log('appendRow error '+ err);
  }
}
