/**
 * Home Maintenance Reminder Script
 * --------------------------------
 * Reads a Google Sheet with home maintenance tasks and sends email reminders
 * for upcoming or overdue items.
 * 
 * Ideal for beginner/intermediate scripters looking to automate household task reminders.
 * Note that authorizing this script ONLY authorizing for your user on your file. 
 * 
 * Setup: Set your email in USER_CONFIG below, then run main() to test.
 *    First time, you will need to grant permissions to this script.  See instructions in README.
 * 
 *    When confirmed working, run setupSchedule() to automate the schedule.
 */

const USER_CONFIG = {
  email: "youremail@gmail.com"  //TODO SET YOUR EMAIL HERE IN QUOTES - required for script to work, e.g `email: "yourname@gmail.com"`
};

//Feel free to customize any of these config settings to your taste. These are smart defaults.
const DEFAULT_CONFIG = {
  sheetName: "house_log", // name of sheet containing maintenance data
  sheetRange: "A1:L100",  // range to read from maintenance sheet
  triggerDays: ["SATURDAY"], // comma-separated day(s) of week to run e.g. ["MONDAY","TUESDAY"], etc.; defaults to every Saturday
  triggerHours: [7], // comma-separated hour(s) to run on the triggerDays, e.g. [7,14] (0-23, where 7 = 7am, 14 = 2pm); defaults to 7am
  
  daysDueThreshold: 7, // items due within this many days will trigger reminder
  daysOverdueThreshold: 14, // items overdue by this many days or more are flagged as overdue
  
  emailSubject: "🏠House maintenance due: {count} item(s)", // email subject template ({count} gets auto-filled)
  emailFooter: "\nSee instructions and update Last Done date when complete in {url}.\n\n Thanks for the house love. 🔧❤️🔧", // footer template ({url} gets autofilled)
  emailBodyHeader: "⚠️The following maintenance is {status}:", // header for each section ({status} gets autofilled w/ DUE or OVERDUE)
  dueLabels: { // templates for due date descriptions
    today: "due today",
    future: "due in {days} days", // {days} replaced with number of days
    past: "due {days} days ago" // {days} replaced with number of days
  },
  statusLabels: { // section headers
    due: "DUE",
    overdue: "OVERDUE"
  },
  columns: { // column header names to look for (case-insensitive partial match)
    maintenance: "Maintenance",
    due: "Next due",
    frequency: "Frequency",
    archived: "Archived"
  }
};
// ======= END CONFIG ======= //

/* Main logic - check for due/overdue house maintenance and send email for anything due */
function main(){
  try {
    loadConfig();
    readMaintenanceLog();
  } catch (error){
    sendEmailError(error);
  }
}

/* MANUAL - Set up scheduled triggers based on config - run once to set up or change the automated schedule */
function setupSchedule() {
  loadConfig();
  deleteTriggers('main');//clear the old settings
  
  for (const day of CONFIG.triggerDays) {
    for (const hour of CONFIG.triggerHours) {
      ScriptApp.newTrigger('main')
        .timeBased()
        .onWeekDay(ScriptApp.WeekDay[day])
        .atHour(hour)
        .create();
      Logger.log(`Set up schedule for ${day} at ${hour}:00`);
    }
  }
}

let CONFIG = null;

/* Load and merge configuration from defaults and user config 
 * @returns {Object} The merged configuration object
 */
function loadConfig() {
  if (CONFIG) return CONFIG;
  
  CONFIG = {...DEFAULT_CONFIG, ...USER_CONFIG};
  
  if (!CONFIG.email || isFakeEmail()) {
    sendEmailError("Email must be set in USER_CONFIG before running. (Replace sample email.)",true);
  }
  
  return CONFIG;
}

/* Delete all existing triggers for specified function 
 * @param {string} triggerFunctionName - Name of function to delete triggers for
 */
function deleteTriggers(triggerFunctionName){ 
  const allTriggers = ScriptApp.getProjectTriggers();
  for (const trigger of allTriggers) {
    if (trigger.getHandlerFunction() === triggerFunctionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted prior schedule for ${triggerFunctionName}`);
    }
  }
}

/* A small container object for each item in our sheet 
 * @param {string} desc - Description of the maintenance task
 * @param {Date|string} dueDate - Due date for the task
 */
class MaintenanceItem {
  constructor(desc, dueDate) {
    this.desc = desc;
    this.dueDate = dueDate;
    this.today = new Date();
    this.today.setHours(0,0,0,0);//start of day, hour wise

  }
  /* Return number of days til due. <0 is overdue
   * @returns {Integer} Days until due (negative if overdue)
   */
  daysTilDue(){
    return getDaysBetween(this.today, new Date(this.dueDate));
  }
  /* Check if maintenance item is due within threshold 
   * @returns {boolean} True if item is due
   */
  isDue(){
    return (this.daysTilDue() <= CONFIG.daysDueThreshold)
  }
  /* Check if maintenance item is overdue beyond threshold 
   * @returns {boolean} True if item is overdue
   */
  isOverdue(){
    return (this.daysTilDue() <= -1*CONFIG.daysOverdueThreshold)
  }
  /* Text description of when this is due
   * @returns {string} Formatted due date description
   */
  dueLabel(){
    const dueDays = this.daysTilDue();
    if (dueDays === 0) {
      return CONFIG.dueLabels.today;
    } else if (dueDays >= 1) {
      return CONFIG.dueLabels.future.replace("{days}", dueDays.toString());
    } else {
      return CONFIG.dueLabels.past.replace("{days}", Math.abs(dueDays).toString());
    }
  }
}

/* Construct body of email to send 
 * @param {MaintenanceItem[]} items - Array of maintenance items
 * @param {boolean} overdue - Whether this section is for overdue items
 * @returns {string} Formatted email body section
 */
function buildMailBody(items, overdue){
  if (items.length === 0) return '';

  const status = overdue ? CONFIG.statusLabels.overdue : CONFIG.statusLabels.due;
  const header = CONFIG.emailBodyHeader.replace("{status}", status);
  const itemList = items.map(obj => `- ${obj.desc} - ${obj.dueLabel()}`).join('\n');
  return `${header}

${itemList}

-----
`;
}

/* Build and send email for due and overdue items 
 * @param {MaintenanceItem[]} dueItems - Array of due maintenance items
 * @param {MaintenanceItem[]} overdueItems - Array of overdue maintenance items
 * @returns {void}
 */
function sendEmail(dueItems, overdueItems) {
  const emailTo = CONFIG.email;
  const numItems = dueItems.length + overdueItems.length
  if (numItems === 0) return;
  
  const subject = CONFIG.emailSubject.replace("{count}", numItems.toString());
  const body1 = buildMailBody(dueItems, false);
  const body2 = buildMailBody(overdueItems, true);
  const buffer = (body1 && body2) ? "\n" : ''
  
  const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const footer = CONFIG.emailFooter.replace("{url}", spreadsheetUrl);
  const body = body1 + buffer + body2 + footer;
  
  MailApp.sendEmail(emailTo, subject, body);
  Logger.log(subject)
  Logger.log(body)
}

/* Return true if user has not yet customized the email field, else false */
function isFakeEmail(){
  return USER_CONFIG.email === "youremail@gmail.com"
}

/* Send error notification email and optionally terminate script 
 * @param {Error|string} error - Error to report
 * @param {boolean} dieOnError - Whether to throw error after emailing (default: true)
 * @returns {void}
 */
function sendEmailError(error,dieOnError=true) {
  const emailTo = CONFIG?.email || USER_CONFIG.email;
  if (!emailTo) {

    console.error("Cannot send error email - no email configured");
    if (dieOnError) throw error;
    return;
  }
  const subject = "ERROR in Google script";
  const scriptId = ScriptApp.getScriptId();
  const link = `https://script.google.com/home/projects/${scriptId}/executions`
  const body = "Google script failed with error:\n"+error +"\n\n" + link;
  MailApp.sendEmail(emailTo, subject, body);
  if (dieOnError){
    throw error;
  }
}

/* Read sheet to parse due and overdue maintenance tasks 
 * @returns {void}
 */
function readMaintenanceLog(){
  const {now:start} = elapsedTime();
  Logger.log(`Starting maintenance log check`);
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) {
    sendEmailError(`Sheet "${CONFIG.sheetName}" not found`);
    return;
  }
  
  const sheetData = sheet.getRange(CONFIG.sheetRange).getValues();
  const headerRow = sheetData.shift();
  
  const iItem = findColumn(headerRow, CONFIG.columns.maintenance, CONFIG.sheetName);
  const iDueDate = findColumn(headerRow, CONFIG.columns.due, CONFIG.sheetName);
  const iArchived = findColumn(headerRow, CONFIG.columns.archived, CONFIG.sheetName);
  
  if (iItem === null || iDueDate === null) { // don't barf on 0 index
    sendEmailError(`Required columns not found in ${CONFIG.sheetName} sheet in first row of range ${CONFIG.sheetRange}`);
    return;
  }
  
  const overdueItems = [];
  const dueItems = [];
  let nEmpty = 0;
  
  for (const row of sheetData) {
    if (nEmpty === 4) break; // end of populated rows
    
    const desc = row[iItem];
    const due = row[iDueDate];
    const archived = row[iArchived];
    
    // Skip rows without task or due date
    if (!desc || !due) {
      nEmpty++;
      continue; // skip empty
    }
    
    if (archived) {
      continue; // skip archived tasks
    }
    
    const item = new MaintenanceItem(desc, due);
    if (item.isOverdue()) {
      overdueItems.push(item);
    } else if (item.isDue()) {
      dueItems.push(item);
    }
  }
  
  if (overdueItems.length > 0 || dueItems.length > 0) {
    // send email
    sendEmail(dueItems, overdueItems);
  } else {
    Logger.log("No tasks are currently due - Nice!")
  }
  
  const {elapsed} = elapsedTime(start);
  Logger.log(`Maintenance log check complete, elapsed: ${elapsed}`);
}

// ===== UTILS ===== 
const MS_PER_SECOND = 1000;
const MS_PER_MINUTE = MS_PER_SECOND * 60;
const MS_PER_HOUR = MS_PER_MINUTE * 60;
const MS_PER_DAY = MS_PER_HOUR * 24;

/* Figure elapsed time from a prior time to now, and return both now and elapsed time in data object 
 * @param {Integer} [sinceTime] - Optional start time in milliseconds
 * @returns {{now: Integer, elapsed: string|null}} Object with current time and elapsed time string
 */
function elapsedTime(sinceTime){
  //return object with now = current time, elapsed=string with elapsed time since optional sinceTime arg
  const now = (new Date()).getTime();//milliseconds
  let elapsed = null;
  if (sinceTime){
    let ms = now - sinceTime;
    if (ms >= MS_PER_MINUTE){
      elapsed = `${(ms/MS_PER_MINUTE).toFixed(1)} min`; // convert ms to min
    } else if (ms >= MS_PER_SECOND){
      elapsed = `${ms/MS_PER_SECOND} second`; // convert ms to seconds
    } else {
      elapsed = `${ms} ms`;
    } 
  }
  return {now:now,elapsed:elapsed}
}

/* Look for a match of a string in a row of data for a given sheet and return found column index number or null 
 * @param {Array} rowData - Array of header values to search
 * @param {string} label - Label to search for (case-insensitive substring match)
 * @param {string} [sheetName] - Optional sheet name for error logging
 * @returns {Integer|null} Column index if found, null if not found
 */
function findColumn(rowData,label,sheetName=null) {
  const regex = new RegExp(label, 'i')//case insensitive substring match
  const i = rowData.findIndex(item => regex.test(item));
  if (i < 0) {
    Logger.log("No column found matching label \""+label +"\"" + ((sheetName)? ` on sheet ${sheetName}` : ''));
    return null;
  }else {
    return i;
  }
}
/* Return number of days between 2 dates - calculate milliseconds between times then convert to days and return as int 
 * @param {Date} earlierDate - The earlier date
 * @param {Date} laterDate - The later date
 * @returns {Integer} Number of days between the dates (integer)
 */
function getDaysBetween(earlierDate, laterDate){
  const result = Math.round(laterDate.getTime() - earlierDate.getTime()) / MS_PER_DAY; 
  return Math.trunc(result); // convert to integer
}