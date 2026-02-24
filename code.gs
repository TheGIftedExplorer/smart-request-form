//CONFIGURATION
const COLS = {
  STATUS: "Approval Status",    
  REQ_EMAIL: "Email Address",   
  REQ_NAME: "Requester Name",   
  REQ_TYPE: "Request Type",     
  MANAGER_EMAIL: "Manager’s Email Address" 
};

const WEB_APP_URL = "https://script.google.com/macros/s/AKfycby_AqRcnAxx25ZczhO3sgI8NsBaqpBV3NDll0YpkXKu_XdCgqwCkQ3aCzjGPy8kzqvk/exec"; 

function doGet(e) {
  try {
  if (!e || !e.parameter || !e.parameter.row) {
    return HtmlService.createHtmlOutput("Invalid Request Parameters.");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0];
  const row = Number(e.parameter.row);
  const action = e.parameter.action;

  //We need to read the headers to know which column is which
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap = getColumnMap(headers);
  const statusColIndex = colMap[COLS.STATUS];
  const currentStatus = sheet.getRange(row, statusColIndex + 1).getValue();

  //If an approval link is clicked multiple times, display "Already Processed"
  if (currentStatus !== "") {
    return renderMessage("Already Processed", `This was already <strong>${currentStatus}</strong>.`, "warning");
  }
  //action stores the status that has to be displayed in the webApp
  sheet.getRange(row, statusColIndex + 1).setValue(action);
  const decisionDate = new Date().toLocaleDateString(); 

  //get data for email
  const email = sheet.getRange(row, colMap[COLS.REQ_EMAIL] + 1).getValue();
  const type  = sheet.getRange(row, colMap[COLS.REQ_TYPE] + 1).getValue();
  const name  = sheet.getRange(row, colMap[COLS.REQ_NAME] + 1).getValue();
  const managerEmail  = sheet.getRange(row, colMap[COLS.MANAGER_EMAIL] + 1).getValue();

  //Generate and send HTML email to the requester using template
  const emailHtml = createEmailTemplate(name, type, action, decisionDate, managerEmail);
  GmailApp.sendEmail(email, "Update: " + type + " Request", "Your email client does not support HTML.", {
    htmlBody: emailHtml,
    name: "Request Handler"
  });
  //Display status in webApp
  return renderMessage(action, `The request has been <strong>${action}</strong>.`, action);
}catch (err) {
    // Log the error
    console.error("doGet Error: " + err.toString());
    return renderMessage("Error", "The system encountered an issue. Please try again or contact admin.", "warning");
  }

}

//Triggered function on clicking submit in the form
function onFormSubmit(e) {
  const row = e.range.rowStart;
  const responses = e.namedValues; 
  //Get data from event responses
  const managerEmail = responses[COLS.MANAGER_EMAIL] ? responses[COLS.MANAGER_EMAIL][0] : ""; 
  const type = responses['Request Type'] ? responses['Request Type'][0] : "General";
  const priority = responses['Priority'] ? responses['Priority'][0] : "Normal";
  const desc = responses['Request Description'] ? responses['Request Description'][0] : "";
  const requester = responses['Requester Name'] ? responses['Requester Name'][0] : "An employee";
  const attachmentLink = responses['Supporting Attachment(Optional)'] ? responses['Supporting Attachment(Optional)'][0] : "";
  const startDate = responses['Leave Start Date'] ? responses['Leave Start Date'][0] : "";
  const endDate   = responses['Leave End Date'] ? responses['Leave End Date'][0] : "";
  
  const approveLink = WEB_APP_URL + "?action=Approved&row=" + row;
  const rejectLink  = WEB_APP_URL + "?action=Rejected&row=" + row;

  if (!managerEmail) {
    console.log("Error: No Manager Email found.");
    return;
  }
  //Generate mail to manager sending request details with approval links
  const managerHtml = createManagerEmailTemplate(requester, type, desc, priority, approveLink, rejectLink, attachmentLink, startDate, endDate);
  GmailApp.sendEmail(managerEmail, `[${priority}] New ${type} Request from ${requester}`, "Please use an HTML-compatible email client.", {
    htmlBody: managerHtml,
    name: "Request Handler"
  });
}

//HELPER FUNCTION used in doGet method to get data using column names
function getColumnMap(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[header] = index; 
  });
  return map;
}

//HTML for webApp that displays status of approval using HtmlService
function renderMessage(title, message, type) {
  const colors = { "Approved": "#28a745", "Rejected": "#dc3545", "warning": "#ffc107", "default": "#007bff" };
  const t = HtmlService.createTemplateFromFile('WebApp');
  t.title = title;
  t.message = message;
  t.themeColor = colors[type] || colors["default"];
  //Return the evaluated HTML
  return t.evaluate().setTitle("Request System");
}

//HTML and styling of generated mail for requester
function createEmailTemplate(name, type, action, date, managerEmail) {
  const isApproved = action === "Approved";
  const t = HtmlService.createTemplateFromFile('EmployeeEmail');
  t.name = name;
  t.type = type;
  t.action = action;
  t.date = date;
  t.managerEmail = managerEmail;
  t.headerColor = isApproved ? "#28a745" : "#dc3545";
  t.statusText = isApproved ? "APPROVED" : "REJECTED";
  
  return t.evaluate().getContent();
}

//HTML and styling of generated mail for manager
function createManagerEmailTemplate(requesterName, type, desc, priority, approveUrl, rejectUrl, attachmentLink, start, end) {
  const t = HtmlService.createTemplateFromFile('ManagerEmail');
  t.requesterName = requesterName;
  t.type = type;
  t.desc = desc;
  t.priority = priority;
  t.approveUrl = approveUrl;
  t.rejectUrl = rejectUrl;
  t.priorityColor = priority === "Urgent" ? "#dc3545" : "#6c757d";

  t.dateHtml = (type.toLowerCase().includes("leave") && start && end) 
    ? `<div style="padding:10px;background:#e8f5e9;color:#2e7d32;margin-bottom:15px;border:1px dashed #2e7d32;"><strong>Dates:</strong> ${start} - ${end}</div>` 
    : "";

  t.attachmentSection = attachmentLink ? `
    <div style="margin: 20px 0; padding: 10px; background-color: #f0f7ff; border-radius: 5px; text-align: center;">
      <p style="margin: 0 0 10px 0; font-size: 14px; color: #555;">Review Attached Supporting Document:</p>
      <a href="${attachmentLink}" style="color: #007bff; text-decoration: underline; font-weight: bold;">View Attachment</a>
    </div>
  ` : "";

  return t.evaluate().getContent();
}