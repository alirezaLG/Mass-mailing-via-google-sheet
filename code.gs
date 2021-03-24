function onInstall(e) {
  onOpen(e);
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
  
    ui.createMenu('Mail Sheet')
        .addItem('Send campaign...', 'sendCampaign')
    .addSeparator()
        .addItem('Check bounced', 'checkBounced')
        .addItem('Check unsubscribed', 'checkUnsubscribed')
      .addToUi();
}

function sendCampaign() {
  var html = doGet();
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function doGet(request) {
  return HtmlService.createTemplateFromFile('Campaign')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Mail Sheet')
      .setWidth(300);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getCampaigns() {
  var leadSheet = SpreadsheetApp.getActiveSheet();
  var leadRange = leadSheet.getRange(2, 4, leadSheet.getLastRow());
  var leadData = leadRange.getValues();
  
  var unique = getUnique(leadData);
  return unique;
}

function getTemplates() {
  var leadSheet = SpreadsheetApp.getActiveSheet();
  var leadRange = leadSheet.getRange(2, 3, leadSheet.getLastRow());
  var leadData = leadRange.getValues();
  
  var unique = getUnique(leadData);
  return unique;
}

function getUnique(a) {
    var seen = {};
    var out = [];
    var len = a.length;
    var j = 0;
    for(var i = 0; i < len; i++) {
         var item = a[i];
         if(seen[item] !== 1) {
               seen[item] = 1;
               out[j++] = item;
         }
    }
    return out;
}

function sendEmails(campaign, template, subject) {
  
  var email = Session.getActiveUser().getEmail();
  
  var leadSheet = SpreadsheetApp.getActiveSheet();
  var leadRange = leadSheet.getRange(2, 1, leadSheet.getLastRow(), leadSheet.getLastColumn());
  var leadData = leadRange.getValues();
  
  var logSent = Array();
  var logFailed = Array();

  var contacted = getDateString();
  var nextStatus = template; //+" Sent";
 
  for (var i = 0; ((i < leadData.length-1) && (logSent.length <= 1500)); ++i) {
    var row = leadData[i];
     
    sendEmail(row , subject);
    
    SpreadsheetApp.flush();
  }
  
  
  // Prepare email //this is confirmation email 
  var subject = campaign +" Send Emails";
  var body = "";
  // Sent
  if (logSent.length >= 1900)
    body += "You may send batches of 1900 emails. To continue, send it again.\n\n";
  if (logSent.length > 0) body += "- Emails sent:\n\n";
  for (var r = 0; r < logSent.length; ++r) {
    body += logSent[r] + "\n";
  }
  // Failed
  if (logFailed.length > 0) body += "\n\n" + "- Emails failed:\n\n";
  for (var r = 0; r < logFailed.length; ++r) {
    body += logFailed[r] + "\n";
  }
  
  if (logSent.length + logFailed.length == 0) {
    body = "No emails sent.";
  }
  
  //GmailApp.sendEmail(email, subject, body);
}



function base64_decoder(data){
  return Utilities.base64Encode(data ,Utilities.Charset.UTF_8);
}
function base64_fix(data){
  return '=?UTF-8?B?' + data + '?=';
}


function sendEmail(row , subject) {
  
  
  var myEmail = Session.getActiveUser().getEmail();
  
  var myName = base64_decoder( "👋 "+ row[10] +" سلام");
  var status = "T2"
  var email = row[12];
  subject = '✨ '+row[14] + ' ✨ ';
  

  //var doc = DocumentApp.openById(docId);
  var body = "";
  
  body = include('template.html');
  if (status != "T1") {
    
    var forScope = GmailApp.getInboxUnreadCount();
    var subject = Utilities.base64Encode(subject,Utilities.Charset.UTF_8 ); 

    
    var raw =
      'MIME-Version: 1.0' + '\r\n' +
      'From: ' + base64_fix(myName) + ' <' + myEmail + '>' + '\r\n' +
      'To: ' + email + '\r\n' +
      'Subject: ' + base64_fix(subject) + '\r\n' +
      'Content-Type: text/html; charset=UTF-8' + '\r\n' +
      '\r\n'+ body ;
    
    var eBody = Utilities.base64EncodeWebSafe(raw, Utilities.Charset.UTF_8);
    
    
    var params = { method: "post",
                   contentType: "application/json",
                   headers: {
                     "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
                   },
                   payload: JSON.stringify({
                    "userId": "me",
                    "raw": eBody,
                   // "threadId": threadId
                   })
                 };
    
    var resp = UrlFetchApp.fetch("https://www.googleapis.com/gmail/v1/users/me/messages/send", params);
    
  } else {
    // First email
    GmailApp.sendEmail(email, subject, body, { name: myName });
  }
}

function getDateString() {
  var date = new Date();
  var year = date.getYear();
  // getMonth starts at 0
  var month = date.getMonth() + 1;
  var day = date.getDate();
  var hour = date.getHours();
  var minute = date.getMinutes();
  var currentDate = hour + ":" + minute + "-" + + month + "/" + day + "/" + year;
  Logger.log(currentDate);
  return currentDate;
}



function checkBounced() {
  var myEmail = Session.getActiveUser().getEmail(); 
  
  var leadSheet = SpreadsheetApp.getActiveSheet();
  var leadRange = leadSheet.getRange(2, 1, leadSheet.getLastRow(), leadSheet.getLastColumn());
  var leadData = leadRange.getValues();
  
  var leadEmails = Array();
  var bouncedEmails = Array();
  var logBounced = Array();
  
  var bounced = GmailApp.search("from:mailer-daemon OR from:postmaster ", 0, 500);
//Logger.log(bounced.length);
  
  for (var i = 0; i < bounced.length; i++) {
    var thread = bounced[i];
    var messages = thread.getMessages();
    var message = messages[0];
    thread.markRead();
    if (message.getFrom() == myEmail) {
      Logger.log("hello");
      bouncedEmails.push(message.getTo() ); //sent emails 
    } else {
      
      var body = (messages[1]) ? messages[1].getBody() : messages[0].getBody();

      var at = body.indexOf("@");
      
      var start = body.lastIndexOf(" ", at);
      var last = body.indexOf(" ", at);
      var toEmail = body.slice(start, last);

      var greater = toEmail.indexOf(">");
      if (greater > 0) toEmail = toEmail.slice(greater+1);
      var colon = toEmail.indexOf(":");
      if (colon > 0) toEmail = toEmail.slice(colon+1);
      var semiColon = toEmail.indexOf(";");
      if (semiColon > 0) toEmail = toEmail.slice(semiColon+1);
    
      var less = toEmail.indexOf("<");
      if (less > 0) toEmail = toEmail.slice(0,less);
      var parenthesis = toEmail.indexOf('"');
      if (parenthesis > 0) toEmail = toEmail.slice(0,parenthesis);
      

      bouncedEmails.push(toEmail.replace(/<[^>]+>/g, ""));
    }
  }
//  Logger.log(bouncedEmails);
  //bouncedEmails // [<b>alirhe@ketabcha.com</b></a>, <b>alireza@ketabcha.com</b></a>]
  for (var i = 0; i < leadData.length-1; ++i) {
    var row = leadData[i];
    var status = row[2];
    var email = row[8];
    
    
      var bounced = bouncedEmails.indexOf(email);
//      Logger.log(email);
//      Logger.log(bounced);
      if (bounced >= 0) {
        leadSheet.getRange(2 + i, 3).setValue("BOUNCED");
        logBounced.push(email);
      }
    
  }
  // Prepare email
  var subject = "Sites.af  Check Bounced Emails";
  var body = "";
  // Bounced Emails
  if (logBounced.length > 0) body += "- Bounced emails:\n\n";
  for (var r = 0; r < logBounced.length; ++r) {
    body += logBounced[r] + "\n";
  }
  
  if (logBounced.length == 0)
    body = "No new bounced emails found.";
  
GmailApp.sendEmail(myEmail, subject, body);
SpreadsheetApp.flush();
}

function checkUnsubscribed() {
  var myEmail = Session.getActiveUser().getEmail();
  
  var leadSheet = SpreadsheetApp.getActiveSheet();
  var leadRange = leadSheet.getRange(2, 1, leadSheet.getLastRow(), leadSheet.getLastColumn());
  var leadData = leadRange.getValues();
  
  var unsubSheet = SpreadsheetApp.openById("1U7C292jpMnN6geoq5FFgH_95Ed3tfuhw4jDip-tF_bk").getSheets()[0];
  var unsubRange = unsubSheet.getRange(1, 1, unsubSheet.getLastRow(), unsubSheet.getLastColumn());
  var unsubData = unsubRange.getValues();
  
//  Logger.log(unsubData.length);
//  Logger.log(leadData);
  
  var leadEmails = Array();
  var unsubEmails = Array();
  var unsubIndex = Array();
  var logUnsub = Array();
  
  for (var i = 0; i < unsubData.length; ++i) {
    var row = unsubData[i];
    var email = row[2]; //in unsubscribe page define email column (start with 0)
    var status = row[4]; //
    
//    Logger.log(status);
    
    if (status != "UNSUBBED") {
      unsubEmails.push(email.toLowerCase());
      unsubIndex.push(i);
//      Logger.log("hello");
    }
    
  }
    Logger.log(leadData);
  for (var i = 0; i < leadData.length-1; ++i) {
    var row = leadData[i];
    var status = row[2];
    var email = row[8];
    
    if (!(status == "UNSUBBED")) {
      var unsub = unsubEmails.indexOf(email.toLowerCase())
      if (unsub > 0) {
        var unsubIdx = unsubIndex[unsub];
        unsubSheet.getRange(1 + unsubIdx, 5).setValue("UNSUBBED"); //col 5 is unsubscribe columns
        leadSheet.getRange(2 + i, 3).setValue("UNSUBBED"); // col 3 is statue columns
        logUnsub.push(email);
      }
    }
  }
  // Prepare email
  var subject = "Sites.af  Check Unsubscribed";
  var body = "";
  // Unsusbscribed
  if (logUnsub.length > 0) body += "- Unsubscribed:\n\n";
  for (var r = 0; r < logUnsub.length; ++r) {
    body += logUnsub[r] + "\n";
  }
  
  if (logUnsub.length == 0)
    body = "No new unsubscribed.";
  
  GmailApp.sendEmail(myEmail, subject, body);
  SpreadsheetApp.flush();
}

