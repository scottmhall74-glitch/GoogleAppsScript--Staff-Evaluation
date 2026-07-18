var eval_info = '1hkNcwLNWBHiKY8OkRUpMBBjubw-WYKea70quwvu_Nqk';
//var staffEval ='16Oy4-MO9EujWmv7YaQkOgtNy4X1qw5Plmh2bz_hjjpM';

var staffEval = '1tb9HazI7NQ39esGCYktqR7KgW_F_eMfDLl9__63jW30';
var supertab = 'Supervisors';
var email = Session.getActiveUser().getEmail();


function doGet() {
 
  var cond = '';
  var para = PropertiesService.getUserProperties();
  var cache = CacheService.getUserCache();
  cond = isSuper();
  
  if (cond && !cache.get(email + "New") && !para.getProperty(email + 'ID') && !para.getProperty(email + 'User')) {
    //Admin Index
   var html = HtmlService.createTemplateFromFile('index').evaluate().setTitle('Teacher Eval Form').setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    para.deleteProperty(email+"Access");
  }
  else if (cache.get(email + "New") == "New" && cond) {
    //get ID and set it into user properties
    //New Eval
    var html = HtmlService.createTemplateFromFile('NewEvaluation').evaluate().setTitle('New Teacher Eval').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    para.deleteProperty(email+"Access");
    cache.remove(email + "New")
  }
   
  else if (para.getProperty(email + 'ID') && cond) {
    // var html = HtmlService.createTemplateFromFile('Print').evaluate().setTitle('Teacher Eval Form').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    //Edit an Evaluation
    var html = HtmlService.createTemplateFromFile('ViewEval').evaluate().setTitle('Teacher Eval').setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    para.deleteProperty(email+"Access");
  }
  else if (para.getProperty(email + 'User')) {
    var html = HtmlService.createTemplateFromFile('StaffEvals').evaluate().setTitle('Teacher Evals').setSandboxMode(HtmlService.SandboxMode.IFRAME); 
  }
  else {
    //non-admin user looking at their evaluation list
   var html = HtmlService.createTemplateFromFile('ShowMyEval').evaluate().setTitle('My Evaluations').setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    para.deleteProperty(email+"Access");
 
  }
  
  return html;
}

//--------------------------------------------------------------------

function isSuper() {
 
  var para = PropertiesService.getUserProperties();
  
  var evalSheet = SpreadsheetApp.openById(eval_info).getSheetByName(supertab);
  var evalData = evalSheet.getDataRange().getValues();
  
  
  var i = 1;
  
  while (i < evalData.length && evalData[i][1] != email ) {
   i++; 
  }
  
  if (i == evalData.length) {
     para.setProperty(email + 'Access', 'No');
     return false;
    }// end if i == 8
    else {
      //is a principal or director
      para.setProperty(email + 'Access', evalData[i][2]);
      para.setProperty(email + "Name", evalData[i][0]);
     return true;
    }//end else
  
  return false;
  
}

//-----------------------------------------------------------------

function getStaff()  {
  
  var staffSpread = SpreadsheetApp.openById(eval_info).getSheetByName(email);
  var staffData = staffSpread.getDataRange().getValues();
  
  return staffData;
  
  
}

//-----------------------------------------------------------------

function setNew(usertmp, date, time) {
 
    var para = PropertiesService.getUserProperties();
    var cache = CacheService.getUserCache();
    Logger.log(usertmp);
    Logger.log(time);
    var user = usertmp.split("&");
   // var ss = SpreadsheetApp.openById(staffEval);
    
    cache.put(email + "New", "New");
    para.setProperty(email + "New-Values", user[0] + "&&" + date + "&&" + time + "&&" + user[1] + "&&" + user[2]);
  
    //if (!ss.getSheetByName(user[0]) ) {
    //   ss.insertSheet(user[0]);
   // }
  
   //Put in find for a file name user[0]
   // If not there copy template
  
   var staff = DriveApp.getFilesByName(user[0]);
   
   if(!staff.hasNext()) {
     var title = DriveApp.getFileById('11a4Jh2Xrz_BjXU3mXsHBPKgDzNKQR7WdvXp0GK8nMNU');
     //var doc = title.next();
     var newdoc = title.makeCopy(user[0]);
    }
  
   
   return true;
}

//----------------------------------------------------------------

function getEmail() {
 return Session.getActiveUser().getEmail(); 
}

//-----------------------------------------------------------------

function getEvalSheets(erole) {
  var ss = SpreadsheetApp.openById(eval_info).getSheetByName("Roles");
  var roles = ss.getDataRange().getValues();
  var i = 1;
  var cache = CacheService.getUserCache();
  var para = PropertiesService.getUserProperties();
  
  cache.remove(email + "New");
  while(i < roles.length && erole != roles[i][0]) {
    i++;
  }
  
  if (i == roles.length) {
   return "error"; 
  }
  else {
   var tab = roles[i][1];
   var rss = SpreadsheetApp.openById(tab);
   var tabs = rss.getSheets();
   para.setProperty(email + "SS", tab);
    
   for (var i=0; i < tabs.length; i++)
     tabs[i] = tabs[i].getName();
    
   return tabs;
  } 
  
  return "error";
}

//-----------------------------------------------------

function getProp(tab) {
  var para = PropertiesService.getUserProperties();
  var ss = SpreadsheetApp.openById(para.getProperty(email + "SS")).getSheetByName(tab);
  var prop = ss.getDataRange().getValues(); 
  
  return prop;
}

//-----------------------------------------------------

function createID() {
  var ss = SpreadsheetApp.openById(eval_info).getSheetByName("Misc");
  var misc = ss.getDataRange().getValues();
  
  var id = misc[0][0];
  var newcounter = id+1;
  var range2 = ss.getRange(1, 1, 1, 1);
  var countvalues = [[newcounter]];
  range2.setValues(countvalues); 
  
  return id;
  
}

//-------------------------------------------------

function postData(array) {
 
  var cache = CacheService.getUserCache();
  var para = PropertiesService.getUserProperties();
  //var ss = SpreadsheetApp.openById(staffEval).getSheetByName(array[2]);
  //var evals = ss.getDataRange().getValues();
  var i = 0;
  var row = '';
  
  //Get spreadsheet named after array[2]
  var staff = DriveApp.getFilesByName(array[2]);
   
   /*if(!staff.hasNext()) {
     var title = DriveApp.getFileById('11a4Jh2Xrz_BjXU3mXsHBPKgDzNKQR7WdvXp0GK8nMNU');
     //var doc = title.next();
     var staff = title.makeCopy(array[2]);
    }*/
  
  Logger.log('Before Opening SS');
  
   var ss = SpreadsheetApp.open(staff.next()).getSheetByName('Eval');
   //var ss = SpreadsheetApp.open(staff.next()).getSheetName('Eval');
   var evals = ss.getDataRange().getValues();
  
   Logger.log('After Opening SS');
  
  //see if ID is saved yet
  //if (!cache.get(array[0])) {
    
  while (i < evals.length && evals[i][1] != array[0]) {
    i++;
  }//end while
   if (i == evals.length) {
     if (i != 0) {
       i++;
     }
    cache.put(array[0], i, 3600);
    row = i;
   }
   else {
     cache.put(array[0], i, 3600);
     row = i;
    }
    
  /*}// end if ! cache
  else {
    row = cache.get(array[0]);
  }*/
  
  //save info into spreadsheet
  var update = [];
  var tmp = [];
  var propAvg = 0;
  var arrayloc = 8;
  var count = 0;
  var sumAvg = 0;
  
  update[0] = array[4]; //staff name
  update[1] = array[0]; //observation ID
  update[2] = array[1]; //eval spreadsheet id
  update[3] = para.getProperty(email +"Name"); //super's email
  update[4] = array[3]; //status
  update[5] = array[5]; //Date
  update[6] = array[6]; //time
  
  //the actual observation storage
  for (i=7; i < array.length; i++) {
    if (array[i].indexOf('&avg') >= 0) {
      update[arrayloc] = propAvg + '&avg';
      count = 0;
      sumAvg = 0;
      propAvg = 0;
      arrayloc = i;
    }
    else {
    
      if (array[i].indexOf('$~$') >= 0){
        tmp = array[i].split('$~$');
        count++;
        tmp[2] = parseInt(tmp[2], 10);
        sumAvg = sumAvg + tmp[2];
        propAvg = sumAvg/count;
      }
      
     update[i] = array[i]; 
      
    }
 
  }// end for
  update[arrayloc] = propAvg + '&avg';
  
  
  var save = ss.getRange(row+1, 1, 1, update.length);
  save.setValues([update]);
  
}

//---------------------------------------------------

function getEvalID(erole) {
  
   var ss = SpreadsheetApp.openById(eval_info).getSheetByName("Roles");
  var roles = ss.getDataRange().getValues();
  var i = 1;
  var cache = CacheService.getUserCache();
  var para = PropertiesService.getUserProperties();
  
  cache.remove(email + "New");
  while(i < roles.length && erole != roles[i][0]) {
    i++;
  }
  
  if (i == roles.length) {
   return "error"; 
  }
  else {
   return roles[i][1]; 
  }
  
}

//-------------------------------------------------------

function getEvals(staff) {
  //var ss = SpreadsheetApp.openById(staffEval);
  
  //if (!ss.getSheetByName(staff) ) {
    //   ss.insertSheet(staff);
  //  }
  
  //find spreadsheet by name 
  //if not there copy template
  
  var staffev = DriveApp.getFilesByName(staff);
   
   if(!staffev.hasNext()) {
     var title = DriveApp.getFileById('11a4Jh2Xrz_BjXU3mXsHBPKgDzNKQR7WdvXp0GK8nMNU');
     //var doc = title.next();
     staffev = title.makeCopy(staff);
    }
  
  var staffSpread = SpreadsheetApp.open(staffev.next()).getSheetByName('Eval');
  //var staffSpread = SpreadsheetApp.openById(staffEval).getSheetByName(staff);
  var staffData = staffSpread.getDataRange().getValues();
 
  
  return staffData;
}

//-------------------------------------------------------

function getStartDate() {
  var ss = SpreadsheetApp.openById(eval_info).getSheetByName("Misc");
  var misc = ss.getDataRange().getValues();
  
  var date = misc[1][0];
  
  return date;
  
}

//--------------------------------------------------------

function setID(id) {
  var para = PropertiesService.getUserProperties();
  
  para.setProperty(email + 'ID', id);
  
  return true;
}

//--------------------------------------------------------

function setUser(id) {
  var para = PropertiesService.getUserProperties();
  
  para.setProperty(email + 'User', id);
  
  return true;
}

//--------------------------------------------------------

function setPrint(id) {
  var para = PropertiesService.getUserProperties();
  
  para.setProperty(email + 'Print', id);
  
  return true;
}

//-----------------------------------------------------------

function openEvalSpread(tab) {
  
  //var para = PropertiesService.getUserProperties();
  var rss = SpreadsheetApp.openById(tab);
  var tabs = rss.getSheets();
  //para.setProperty(email + "SS", tab);
    
   for (var i=0; i < tabs.length; i++)
     tabs[i] = tabs[i].getName();
    
   return tabs;
  
}

//----------------------------------------------------------

function getObservation() {
 
  var para = PropertiesService.getUserProperties();
  var tmp = para.getProperty(email + 'ID').split('&');
  para.deleteProperty(email +'ID');
  //var ss = SpreadsheetApp.openById(staffEval).getSheetByName(tmp[0]);
  //var evals = ss.getDataRange().getValues();
  var i = 0;
  
  //get spreadsheet by name tmp[0]
  var staffev = DriveApp.getFilesByName(tmp[0]);
   
   if(!staffev.hasNext()) {
     var title = DriveApp.getFileById('11a4Jh2Xrz_BjXU3mXsHBPKgDzNKQR7WdvXp0GK8nMNU');
     //var doc = title.next();
     staffev = title.makeCopy(staff);
    }
  
  var ss = SpreadsheetApp.open(staffev.next()).getSheetByName('Eval');
  var evals = ss.getDataRange().getValues();
  
  while (i < evals.length && evals[i][1] != tmp[1] )  {
    i++;
  }//end while
  
  if (i == evals.length) {
    return false;
  } //end of
  else {
   var array = [];
    
    for (var j = 0; j < evals[i].length; j++)  {
        array[j] = evals[i][j];
    }
    
   return array; 
  }
  
}

//------------------------------------------------------------


function getGPrint(info) {
  

    
   //var info = para.getProperty(email + 'Print');
   var tmp = info.split('&');  //before & is eval email, after is eval id
  
  
    var para = PropertiesService.getUserProperties();
    para.setProperty(email + 'ID', info);
  
    var html = HtmlService.createTemplateFromFile('Print').evaluate().setTitle('Print Eval').setSandboxMode(HtmlService.SandboxMode.IFRAME); 
    //para.deleteProperty(email+"Access");
           
    var blob = html.getBlob();
    var pdf = blob.getAs("application/pdf");
    

    var file = DriveApp.createFile(pdf).setName(tmp[0] + " " + tmp[1] + " Eval.pdf");
   //  driveService.files().get(file.getId())
    //.executeMediaAndDownloadTo(outputStream);
   para.deleteProperty(email + 'ID');
   //file.addViewer(email);
   var fileId = file.getId();

   try { 
    Drive.Permissions.create(
      {
        role: 'reader',
        type: 'user',
        emailAddress: email
      },
      fileId,
      {
        sendNotificationEmail: false // This suppresses the default Google Drive email
      }
    );
  } catch (e) {
    console.error("Error sharing file: " + e.message);
    return;
  }


 var URL = file.getDownloadUrl().replace("&e=download&gd=true","");
   URL = 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=sharing';
 //  Logger.log(URL);

  //var file = DriveApp.getFileById(fileId);
  var fileUrl = file.getUrl();
  var fileName = file.getName();
  
  var emailSubject = "New document shared with you: " + fileName;
  var emailBody = "Hello,\n\nA File has been shared with you.\n\n" +
                  "You can view it here: " + fileUrl + "\n\n" +
                  "Please do not reply to this email.";

  MailApp.sendEmail({
    to: email,
    subject: emailSubject,
    body: emailBody,
    noReply: true // Forces the email to come from a noreply address
  });

  
    
    
    
  
    file.getParents().next().removeFile(file);
   DriveApp.getFolderById('1X2RlRTpc6aFh6n26aMZ2Z0ezSbzaK_3o').addFile(file);

  
  
  return URL;

  
}


//---------------------------------------------------------------

function getPrintID(){
  
 var para =  PropertiesService.getUserProperties();
 var info = para.getProperty(email + 'Print');
 para.deleteProperty(email + 'Print');
 return info;
  
}

//------------------------------------------------

function emailStaff(array) {
  
  var message = "";
  var subject = "";
  
  var para =  PropertiesService.getUserProperties();
  
  para.setProperty(email + 'ID', array[2]+'&'+array[0]);
  var obs = getObservation();
  para.deleteProperty(email + 'ID');
  
  //check to see if array[3] is Released or Finalized
  if (array[3] == "Released") {
    message = array[4] + " your observation, completed by " + obs[3] + " has been released for your review. To view your observations please go to: http://eval.carrabec.org";
    message = message + "\n Please email " + obs[3] + " if you have any questions.";
 
    subject = "Your observation on " + array[5] + " has been released for review";
  }
  else {
     message = array[4] + " your observation, completed by " + obs[3] + " and is now finalized. To view your observations please go to: http://eval.carrabec.org";
    message = message + "\n Please email " + obs[3] + " if you have any questions.";
    
    subject = "Your observation on " + array[5] + " has been finalized";
  }
  
  MailApp.sendEmail(array[2], subject, message, {noReply: true, htmlBody: message});

  
  
  
}

//------------------------------------------------
//sends data needed to evaluation calculations
function sendFinalized(array) {
//Post the data  
  postData(array);

 // Find the name of the rubric 
 var spreadtemp = SpreadsheetApp.openById(array[1]);
 var rubricName = spreadtemp.getName();

 console.log(rubricName);

   if (rubricName == "Teacher Eval 1-3") {
  // is a teacher with Prop 1-3 so do all the rest of the code
  
  var ssid = '1vFRG5kVA1Gu7r7CINXTBhMA6TECEGFb_mTM-boernpE';
  var staffemail = array[2];

  Logger.log("In Teacher");

  var prop13 = 9;
  var prop45 = 4;
  
  var para =  PropertiesService.getUserProperties();
  para.setProperty(email + 'ID', array[2]+'&'+array[0]);
  var obs = getObservation();
  para.deleteProperty(email + 'ID');

var section = 1;
var count = 1;
var data = [[]];
var tmp = [];
var count2 = 0;
var ss = SpreadsheetApp.openById(ssid).getSheetByName('Observation');
var calceval = ss.getDataRange().getValues();
var i = 1;

while(i < calceval.length && calceval[i][0] != staffemail){
i++;
}//end while

if (calceval[i][0] == staffemail) {
var ib;
var text;
//prep data for calculations
data[0][0] = array[5];

Logger.log(i);

for(var ia = 0; ia < obs.length; ia++) {

 //console.info(obs[ia]);
 
 text = obs[ia].toString();
 if(text.indexOf('&avg') >= 0){

 ib = ia-1;
 
 //console.info(obs[ib]);
 
 //take data from previous cell
 if (obs[ib] != "NS") {
 //text = obs[ib].toString();
 tmp = obs[ib].split("$");
 data[0][count] = Number(tmp[4]);
 
 console.info(data[0][count]);
 
 count++;
 count2++;
 
 //console.info(tmp[4] + " " + ib);
 
 }//end if

else {
data[0][count] = "";
count++;

}//end else

 }//end if

   }//end for

} //if is in db
//console.info(count2);

//send data to calculations spreadsheet
//Test to see if is Prop 1-3 or 4-5
var loc = 0;

if(count2 > prop45){
//is Prop 1-3

//check to see if data exists
if(calceval[i][4] == ""){
 
 loc = 4;
} //end if
else if(calceval[i][14] == ""){
 loc = 14;

}//end else if
else if(calceval[i][24] == "") {
 loc = 24;
}//
else if(calceval[i][34] == "") {
 loc = 34;
}//
else if(calceval[i][44] == "") {
 loc = 44;
}//
else if(calceval[i][54] == "") {
 loc = 54;
}//
else if(calceval[i][64] == "") {
 loc = 64;
}//
else if(calceval[i][74] == "") {
 loc = 74;
}//


if (loc != 0) {
// write to DB
   var values =[[data[0][0], data[0][1], data[0][2], data[0][3], data[0][4], data[0][5], data[0][6], data[0][7], data[0][8], data[0][9]]];
   var save = ss.getRange(i+1, loc, 1, 10);
   save.setValues(values);
   
}//end if loc not 0
} //if prop 1-3

/*else {
//is Prop 45

if(calceval[i][44] == "") {
 loc = 44;
}//
else if (calceval[i][49] == "") {
 loc = 49;
}//)


if (loc != 0) {
 // Write to DB
  var values =[[data[0][0], data[0][10], data[0][11], data[0][12], data[0][13]]];
  var save = ss.getRange(i+1, loc, 1, 5);
  save.setValues(values);
 } 
 


}//end else*/

} //end if teacher 1-3
else if (rubricName == "Teacher Eval 4-5") {

  // is a teacher with Prop 4-5 so do all the rest of the code
  
  var ssid = '1vFRG5kVA1Gu7r7CINXTBhMA6TECEGFb_mTM-boernpE';
  var staffemail = array[2];

  Logger.log("In Teacher");

  var prop13 = 9;
  var prop45 = 4;
  
  var para =  PropertiesService.getUserProperties();
  para.setProperty(email + 'ID', array[2]+'&'+array[0]);
  var obs = getObservation();
  para.deleteProperty(email + 'ID');

var section = 1;
var count = 1;
var data = [[]];
var tmp = [];
var count2 = 0;
var ss = SpreadsheetApp.openById(ssid).getSheetByName('Prop 4-5');
var calceval = ss.getDataRange().getValues();
var i = 1;

while(i < calceval.length && calceval[i][0] != staffemail){
i++;
}//end while

if (calceval[i][0] == staffemail) {
var ib;
var text;
//prep data for calculations
data[0][0] = array[5];

Logger.log(i);

for(var ia = 0; ia < obs.length; ia++) {

 //console.info(obs[ia]);
 
 text = obs[ia].toString();
 if(text.indexOf('&avg') >= 0){

 ib = ia-1;
 
 //console.info(obs[ib]);
 
 //take data from previous cell
 if (obs[ib] != "NS") {
 //text = obs[ib].toString();
 tmp = obs[ib].split("$");
 data[0][count] = Number(tmp[4]);
 
 console.info(data[0][count]);
 
 count++;
 count2++;
 
 //console.info(tmp[4] + " " + ib);
 
 }//end if

else {
data[0][count] = "";
count++;

}//end else

 }//end if

   }//end for

} //if is in db
//console.info(count2);

//send data to calculations spreadsheet
//Test to see if is Prop 1-3 or 4-5
var loc = 0;

//check to see if data exists
if(calceval[i][4] == ""){
 
 loc = 4;
} //end if
else if(calceval[i][9] == ""){
 loc = 9;

}//end else if
else if(calceval[i][14] == "") {
 loc = 14;
}//
else if(calceval[i][19] == "") {
 loc = 19;
}//
else if(calceval[i][24] == "") {
 loc = 24;
}//

if (loc != 0) {
 // Write to DB
  var values =[[data[0][0], data[0][1], data[0][2], data[0][3], data[0][4]]];
  var save = ss.getRange(i+1, loc, 1, 5);
  save.setValues(values);
 } 

} //end else if 4-5

} 

