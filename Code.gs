var eval_info = '1hkNcwLNWBHiKY8OkRUpMBBjubw-WYKea70quwvu_Nqk';
var tab = 'Complete Data';
var supertab = 'Supervisors';
var email = Session.getActiveUser().getEmail();
var teacherdata = '1vFRG5kVA1Gu7r7CINXTBhMA6TECEGFb_mTM-boernpE';
var teachertab = 'Complete Data';
var evalDataTab = 'Supervisor Entries'
var type = '';

function doGet(e) {
 
 //get list of administrators
 var adminsheet = SpreadsheetApp.openById(eval_info).getSheetByName('Supervisors');
 var adminData = adminsheet.getDataRange().getValues();
 var i = 1;

 while (i <= adminData.length-1 && email != adminData[i][1]) {
  i++;
 } //end while

 if(i != adminData.length && email == adminData[i][1]) {
  if (e.parameter.Type)
  {
    //Will open the correct page based on type. 
    type = e.parameter.Type;
  // add check to see if admin has access to this teacher
  var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
  var teacherData = teachersheet.getDataRange().getValues();
  var j = 1;
  var test = false;

 while (j <teacherData.length && !test) {
    if (teacherData[j][2] == adminData[i][0]) {
      test = true;
     } 
  
  j++;

  }

  if (test) {
  
  if (type == "Peer") {
  var html = HtmlService.createTemplateFromFile('PeerEdit').evaluate().setTitle('Teacher Eval Peer Observation').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
  }// end if peer
  else if (type == "StuPer") {
  var html = HtmlService.createTemplateFromFile('StuPerEdit').evaluate().setTitle('Teacher Eval Status Student Perception Survey').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
  }// end if student perception survey
  else if (type == "PD") {
  var html = HtmlService.createTemplateFromFile('PD').evaluate().setTitle('Teacher Eval Status PD').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;
  } // end if pd
  }// end if test

else {
    //load page with all their staff
    var html = HtmlService.createTemplateFromFile('NotAdmin').evaluate().setTitle('Not an admin').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
  }
  }
  else {
    //load page with all their staff
    var html = HtmlService.createTemplateFromFile('NotAdmin').evaluate().setTitle('Not an admin').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
  }
 
 }//end if type
 else {
   var html = HtmlService.createTemplateFromFile('NotAdmin').evaluate().setTitle('Not an admin').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
 }
}

//-----------------------------------------------------------------
//Returns one teacher's status of their evaluation to the index.html file when called.

function getStatus () {


var tab = 'Complete Data'

  var ss = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][1] != ID) {
      i++;
      //Logger.log(i);
  }//end while ! ID
  
  if (ID == status[i][1])
   return status[i];
  else 
    return 0;
 
}

//-----------------------------------------------------------------
//Returns one teacher's status of their PD to the index.html file when called.

function getPD (teacher) {


  var ss = SpreadsheetApp.openById(teacherdata).getSheetByName(evalDataTab);
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][0] != teacher) {
      i++;
      //Logger.log(i);
  }//end while ! ID
  
  if (teacher == status[i][0])
   return status[i];
  else 
    return 0;
 
}

//----------------------------------------------------------------
// Gets the user's currently logged in Google email

function getEmail() {
 return Session.getActiveUser().getEmail(); 
}

//-----------------------------------------------------------------
//returns the evals that have this supervisor listed

function getEvals() {
  var supervisor = getName();

  var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
  var teacherData = teachersheet.getDataRange().getValues();
  var i = 1;
  var count = 0;
  var evals = new Array(new Array());
  var j;

 while (i <teacherData.length) {
    if (teacherData[i][2] == supervisor) {
       evals[count] = teacherData[i];
      
      //Logger.log(evals[count]);

      count++;
      
    }
   // Logger.log(i);
    i++;

 }//end while

 return evals; 

}

//-----------------------------------------------------------------
// returns the name used in the teacher eval spreadsheet

function getName() {
  //get list of administrators
 var adminsheet = SpreadsheetApp.openById(eval_info).getSheetByName('Supervisors');
 var adminData = adminsheet.getDataRange().getValues();
 var i = 1;

 while (i <= adminData.length-1 && email != adminData[i][1]) {
  i++;
 } //end while

 if(i != adminData.length && email == adminData[i][1]) {
   return adminData[i][0];
 }
 else return 0;

}

//-----------------------------------------------------------------

function postPeer(x) {
 
 var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
 var teacherData = teachersheet.getDataRange().getValues();
 var j = 0;
 var teacher = "";
 var tloc = 0;
 var values = ""
 //var array = [];

Logger.log(x.length);
Logger.log(x);
 
 if (x.length > 0) {

 while (j < x.length && x[j] != "") {
  var array = x[j].split("&");
  teacher = array[0];
  tloc = 0;

  Logger.log(teacher);
  Logger.log(array[1]);

  while (tloc <= teacherData.length && teacherData[tloc][1] != teacher) {
           tloc++
  }


 //Logger.log(tloc);
 
  if (tloc <= teacherData.length && teacherData[tloc][1] == teacher) {
    
    values = array[1];
    if (values == "Null") {
      values = "";
    }
    var save = teachersheet.getRange(tloc+1, 17, 1, 1);
    save.setValue(values);
  }
  j++;
 }
 }

}

//-----------------------------------------------------------------

function postStuPer(x) {
 
 var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(evalDataTab);
 var teacherData = teachersheet.getDataRange().getValues();
 var j = 0;
 var teacher = "";
 var tloc = 0;
 var values = ""
 //var array = [];

Logger.log(x.length);
Logger.log(x);
 
 if (x.length > 0) {

 while (j < x.length && x[j] != "") {
  var array = x[j].split("&");
  teacher = array[0];
  tloc = 0;

  Logger.log(teacher);
  Logger.log(array[1]);

  while (tloc <= teacherData.length && teacherData[tloc][0] != teacher) {
           tloc++
  }


 //Logger.log(tloc);
 
  if (tloc <= teacherData.length && teacherData[tloc][0] == teacher) {
    
    values = array[1];
    if (values == "Null") {
      values = "";
    }
    var save = teachersheet.getRange(tloc+1, 4, 1, 1);
    save.setValue(values);
  }
  j++;
 }
 }

}

//-----------------------------------------------------------------

function postPD(x) {
 
 var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
 var teacherData = teachersheet.getDataRange().getValues();
 var j = 0;
 var teacher = "";
 var tloc = 0;
 var values = ""
 //var array = [];

Logger.log(x.length);
Logger.log(x);
 
 if (x.length > 0) {

 while (j < x.length && x[j] != "") {
  var array = x[j].split("&");
  teacher = array[0];
  tloc = 0;

  Logger.log(teacher);
  Logger.log(array[1]);

  while (tloc <= teacherData.length && teacherData[tloc][1] != teacher) {
           tloc++
  }


 //Logger.log(tloc);
 
  if (tloc <= teacherData.length && teacherData[tloc][1] == teacher) {
    console.log('Here I am');
    values = array[2];
    if (values == "Null") {
      values = "";
    }
   // if (array[1] == 1) { 
   // var save = teachersheet.getRange(tloc+1, 18, 1, 1);
   // }
   // else if (array[1] == 2) {
      var save = teachersheet.getRange(tloc+1, 20, 1, 1);
   // }
    save.setValue(values);
  }
  j++;
 }
 }

}

