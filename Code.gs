var eval_info = '1hkNcwLNWBHiKY8OkRUpMBBjubw-WYKea70quwvu_Nqk';
var tab = 'Complete Data';
var supertab = 'Supervisors';
var email = Session.getActiveUser().getEmail();
var teacherdata = '1vFRG5kVA1Gu7r7CINXTBhMA6TECEGFb_mTM-boernpE';
var teachertab = 'Complete Data';
var ID = '';

function doGet(e) {
 
 //get list of administrators
 var adminsheet = SpreadsheetApp.openById(eval_info).getSheetByName('Supervisors');
 var adminData = adminsheet.getDataRange().getValues();
 var i = 1;

 while (i <= adminData.length-1 && email != adminData[i][1]) {
  i++;
 } //end while

 if(i != adminData.length && email == adminData[i][1]) {
  if (e.parameter.ID)
  {
    //Has an email
    ID = e.parameter.ID;
  // add check to see if admin has access to this teacher
  var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName(tab);
  var teacherData = teachersheet.getDataRange().getValues();
  var j = 1;
  var test = false;

 while (j <teacherData.length && !test) {
    if (teacherData[j][2] == adminData[i][0] && teacherData[j][1] == ID) {
      test = true;
     } 
  
  j++;

  }

  if (test) {
  
  var html = HtmlService.createTemplateFromFile('TeacherEvalStatus').evaluate().setTitle('Teacher Eval Status').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  return html;

  }

else {
    //load page with all their staff
    var html = HtmlService.createTemplateFromFile('NotAdmin').evaluate().setTitle('Not an admin').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
  }
  }
if(e.parameter.print) 
{
  MailMerge();
  //load page with all their staff
    var html = HtmlService.createTemplateFromFile('PrintEval').evaluate().setTitle('Status of Teacher Evaluations').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
}
  else {
    //load page with all their staff
    var html = HtmlService.createTemplateFromFile('TeachersStatus').evaluate().setTitle('Status of Teacher Evaluations').setSandboxMode(HtmlService.SandboxMode.IFRAME);
    return html;
  }
 
 }//end if email
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
      Logger.log(i);
  }//end while ! ID
  
  if (ID == status[i][1])
   return status[i];
  else 
    return 0;
 
}

//-----------------------------------------------------------------
//Returns one teacher's status of their evaluation to the index.html file when called.

function getPD () {


var tab = 'Complete Data'

  var ss = SpreadsheetApp.openById(teacherdata).getSheetByName('Supervisor Entries');
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][0] != ID) {
      i++;
      Logger.log(i);
  }//end while ! ID
  
  if (ID == status[i][0])
   return status[i];
  else 
    return 0;
 
}

//----------------------------------------------------------------

//Returns one teacher's status of their finalized observation to the index.html file when called.

function getObs () {


var tab = 'Complete Data'

  var ss = SpreadsheetApp.openById(teacherdata).getSheetByName('Observation');
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][0] != ID) {
      i++;
      Logger.log(i);
  }//end while ! ID
  
  if (ID == status[i][0])
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
      
      Logger.log(evals[count]);

      count++;
      
    }
    Logger.log(i);
    i++;

 }//end while

 return evals; 

}

//-----------------------------------------------------------------


function getAllPD() {
  var supervisor = getName();

  var teachersheet = SpreadsheetApp.openById(teacherdata).getSheetByName('Supervisor Entries');
  var teacherData = teachersheet.getDataRange().getValues();
  var i = 1;
  var count = 0;
  var evals = new Array(new Array());
  var j;

 while (i <teacherData.length) {
    if (teacherData[i][7] == supervisor) {
       evals[count] = teacherData[i];
      
      Logger.log(evals[count]);

      count++;
      
    }
    Logger.log(i);
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
//Create the finalized evaluation forms for appropriate staff

function MailMerge()
{

//get the evals and name of the user
var info = getEvals();
var supervisor = getName();

Logger.log(supervisor);

//Get date
var date = new Date();
var month =date.getMonth() + 1;
var day = date.getDate();
var year = date.getFullYear();

//Document copy and paste
  var title = "Teacher Evaluations for " + supervisor + " " + month + "/" + day + "/" + year;
  var doc = DriveApp.getFileById("1ebgW-JX-FEaYFZQaC5RomdV8XsYZg2z3R1WrSFgghA4"); //getting the template document
  var newdoc = doc.makeCopy(title); //create the document
  var num = newdoc.getId(); //get the ID of the new document
  var final = DocumentApp.openById(num); //open the new document for editting
  var body = final.getBody(); //get the body of the document
  var otherBody = final.getBody();
  var pars = final.getParagraphs();
  var blank = "";

Logger.log(title);

  //body.insertPageBreak
  for( var i in pars ) //loop to keep a copy of the original paragraphs
   pars[i] = pars[i].copy();

  for(i=0; i < info.length; i++)
  {
    if (info[i][5] != '1 of 2') {   
      // Ensure the staff member should get an eval
      body.replaceText('<<TeacherName>>', info[i][0]);
      body.replaceText('<<SchoolYear>>', info[i][3]);
      body.replaceText('<<Principal>>', info[i][2]);
      body.replaceText('<<Cycle>>', info[i][5]);
      body.replaceText('<<ProfObsFactor>>', info[i][7]);
      
      if (info[i][15] != "")
       body.replaceText('<<ObsScore>>', info[i][15].toFixed(2));

      if (info[i][22] != "")
       body.replaceText('<<WeightedObs>>', info[i][22].toFixed(2));
      
      body.replaceText('<<PDFactor>>', info[i][8]);
      body.replaceText('<<PD>>', info[i][19]);
     
      if (info[i][6] != "1. No PD")
       body.replaceText('<<WeightedPD>>', info[i][23].toFixed(2));
      if (info[i][6] == "1. No PD")
        body.replaceText('<<WeightedPD>>', blank);
      
      body.replaceText('<<PeerObsFactor>>', info[i][9]);
      
      if (info[i][16] != "")
       body.replaceText('<<PeerObs>>', info[i][16].toFixed(2));
      
      if (info[i][24] != "")
       body.replaceText('<<WeightedPeer>>', info[i][24].toFixed(2));
      
      if (info[i][25] != "")
       body.replaceText('<<SUM>>', info[i][25].toFixed(2));
      
      body.replaceText('<<Rating>>', info[i][27]);
      body.replaceText('<<TotalFactors>>', info[i][10]);
    
    if (i < info.length) 
    {
      Logger.log(info[i][0]);
     //final.appendPageBreak();
     body.appendPageBreak();
     var otherBody = DocumentApp.openById("1ebgW-JX-FEaYFZQaC5RomdV8XsYZg2z3R1WrSFgghA4").getBody();
     var totalElements = otherBody.getNumChildren();

    //for (var j = 0; j < pars.length; ++j) {
      //final.appendParagraph(pars[j].copy());

     for( var j = 0; j < totalElements; j++ ) {
      var element = otherBody.getChild(j).copy();
      var type = element.getType();
      //var type = pars[j].getType();
      
      //console.log(type);
     
      if( type == DocumentApp.ElementType.PARAGRAPH )
        body.appendParagraph(element);

      else if( type == DocumentApp.ElementType.TABLE )
      //if( type == DocumentApp.ElementType.TABLE )
        body.appendTable(element);
    }
    
    }
   } //end if not 1 of 2
  }
  Logger.log('Done');
  //body = finalbody;
  final.saveAndClose();

var user = Session.getActiveUser().getEmail();
MailApp.sendEmail (user, title, newdoc.getUrl(), {noReply: true});

}
