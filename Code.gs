//This shows teachers their eval status for the current year. 

var staffEval = '1vFRG5kVA1Gu7r7CINXTBhMA6TECEGFb_mTM-boernpE';
var tab = 'Complete Data';
var email = Session.getActiveUser().getEmail();


function doGet() {
 

  var ss = SpreadsheetApp.openById(staffEval).getSheetByName(tab);
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  //Check to see if the user is in the teacher list for evaluations.
  while ( i < status.length-1 && status[i][1] != email) {
      i++;
      //Logger.log(i);
  }//end while ! email
  
  //show status if in the teacher list, otherwise shows a screen that states that they are not in the list. 
  if (email == status[i][1])
   var html = HtmlService.createTemplateFromFile('index').evaluate().setTitle('My Eval Status').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  else 
    var html = HtmlService.createTemplateFromFile('notteacher').evaluate().setTitle('Not in the teacher list').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
 
  return html;
}


//-----------------------------------------------------------------
//Returns the teacher's status of their PD Sheet to the index.html file when called.

function getPD () {

  var ss = SpreadsheetApp.openById(staffEval).getSheetByName('Supervisor Entries');
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][0] != email) {
      i++;
      Logger.log(i);
  }//end while ! email
  
  if (email == status[i][0])
   return status[i];
  else 
    return 0;
 
}

//-----------------------------------------------------------------
//Returns the teacher's status of their evaluation to the index.html file when called.

function getStatus () {

  var ss = SpreadsheetApp.openById(staffEval).getSheetByName(tab);
  var status = ss.getDataRange().getValues();
  var i = 1; 
  
  while ( i < status.length-1 && status[i][1] != email) {
      i++;
      Logger.log(i);
  }//end while ! email
  
  if (email == status[i][1])
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

function postType(array) {
  var ss = SpreadsheetApp.openById(staffEval).getSheetByName(tab);
  var status = ss.getDataRange().getValues();
  var i =1;
  console.log(array[0]);

  while (i < status.length && status[i][1] != email) {
    i++;
  }

  if (status[i][1] == email) {
    var cell = "4. Needs to Select";

    if (array[0] == "1. No PD"){
      cell = "1. No PD";
      console.log('No PD');
    }
    else if (array[0] == "2. With PD"){
      cell = "2. With PD";
      console.log('PD');
    }

     ss.getRange(i+1, 7).setValue(cell);

  } //end of selected user 


} //end function postType
