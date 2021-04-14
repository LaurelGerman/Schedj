function onOpen() { //onOpen runs automatically when the spreadsheet is opened
  //var spreadsheet = SpreadsheetApp.openById(1nY4w8a_rFhG1S5GmZKID7XmYdFa824X5vv1j9VbE3Bg);
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  //--------Make Menu--------//
  
  var myMenu = []; 
  
  myMenu.push({name: "Read Me", functionName: "readMe"});
  myMenu.push({name: "Set Up: ST", functionName: "setupST"});
  myMenu.push({name: "Set Up: MT", functionName: "setupMT"});
  //myMenu.push({name: "Set Up: MT Eq", functionName: "setupMTEQ"});  
  myMenu.push(null);
  myMenu.push({name: "Run Schedule: ST", functionName: "runScheduleST"});
  myMenu.push({name: "Run Schedule: MT", functionName: "runScheduleMT"});
  myMenu.push(null);
  myMenu.push({name: "Create Deputy Import", functionName: "deputyConvert"});

  spreadsheet.addMenu("STC Scheduler", myMenu);
}


//################################ Menu Buttons ################################//

function readMe(){
  var ui = SpreadsheetApp.getUi();
  ui.alert('Welcome to the STC Shift Scheduler! \n\nBefore you start: \n1. Paste the whenisgood availability into a blank tab and name that tab Whenisgood_Input \n2. Paste the Google form results into another tab and name that tab Google_Form_Input \n3. Create a third tab named Timeslots_Input and indicate when shifts can be scheduled \n4. Create a fourth tab named Locations and indiate attributes about each location \n5. Create a fifth tab named Meals and indicate Yale Dining Hall hours \nThen: \n6. Go to STC Scheduler > Run Setup \n7. If you get no errors, go to STC Scheduler > Run Schedule!');
}

function setupST(){
  var Roster = setupGeneric(4);
  //formatSTPrefs(Roster);
}

function setupMT(){ //not done yet but in theory this could use the same formatAvailability subroutine as the STs
  var Roster = setupGeneric(4);
  formatMTPrefs(Roster);
}

function setupGeneric(slotsize){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var oldAvailability = spreadsheet.getSheetByName("Availability");
  var oldRoster = spreadsheet.getSheetByName("Roster");
  var oldTimeslots = spreadsheet.getSheetByName("Timeslots");
  var oldSchedule = spreadsheet.getSheetByName("Schedule");
  var whenIsGood_usr = spreadsheet.getSheetByName("Whenisgood_Input");
  //var gForm_usr = spreadsheet.getSheetByName("Google_Form_Input");
  var gForm_usr = spreadsheet.getSheetByName("ROSTERTEMPLATE");
  var timeslots_usr = spreadsheet.getSheetByName("Timeslots_Input");
    
  //--------Create Availability, Roster, Timeslots Sheets--------//
  //these are copies of Whenisgood_Input, Google_Form_Input, and Timeslots_Input
  //that we'll modify so we don't touch the original data
  
  if(oldAvailability != null){
    spreadsheet.deleteSheet(oldAvailability);
  }
  var Availability = whenIsGood_usr.copyTo(spreadsheet);
  Availability.setName("Availability"); 
  
  if(oldRoster != null){
    spreadsheet.deleteSheet(oldRoster);
  }
  var Roster = gForm_usr.copyTo(spreadsheet);
  Roster.setName("Roster");
  
  if(oldTimeslots != null){
    spreadsheet.deleteSheet(oldTimeslots);
  }
  var Timeslots = timeslots_usr.copyTo(spreadsheet);
  Timeslots.setName("Timeslots");
  
  if(oldSchedule != null){
    spreadsheet.deleteSheet(oldSchedule);
  }
  
  //--------Run "Setup" Steps--------//
  //Format and add to Availability, Timeslot, Roster tabs
  formatAvailability(Availability);
  formatTimeslots(Timeslots);
  var Schedule = Timeslots.copyTo(spreadsheet);
  Schedule.setName("Schedule");
  //formatPrefs(Roster, Availability, slotsize);
  return Roster;
}

//############################################################################################ Setup ############################################################################################//
//This runs separately from "Run Schedule" to prevent Google timeout error and to leave a buffer to fix any user input errors
//It parses user input and creates the data tables used in the actual "Run schedule" command
//It duplicates the user data tabs and the scheduler works off the duplicates, so the original user input data never gets modified
//The Format Roster Tab section also determines location priority and target scheduling hours and adds this information to the Roster for each user to cut down on time spent sorting in the scheduling algorithm


//################################ Format Availability Tab ################################//
//This function is called from multiple setup buttons

function formatAvailability(Availability){
  var avail = Availability.getDataRange().getValues();
  
  //--------Delete and Add Columns--------//
  avail = deleteCol(Availability, avail, "Can Make It");
  avail = addIndexRow(Availability, avail, "Name Index", 2, 2, 2);
  avail = addIndexCol(Availability, avail, "Date Index", 2, 3, 2);
  
  //--------Aesthetic Formatting--------//
  //Change names to all uppercase to make string searches easier
  //Freeze rows and colummns
  for(var personCell=3; personCell<=avail[0].length; personCell++){
    var cell = Availability.getDataRange().getCell(1, personCell);
    cell.setFormula('=UPPER("' + avail[0][personCell-1] + '")');
  }  
  Availability.setFrozenRows(2);
  Availability.setFrozenColumns(2);
}

//################################ Format Timeslots Tab ################################//
function formatTimeslots(Timeslots){
  var times = Timeslots.getDataRange().getValues();
  
  //--------Add Index Columns--------//
  times = addIndexRow(Timeslots, times, "Location Index", 2, 2, 2);
  times = addIndexCol(Timeslots, times, "Date Index", 2, 3, 2);
  var timesRLength = times[0].length;
  var timesCLength = times.length;
  
  //--------Add Dashes--------//
  for(var c=2; c<timesRLength; c++){
    for(var r=2; r<timesCLength; r++){
      if(typeof times[r][c] === 'string' && times[r][c]==""){
        times[r][c]="--";
      }
    }
  }
  
  Timeslots.getDataRange().setValues(times);
  
  //--------Aesthetic Formatting--------//
  Timeslots.setFrozenRows(2);
  Timeslots.setFrozenColumns(2); 
} 


//################################ Format Roster Tab ################################//
//This function is called from multiple setup buttons.
function formatPrefs(Roster, Availability, slotsize){
  var prefs = Roster.getDataRange().getValues();
  
  //--------Add Columns--------//
  //Add a column with all-caps firstname lastname to match names on the Availability tab
  prefs = addCol(Roster, prefs, "Name", 3);
  prefs = addCol(Roster, prefs, "Name Index", 4);
  prefs = addCol(Roster, prefs, "Hours Offered", 5);
  prefs = addCol(Roster, prefs, "Target", 6);
  prefs = addCol(Roster, prefs, "Hrs Given", 7);
  prefs = addCol(Roster, prefs, "% of Goal", 8);
  
  //--------Parse Addresses Into Names--------//
  var prefRLength = prefs[1].length;
  var prefCLength = prefs.length;
  
  var emailorname = 0;
  if(prefs[0][1]=="Name_email"){
    emailorname = 2;
  }
  if(prefs[0][1]=="Name_firstlast"){
    emailorname = 1;
  }
  var prefRange = Roster.getDataRange();
  
  if(emailorname == 1){ //firstlast
    for(var p=1; p<prefCLength; p++){
      var cell = prefRange.getCell(p+1,3); 
      cell.setFormulaR1C1('=UPPER(CONCATENATE(LEFT(LEFT(INDIRECT("R[0]C[-1]",FALSE),SEARCH("@",INDIRECT("R[0]C[-1]",FALSE))-1),SEARCH(".",LEFT(INDIRECT("R[0]C[-1]",FALSE),SEARCH("@",INDIRECT("R[0]C[-1]",FALSE))-1))-1)," ", RIGHT(LEFT(INDIRECT("R[0]C[-1]",FALSE),SEARCH("@",INDIRECT("R[0]C[-1]",FALSE))-1),LEN(LEFT(INDIRECT("R[0]C[-1]",FALSE),SEARCH("@",INDIRECT("R[0]C[-1]",FALSE))-1))-SEARCH(".",LEFT(INDIRECT("R[0]C[-1]",FALSE),SEARCH("@",INDIRECT("R[0]C[-1]",FALSE)))))))');
    }
  } 
  if(emailorname == 2){//email
    for(var p=1; p<prefCLength; p++){
      var cell = prefRange.getCell(p+1,3);
      cell.setFormulaR1C1('=UPPER(R[0]C[-1])');
    }
  }
  prefs = prefRange.getValues();
  
  //--------Add Name Index Numbers--------//
  var avail = Availability.getDataRange().getValues();
  var availRLength = avail[0].length;
  var availCLength = avail.length;
  var neednamewarning = 0;

  for(var p=1; p<prefCLength; p++){
    if(prefs[p][3] == "ERROR" || prefs[p][3] == ""){
      var thisname = 0;
      for(var a=2; a<availRLength; a++){
        if(prefs[p][2] == avail[0][a]){
          prefs[p][3] = avail[1][a];
          a+= availRLength;
          thisname = 1;
        }
      }
      if(thisname == 0){
        prefs[p][3] = "ERROR"
        neednamewarning = 1;
      }
    }
  }
  
  //--------Add Hours Offered--------//  
  //Calculate total hours each person listed as available
  var hoursOffered=[avail[1],avail[0]];  
  for(var person=2; person<availRLength; person++){
    var thissum = 0;
    for(var time=2; time<availCLength; time++){
      if(avail[time][person] != "NO"){
        thissum++;
      }
    }
    hoursOffered[1][person]=thissum/slotsize;
  }
  for(var name = 1; name<prefCLength; name++){
    prefs[name][4]=hoursOffered[1][prefs[name][3]];
  }

  //--------Format Priority Column if it exists--------//
  //Reduce answers to a single word: "Hours" or "Location"
  var priorityLabel = findLabel("Which of the following would you prefer?", prefs[0], "Priority");
  var priority = 0;
  
  if(priorityLabel > -1){
    priority = 1;
    for(var person=1; person<prefCLength; person++){
      if(prefs[person][priorityLabel] == "Working in my preferred locations"){
        prefs[person][priorityLabel] = "Location";}
      if(prefs[person][priorityLabel] == "Getting as many hours as possible"){
        prefs[person][priorityLabel] = "Hours";}
    }
  }
  
  prefRange.setValues(prefs);
  
  //--------Aesthetic Formatting--------//
  if(neednamewarning == 1){
    var ui = SpreadsheetApp.getUi();
    ui.alert('Some student names could not be parsed. Check the Roster tab and fix errors in the Name Index column before proceeding.');
  }
  Roster.setFrozenRows(1);
  Roster.setFrozenColumns(2); 
}


//################################ Format ST Roster Tab ################################//
//This function is called from SetupST button after formatAvailability has run.

function formatSTPrefs(Roster){
  var prefs = Roster.getDataRange().getValues();
  var certLabel = findLabel("Check off all of the following that you are, if applicable:", prefs[0], "Certs");
  var wantCerts = findLabel("Check off all of the following that you would like to become this semester, if applicable:", prefs[0], "Cert Goals");
  var goalLabel = findLabel("% of Goal", prefs[0], "% of Goal");
  Roster.getDataRange().setValues(prefs);
  
  
  //--------Add Cert & Location Columns--------//  
  prefs = addCol(Roster, prefs, "MacSpec", certLabel+3);
  prefs = addCol(Roster, prefs, "HWSpec Cert", certLabel+4);
  prefs = addCol(Roster, prefs, "Coord", certLabel+5);
  prefs = addCol(Roster, prefs, "Want MacSpec", certLabel+6);
  prefs = addCol(Roster, prefs, "Want HWSpec", certLabel+7);
  prefs = addCol(Roster, prefs, "Want Coord", certLabel+8);
  prefs = addCol(Roster, prefs, "IO Hrs", goalLabel+2);
  prefs = addCol(Roster, prefs, "IO Target", goalLabel+3);
  prefs = addCol(Roster, prefs, "TTO Hrs", goalLabel+4);
  prefs = addCol(Roster, prefs, "TTO Target", goalLabel+5);
  //prefs = addCol(Roster, prefs, "SAGE Hrs", goalLabel+6);
  //prefs = addCol(Roster, prefs, "SAGE Target", goalLabel+7);
  
  //this is weird and hard-coded
  var MacSpecLabel = findLabel("MacSpec", prefs[0], "MacSpec");
  var HWSpecLabel = findLabel("HWSpec Cert", prefs[0], "HWSpec Cert");
  var CoordLabel = findLabel("Coord", prefs[0], "Coord");
  var WantMacSpecLabel = findLabel("Want MacSpec", prefs[0], "Want MacSpec");
  var WantHWSpecLabel = findLabel("Want HWSpec", prefs[0], "Want HWSpec");
  var WantCoordLabel = findLabel("Want Coord", prefs[0], "Want Coord");
  
  var prefRange = Roster.getDataRange();
  for(var student = 2; student < prefs.length+1; student++){
    var cell01 = prefRange.getCell(student,MacSpecLabel+1);
    cell01.setFormulaR1C1('=if(isnumber(search("MacSpec",indirect("R[0]C[-2]",false))),1,0)');
    var cell02 = prefRange.getCell(student,HWSpecLabel+1);
    cell02.setFormulaR1C1('=if(isnumber(search("HWSpec",indirect("R[0]C[-3]",false))),1,0)');
    var cell03 = prefRange.getCell(student,CoordLabel+1);
    cell03.setFormulaR1C1('=if(isnumber(search("Coord",indirect("R[0]C[-4]",false))),1,0)');
    var cell04 = prefRange.getCell(student,WantMacSpecLabel+1);
    cell04.setFormulaR1C1('=if(isnumber(search("MacSpec",indirect("R[0]C[-4]",false))),1,0)');
    var cell05 = prefRange.getCell(student,WantHWSpecLabel+1);
    cell05.setFormulaR1C1('=if(isnumber(search("HWSpec",indirect("R[0]C[-5]",false))),1,0)');
    var cell06 = prefRange.getCell(student,WantCoordLabel+1);
    cell06.setFormulaR1C1('=if(isnumber(search("Coord",indirect("R[0]C[-6]",false))),1,0)');
  } 
  prefs = prefRange.getValues();
  
  //--------Rename Cert & Location Columns--------//
  var minTTOLabel = findLabel("What is the minimum number of hours that you would like to work in the TTO?", prefs[0], "Min TTO");
  var minioLabel = findLabel("What is the minimum number of hours that you would like to work in the io?", prefs[0], "Min IO");
  //var minSageLabel = findLabel("What is the minimum number of hours that you would like to work in Sage Hall?", prefs[0], "Min SAGE");
  var maxTTOLabel = findLabel("What is the maximum number of hours that you would like to work in the TTO?", prefs[0], "Max TTO");
  var maxioLabel = findLabel("What is the maximum number of hours that you would like to work in the io?", prefs[0], "Max IO");
  //var maxSageLabel = findLabel("What is the maximum number of hours that you would like to work in Sage Hall?", prefs[0], "Max SAGE");
  var shortLabel = findLabel("What is the shortest shift, in hours, that you would like to take?", prefs[0], "Shortest Shift");
  var longLabel = findLabel("What is the longest shift, in hours, that you would like to take?", prefs[0], "Longest Shift");
  var minLabel = findLabel("What is the minimum number of hours that you would like to work in total?", prefs[0], "Min Hrs");
  var maxLabel = findLabel("What is the maximum number of hours that you would like to work in total?", prefs[0], "Max Hrs");
  
  //--------Add Targets--------//
  //Targets Based on Priority, Hours Offered, Min and max hrs columns, calculate target hours for each student.
  //Students must list 3x as many hours as they want and cannot exceed 18 shift hrs per week.
  var priorityLabel = findLabel("Priority", prefs[0], "Priority");
  var targetLabel = findLabel("Target", prefs[0], "Target");
  var offeredLabel = findLabel("Hours Offered", prefs[0], "Hours Offered");
  var hrsGivenLabel = findLabel("Hrs Given", prefs[0], "Hrs Given");
  var goalLabel = findLabel("% of Goal", prefs[0], "% of Goal");
  var ioGivenLabel = findLabel("IO Hrs", prefs[0], "IO Hrs");
  var ttoGivenLabel = findLabel("TTO Hrs", prefs[0], "TTO Hrs");
  //var sageGivenLabel = findLabel("SAGE Hrs", prefs[0], "SAGE Hrs");
  var ttoTargetLabel = findLabel("TTO Target", prefs[0], "TTO Target");
  var ioTargetLabel = findLabel("IO Target", prefs[0], "IO Target");
  //var sageTargetLabel = findLabel("SAGE Target", prefs[0], "SAGE Target");
  
  var prefCLength = prefs.length;
  
  var indexLabel = findLabel("Name Index", prefs[0], "Name Index");

  Roster.getDataRange().setValues(prefs);
  
  /*for(var name = 1; name<prefCLength; name++){
    var person = prefs[name];
    var givenMinHrs = person[minLabel];
    var givenMaxHrs = person[maxLabel];
    var target = 0;
    var minioHrs = person[minioLabel];
    var minTTOHrs = person[minTTOLabel];
    var minSageHrs = person[minSageLabel];
    var maxioHrs = person[maxioLabel];
    var maxTTOHrs = person[maxTTOLabel];
    var maxSageHrs = person[maxSageLabel];
    var minSum = minTTOHrs + minioHrs + minSageHrs;
    var maxSum = maxTTOHrs + maxioHrs + maxSageHrs;
    
    //Validate
    if(minSum > givenMinHrs){
      Logger.log("Person %s: init minSum > givenmin. chg min %s > %s", person[indexLabel], givenMinHrs, minSum);
      givenMinHrs = minSum;
    }
    if(maxSum < givenMaxHrs){
      Logger.log("Person %s: init maxSum < givenmax. chg max %s > %s", person[indexLabel], givenMaxHrs, maxSum);
      givenMaxHrs = maxSum;
    }
    if(givenMinHrs > person[offeredLabel]/2){
      Logger.log("Person %s: givenmin > offered. chg min %s > %s", person[indexLabel], givenMinHrs, person[offeredLabel]/2);
      givenMinHrs = person[offeredLabel]/2;
    }
    if(givenMinHrs > givenMaxHrs){
      var avg = (givenMinHrs + givenMaxHrs) / 2;
      Logger.log("Person %s: givenmin > givenmax. chg min/max %s/%s > %s", person[indexLabel], givenMinHrs, givenMaxHrs, avg);
      givenMinHrs = avg;
      givenMaxHrs = avg;
    }
    if(givenMaxHrs > 18){
      Logger.log("Person %s: givenmax > 18. chg max %s > %s", person[indexLabel], givenMaxHrs, 18);
      givenMaxHrs = 18;
    }
    if(givenMinHrs > 18){
      Logger.log("Person %s: givenmin > 18. chg min %s > %s", person[indexLabel], givenMinHrs, 18);
      givenMinHrs = 18;
    }
    if(givenMinHrs < 2){
      Logger.log("Person %s: givenmin < 2. chg min %s > %s", person[indexLabel], givenMinHrs, 2);
      givenMinHrs = 2;
    }
    if(givenMaxHrs < 2){
      Logger.log("Person %s: givenmax < 2. chg max %s > %s", person[indexLabel], givenMaxHrs, 2);
      givenMaxHrs = 2;
    }

    if(minioHrs<2){
      Logger.log("Person %s: minio < 2. chg minio %s > %s", person[indexLabel], minioHrs, 2);
      minioHrs = 2;
    }
    if(maxioHrs<2){
      Logger.log("Person %s: maxio < 2. chg maxio %s > %s", person[indexLabel], maxioHrs, 2);
      maxioHrs = 2;
    }
    if(minioHrs>maxioHrs){
      var avg = (minioHrs + maxioHrs) / 2;
      Logger.log("Person %s: minio > maxio. chg minio/maxio %s/%s > %s", person[indexLabel], minioHrs, maxioHrs, avg);
      minioHrs = avg;
      maxioHrs = avg;
    }
    if(minTTOHrs>maxTTOHrs){
      var avg = minTTOHrs + maxTTOHrs / 2;
      Logger.log("Person %s: mintto > maxtto. chg mintto/maxtto %s/%s > %s", person[indexLabel], minTTOHrs, maxTTOHrs, avg);
      minTTOHrs = avg;
      maxTTOHrs = avg;
    }
    if(minSageHrs>maxSageHrs){
      var avg = minSageHrs + maxSageHrs / 2;
      Logger.log("Person %s: minSage > maxSage. chg minSage/maxSage %s/%s > %s", person[indexLabel], minSageHrs, maxSageHrs, avg);
      minSageHrs = avg;
      maxSageHrs = avg;
    }
    if(minSum != minTTOHrs + minioHrs + minSageHrs){
      Logger.log("Person %s: chg minsum %s > %s", person[indexLabel], minSum, minTTOHrs + minioHrs + minSageHrs);
    }
    if(maxSum != maxTTOHrs + maxioHrs + maxSageHrs){
      Logger.log("Person %s: chg maxsum %s > %s", person[indexLabel], maxSum, maxTTOHrs + maxioHrs + maxSageHrs);
    }
    minSum = minTTOHrs + minioHrs + minSageHrs;
    maxSum = maxTTOHrs + maxioHrs + maxSageHrs;
    
    if(minSum > givenMaxHrs){
      var nonio = minSum-minioHrs;
      var remove = minSum-givenMaxHrs;
      if(remove > nonio){
        jkl;
      }
      var ttoRatio = minTTOHrs/nonio;
      var sageRatio = minSageHrs/nonio;
      Logger.log("Person %s: minSum > givenmax. chg minTTO/minSage/minSum %s/%s/%s > %s/%s/%s", person[indexLabel], minTTOHrs, minSageHrs, minSum, minTTOHrs-(ttoRatio*remove), minSageHrs-(sageRatio*remove), givenMaxHrs);
      minTTOHrs = minTTOHrs-(ttoRatio*remove);
      minSageHrs = minSageHrs-(sageRatio*remove);
      minSum = givenMaxHrs;
    }
    if(minSum > person[offeredLabel]/2){
      var nonio = minSum-minioHrs;
      var remove = minSum-(person[offeredLabel]/2);
      if(remove > nonio){
        jkl;
      }
      var ttoRatio = minTTOHrs/nonio;
      var sageRatio = minSageHrs/nonio;
      Logger.log("Person %s: minSum > offered. chg minTTO/minSage/minSum %s/%s/%s > %s/%s/%s", person[indexLabel], minTTOHrs, minSageHrs, minSum, minTTOHrs-(ttoRatio*remove), minSageHrs-(sageRatio*remove), person[offeredLabel]/2);
      minTTOHrs = minTTOHrs-(ttoRatio*remove);
      minSageHrs = minSageHrs-(sageRatio*remove);
      minSum = person[offeredLabel]/2;
    }
    
    if(maxSum < givenMinHrs){
      var add = givenMinHrs - maxSum;
      Logger.log("Person %s: maxSum < givenmin. chg maxio/maxSum %s/%s > %s/%s", person[indexLabel], maxioHrs, maxSum, maxioHrs + add, givenMinHrs);
      maxioHrs = maxioHrs + add;
      maxSum = givenMinHrs;
    }
    if(minSum > givenMinHrs){
      Logger.log("Person %s: minSum > givenmin. chg givenmin %s > %s", person[indexLabel], givenMinHrs, minSum);
      givenMinHrs = minSum;
    }
    if(maxSum < givenMaxHrs){
      Logger.log("Person %s: maxSum < givenmax. chg givenmax %s > %s", person[indexLabel], givenMaxHrs, maxSum);
      givenMaxHrs = maxSum;
    }

    if(person[priorityLabel] == "Location"){
      target = (givenMaxHrs+givenMinHrs)/2;
    }
    if(person[priorityLabel] == "Hours"){
      target = givenMaxHrs;
    }
    if(target > person[offeredLabel]/2){
       target = person[offeredLabel]/2;
    }
    if(givenMinHrs>target){
      givenMinHrs = target;
    }
    
    if(givenMaxHrs - (minTTOHrs + minSageHrs) < maxioHrs){
      maxioHrs = givenMaxHrs - (minTTOHrs + minSageHrs);
      Logger.log("Person %s: adjust maxio", person[indexLabel]);
    }
    if(givenMaxHrs - (minioHrs + minSageHrs) < maxTTOHrs){
      maxTTOHrs = givenMaxHrs - (minioHrs + minSageHrs);
      Logger.log("Person %s: adjust maxtto", person[indexLabel]);
    }
    if(givenMaxHrs - (minioHrs + minTTOHrs) < maxSageHrs){
      maxSageHrs = givenMaxHrs - (minioHrs + minTTOHrs);
      Logger.log("Person %s: adjust maxsage", person[indexLabel]);
    }
    var ioAvg = (minioHrs + maxioHrs)/2;
    var ttoAvg = (minTTOHrs + maxTTOHrs)/2;
    var sageAvg = (minSageHrs + maxSageHrs)/2;
    var sumAvg = ttoAvg + ioAvg + sageAvg;
    var ttoRat = ttoAvg/sumAvg;
    var ioRat = ioAvg/sumAvg;
    var sageRat = sageAvg/sumAvg;
    var ioTarget = ioRat*target;
    if(ioTarget < minioHrs){
      ioTarget = minioHrs;
    }
    if(ioTarget > maxioHrs){
      ioTarget = maxioHrs;
    }
    var ttoTarget = ttoRat*target;
    if(ttoTarget < minTTOHrs){
      ttoTarget = minTTOHrs;
    }
    if(ttoTarget > maxTTOHrs){
      ttoTarget = maxTTOHrs;
    }
    var sageTarget = sageRat*target;
    if(sageTarget < minSageHrs){
      sageTarget = minSageHrs;
    }
    if(sageTarget > maxSageHrs){
      sageTarget = maxSageHrs;
    }
    
    person[targetLabel] = target;
    person[ttoTargetLabel] = ttoTarget;
    person[ioTargetLabel] = ioTarget;
    person[sageTargetLabel] = sageTarget;
    person[minioLabel] = minioHrs;
    person[maxioLabel] = maxioHrs;
    person[minTTOLabel] = minTTOHrs;
    person[maxTTOLabel] = maxTTOHrs;
    person[minSageLabel] = minSageHrs;
    person[maxSageLabel] = maxSageHrs;
    person[minLabel] = givenMinHrs;
    person[maxLabel] = givenMaxHrs;
    person[hrsGivenLabel] = 0;
    person[goalLabel] = 0;
    person[ioGivenLabel] = 0;
    person[ttoGivenLabel] = 0;
    person[sageGivenLabel] = 0;
    prefs[name] = person;    
  }
  */
  prefRange.setValues(prefs);
}



//################################ Format MT Roster Tab ################################//
//This function is called from SetupMT button after general setup has run.

function formatMTPrefs(Roster){
  var prefs = Roster.getDataRange().getValues();
  
  //--------Add Cert & Location Columns--------//  
  var mealLabel = findLabel("Are You On A Meal Plan?", prefs[0], "Meals");
  prefs = addCol(Roster, prefs, "212 YORK Min", mealLabel+2);
  prefs = addCol(Roster, prefs, "212 YORK Max", mealLabel+3);
  prefs = addCol(Roster, prefs, "CCAM Min", mealLabel+4);
  prefs = addCol(Roster, prefs, "CCAM Max", mealLabel+5);
  prefs = addCol(Roster, prefs, "EqSpec Min", mealLabel+6);
  prefs = addCol(Roster, prefs, "EqSpec Max", mealLabel+7);
  prefs = addCol(Roster, prefs, "ELO Min", mealLabel+8);
  prefs = addCol(Roster, prefs, "ELO Max", mealLabel+9);
  var expLabel = findLabel("Which Locations have you worked in last year?", prefs[0], "Experience");
  prefs = addCol(Roster, prefs, "212 YORK Cert", expLabel+2);
  prefs = addCol(Roster, prefs, "ELO Cert", expLabel+3);
  prefs = addCol(Roster, prefs, "CCAM Cert", expLabel+4);
  //prefs = addCol(Roster, prefs, "EQSPEC", expLabel+5);
  //prefs = addCol(Roster, prefs, "CMC", expLabel+6);
  //prefs = addCol(Roster, prefs, "PROJ", expLabel+7);
  var givenLabel = findLabel("Hrs Given", prefs[0], "Hrs Given");
  prefs = addCol(Roster, prefs, "212 YORK Hrs", givenLabel+2);
  prefs = addCol(Roster, prefs, "212 YORK Target", givenLabel+3);
  prefs = addCol(Roster, prefs, "CCAM Hrs", givenLabel+4);
  prefs = addCol(Roster, prefs, "CCAM Target", givenLabel+5);
  prefs = addCol(Roster, prefs, "EqSpec Hrs", givenLabel+6);
  prefs = addCol(Roster, prefs, "EqSpec Target", givenLabel+7);
  prefs = addCol(Roster, prefs, "ELO Hrs", givenLabel+8);
  prefs = addCol(Roster, prefs, "ELO Target", givenLabel+9);
  prefs = addCol(Roster, prefs, "Shortest Shift", givenLabel+10);
  prefs = addCol(Roster, prefs, "Longest Shift", givenLabel+11);
  
  //this is weird and hard-coded
  var YorkLabel = findLabel("212 YORK Cert", prefs[0], "212 YORK Cert");
  var EloLabel = findLabel("ELO Cert", prefs[0], "ELO Cert");
  var CcamLabel = findLabel("CCAM Cert", prefs[0], "CCAM Cert");
  //var WantMacSpecLabel = findLabel("Want MacSpec", prefs[0], "Want MacSpec");
  //var WantHWSpecLabel = findLabel("Want HWSpec", prefs[0], "Want HWSpec");
  //var WantCoordLabel = findLabel("Want Coord", prefs[0], "Want Coord");
  
  var prefRange = Roster.getDataRange();
  for(var student = 2; student < prefs.length+1; student++){
    //var cell01 = prefRange.getCell(student,YorkLabel+1);
    //cell01.setFormulaR1C1('=if(isnumber(search("212 York",indirect("R[0]C[-1]",false))),1,0)');
    var cell02 = prefRange.getCell(student,EloLabel+1);
    cell02.setFormulaR1C1('=if(isnumber(search("ELO",indirect("R[0]C[-2]",false))),1,0)');
    var cell03 = prefRange.getCell(student,CcamLabel+1);
    cell03.setFormulaR1C1('=if(isnumber(search("CCAM",indirect("R[0]C[-3]",false))),1,0)');
    /*var cell04 = prefRange.getCell(student,WantMacSpecLabel+1);
    cell04.setFormulaR1C1('=if(isnumber(search("MacSpec",indirect("R[0]C[-4]",false))),1,0)');
    var cell05 = prefRange.getCell(student,WantHWSpecLabel+1);
    cell05.setFormulaR1C1('=if(isnumber(search("HWSpec",indirect("R[0]C[-5]",false))),1,0)');
    var cell06 = prefRange.getCell(student,WantCoordLabel+1);
    cell06.setFormulaR1C1('=if(isnumber(search("Coord",indirect("R[0]C[-6]",false))),1,0)');*/
  } 
  prefs = prefRange.getValues();
  
  var yorkCert = findLabel("Skill Level (1=inexperienced, 2=experienced)", prefs[0], "Skill Level");
  var trackLabel = findLabel("Which track are you?", prefs[0], "Track");
  
  for(var student = 1; student<prefs.length; student++){
    if(prefs[student][trackLabel]  == "212 York"){
      if(prefs[student][yorkCert] == 2){
        prefs[student][YorkLabel] = 1;
      }
      else{
        prefs[student][YorkLabel]=0;
      }
    }
    else{
      prefs[student][YorkLabel]=0;
    }
  }
  
  //--------Rename Cert & Location Columns--------//
  var minHrsLabel = findLabel("Min Number Of Hours You Would Like To Work?", prefs[0], "Min Hrs");
  var maxHrsLabel = findLabel("Max Number Of Hours You Would Like To Work?", prefs[0], "Max Hrs");
  var shortShiftLabel = findLabel("Shortest Shift", prefs[0], "Shortest Shift");
  var longShiftLabel = findLabel("Longest Shift", prefs[0], "Longest Shift");
  var mealLabel = findLabel("Are you okay with working during the entirety of normal dining hall hours?", prefs[0], "Thru Meals");
  var eqspecRankLabel = findLabel("Rank These Locations By Your Desire To Work There [EqSpec]", prefs[0], "EqSpec Rank");
  var ccamRankLabel = findLabel("Rank These Locations By Your Desire To Work There [CCAM]", prefs[0], "CCAM Rank");
  var eloRankLabel = findLabel("Rank These Locations By Your Desire To Work There [ELO]", prefs[0], "ELO Rank");
  
  var eloGivenLabel = findLabel("ELO Hrs", prefs[0], "ELO Hrs");
  var ccamGivenLabel = findLabel("CCAM Hrs", prefs[0], "CCAM Hrs");
  var yorkGivenLabel = findLabel("212 YORK Hrs", prefs[0], "212 YORK Hrs");
  var hrsGivenLabel = findLabel("Hrs Given", prefs[0], "Hrs Given");
  
  for(var student = 1; student<prefs.length; student++){
    prefs[student][shortShiftLabel] = 1;
    prefs[student][longShiftLabel] = 4;
    prefs[student][hrsGivenLabel] = 0;
    prefs[student][eloGivenLabel] = 0;
    prefs[student][ccamGivenLabel] = 0;
    prefs[student][yorkGivenLabel] = 0;
    if(prefs[student][mealLabel] =="Yes"||prefs[student][mealLabel] =="No, but don't schedule me through an entire meal"){
      prefs[student][mealLabel] = "No";
    }
    if(prefs[student][mealLabel] =="No, and I'm willing to work through meal hours"){
      prefs[student][mealLabel] = "Yes";
    }
  }
    
  
  
  /*var minTTOLabel = findLabel("What is the minimum number of hours that you would like to work in the TTO?", prefs[0], "Min TTO");
  var minioLabel = findLabel("What is the minimum number of hours that you would like to work in the io?", prefs[0], "Min IO");
  var minSageLabel = findLabel("What is the minimum number of hours that you would like to work in Sage Hall?", prefs[0], "Min SAGE");
  var maxTTOLabel = findLabel("What is the maximum number of hours that you would like to work in the TTO?", prefs[0], "Max TTO");
  var maxioLabel = findLabel("What is the maximum number of hours that you would like to work in the io?", prefs[0], "Max IO");
  var maxSageLabel = findLabel("What is the maximum number of hours that you would like to work in Sage Hall?", prefs[0], "Max SAGE");
  var shortLabel = findLabel("What is the shortest shift, in hours, that you would like to take?", prefs[0], "Shortest Shift");
  var longLabel = findLabel("What is the longest shift, in hours, that you would like to take?", prefs[0], "Longest Shift");*/
  
  //--------Add Targets--------//
  //Targets Based on Priority, Hours Offered, Min and max hrs columns, calculate target hours for each student.
  //Students must list 3x as many hours as they want and cannot exceed 18 shift hrs per week.
  var minLabel = findLabel("Min Hrs", prefs[0], "Min Hrs");
  var maxLabel = findLabel("Max Hrs", prefs[0], "Max Hrs");
  var priorityLabel = findLabel("Priority", prefs[0], "Priority");
  var targetLabel = findLabel("Target", prefs[0], "Target");
  var offeredLabel = findLabel("Hours Offered", prefs[0], "Hours Offered");
  var hrsGivenLabel = findLabel("Hrs Given", prefs[0], "Hrs Given");
  var goalLabel = findLabel("% of Goal", prefs[0], "% of Goal");
  var eloGivenLabel = findLabel("ELO Hrs", prefs[0], "ELO Hrs");
  var ccamGivenLabel = findLabel("CCAM Hrs", prefs[0], "CCAM Hrs");
  var yorkGivenLabel = findLabel("212 YORK Hrs", prefs[0], "212 YORK Hrs");
  var eloTargetLabel = findLabel("ELO Target", prefs[0], "ELO Target");
  var ccamTargetLabel = findLabel("CCAM Target", prefs[0], "CCAM Target");
  var yorkTargetLabel = findLabel("212 YORK Target", prefs[0], "212 YORK Target");
  var trackLabel = findLabel("Track", prefs[0], "Track");
  var noCCAMLabel = findLabel("No CCAM", prefs[0], "No CCAM");
  var noELOLabel = findLabel("No ELO", prefs[0], "No ELO");
  var noEqspecLabel = findLabel("No EqSpec", prefs[0], "No EqSpec");
  var yorkMinLabel = findLabel("212 YORK Min", prefs[0], "212 YORK Min");
  var yorkMaxLabel = findLabel("212 YORK Max", prefs[0], "212 YORK Max");
  var ccamMinLabel = findLabel("CCAM Min", prefs[0], "CCAM Min");
  var ccamMaxLabel = findLabel("CCAM Max", prefs[0], "CCAM Max");
  var eloMinLabel = findLabel("ELO Min", prefs[0], "ELO Min");
  var eloMaxLabel = findLabel("ELO Max", prefs[0], "ELO Max");
  var eqSpecMinLabel = findLabel("EqSpec Min", prefs[0], "EqSpec Min");
  var eqSpecMaxLabel = findLabel("EqSpec Max", prefs[0], "EqSpec Max");
  
  var prefCLength = prefs.length;
  
  var indexLabel = findLabel("Name Index", prefs[0], "Name Index");

  for(var name = 1; name<prefCLength; name++){
    var person = prefs[name];
    var givenMinHrs = person[minLabel];
    var givenMaxHrs = person[maxLabel];
    var offered = person[offeredLabel];
    var eqRank = person[eqspecRankLabel];
    var eloRank = person[eloRankLabel];
    var ccamRank = person[ccamRankLabel];
    var track = person[trackLabel];
    var target = (givenMinHrs + givenMaxHrs)/2;
    if(target>offered){
      target = offered;
    }
    if(target>18){
      target = 18;
    }
    prefs[name][targetLabel]=target;
    /*var minioHrs = person[minioLabel];
    var minTTOHrs = person[minTTOLabel];
    var minSageHrs = person[minSageLabel];
    var maxioHrs = person[maxioLabel];
    var maxTTOHrs = person[maxTTOLabel];
    var maxSageHrs = person[maxSageLabel];
    var minSum = minTTOHrs + minioHrs + minSageHrs;
    var maxSum = maxTTOHrs + maxioHrs + maxSageHrs;*/
    
    //Convert ranks to numbers
   
    if(eqRank!=""){
      if(eqRank=="3rd Choice"){
        eqRank = 3;
        prefs[name][eqspecRankLabel]=3;
      }
      if(eqRank=="2nd Choice"){
        eqRank = 2;
        prefs[name][eqspecRankLabel]=2;
      }
      if(eqRank=="1st Choice"){
        eqRank = 1;
        prefs[name][eqspecRankLabel]=1;
      }
    }
    if(ccamRank!=""){
      if(ccamRank=="3rd Choice"){
        ccamRank = 3;
        prefs[name][ccamRankLabel]=3;
      }
      if(ccamRank=="2nd Choice"){
        ccamRank = 2;
        prefs[name][ccamRankLabel]=2;
      }
      if(ccamRank=="1st Choice"){
        ccamRank = 1;
        prefs[name][ccamRankLabel]=1;
      }
    }
    if(eloRank!=""){
      if(eloRank=="3rd Choice"){
        eloRank = 3;
        prefs[name][eloRankLabel]=3;
      }
      if(eloRank=="2nd Choice"){
        eloRank = 2;
        prefs[name][eloRankLabel]=2;
      }
      if(eloRank=="1st Choice"){
        eloRank = 1;
        prefs[name][eloRankLabel]=1;
      }
    }
    
    if(track=="Equipment Specialist"){
      prefs[name][trackLabel]="Eq";
      track = "Eq";
      var locCounter = 0;
      for(var l = 0; l<3; l++){ //HARDCODED. ALL OF THIS.
        if(person[noCCAMLabel+l]!=1){
          locCounter++; //count the places they're allowed/willing to work
        }
      }
      if(locCounter>0){ //if they're working in at least one location
        var locList = [];
        for(var l = 0; l<3; l++){
          if(person[noCCAMLabel+l]!=1){
            locList.push([ccamTargetLabel+(l*2),mealLabel+3+(l*2),person[ccamRankLabel+l],0,0]); //loc target lbl, loc min lbl, loc rank, weight, target
          }
        }
        locList.sort(function(a,b){ //sort first-last choice
          return a[2] - b[2];
        });
        
        var x=1;
        var s=0; //denominator/total weight parts (aka 4 + 2 + 1 = 7)
        
        for(var c=0; c<locCounter; c++){
          s=s+x;
          x=x*2;
          locList[c][3]=Math.pow(2,(locCounter-c-1)); //assign weighting to each location (1st choice has 2x 2nd choice, which has 2x 3rd choice)
        }
        
        for(var c=0; c<locCounter; c++){
          locList[c][4]=((target*locList[c][3])/s);
          prefs[name][locList[c][0]]=locList[c][4]; //assign targets to each location 
          if(c==0){
            prefs[name][locList[c][1]]=(givenMinHrs*locList[c][3])/s; //for first choice, min >0
          }
          if(c>0){
            prefs[name][locList[c][1]]=0; //for other choices, max = 0
          }
          prefs[name][locList[c][1]+1]=givenMaxHrs;
        }
      }
    }
    if(track=="212 York"){
      person[trackLabel]="AV";
      track = "AV";
      prefs[name][yorkTargetLabel] = target;
      prefs[name][yorkMinLabel] = givenMinHrs;
      prefs[name][yorkMaxLabel] = givenMaxHrs;
    }
    for(var l=0; l<4; l++){
      if(prefs[name][mealLabel+1+(l*2)]>0){
      }
      else{
        prefs[name][mealLabel+1+(l*2)]=0;
      }
      if(prefs[name][mealLabel+1+(l*2)+1]>0){
      }
      else{
        prefs[name][mealLabel+1+(l*2)+1]=0;
      }
      if(prefs[name][yorkTargetLabel+(l*2)]>0){
      }
      else{
        prefs[name][yorkTargetLabel+(l*2)]=0;
      }
    }
  }
   
  prefRange.setValues(prefs);
}



//############################################################################################ Make ST Schedule ############################################################################################//
function runScheduleST(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  
  var Availability = spreadsheet.getSheetByName("Availability");
  var availRange = Availability.getDataRange();
  var avail = availRange.getValues();
  var availRowLength = avail[0].length;
  var availColLength = avail.length;
  
  var Roster = spreadsheet.getSheetByName("Roster");
  var rosterRange = Roster.getDataRange();
  var rost = rosterRange.getValues();
  var rostRowLength = rost[0].length;
  var rostColLength = rost.length;
  
  var Timeslots = spreadsheet.getSheetByName("Timeslots");
  var timeRange = Timeslots.getDataRange();
  var times = timeRange.getValues();
  var timeRowLength = times[0].length;
  var timeColLength = times.length;
  
  var Schedule = spreadsheet.getSheetByName("Schedule");
  var schedRange = Schedule.getDataRange();
  var sched = schedRange.getValues();
  
  var locs = spreadsheet.getSheetByName("Locations").getDataRange().getValues();
  var locRowLength = locs[0].length;
  
  var meals = spreadsheet.getSheetByName("Meals").getDataRange().getValues();
  var mealColLength = meals.length;
  
  var availSchedNonstart = timeRange.getValues();
  var availSchedStart = timeRange.getValues();
  
  var nameIndexLabel = findLabel("Name Index", rost[0], "Name Index");
  var nameLabel = findLabel("Name", rost[0], "Name");
  var hrsOfferedLabel = findLabel("Hours Offered", rost[0], "Hours Offered");
  var targetLabel = findLabel("Target", rost[0], "Target");
  var hrsGivenLabel = findLabel("Hrs Given", rost[0], "Hrs Given");
  var minTimeLabel = findLabel("Shortest Shift", rost[0], "Shortest Shift");
  var maxTimeLabel = findLabel("Longest Shift", rost[0], "Longest Shift");
  var maxHrsLabel = findLabel("Max Hrs", rost[0], "Max Hrs");
  var mealLabel = findLabel("Thru Meals", rost[0], "Thru Meals");
  var hwCertLabel = findLabel("HWSPEC Cert", rost[0], "HWSPEC Cert");
  
  
  var slotSize = 4;
  
  //Set up mealTable
  var mealTable = [];
  for(var c=0; c<mealColLength; c++){
    mealTable.push([meals[c][1],meals[c][3],0]);
  }
 
  //Make rosterIndexTable
  //To use: rosterIndexTable.indexOf(name)
  var rosterIndexTable = []; //row of nameindex on roster
  for(var rosterName = 1; rosterName < rostColLength; rosterName++){
    rosterIndexTable[rosterName] = rost[rosterName][3];
  }
  
  //Make locLabelTable
  var locLabelTable = [];
  for(var l=0; l<locRowLength; l++){
    var thisLoc = locs[0][l];
    var given = findLabel(thisLoc.concat(" Hrs"),rost[0],thisLoc.concat(" Hrs"));
    var min = findLabel("Min ".concat(thisLoc),rost[0],"Min ".concat(thisLoc));
    var max = findLabel("Max ".concat(thisLoc), rost[0], "Max ".concat(thisLoc));
    var targ = findLabel(thisLoc.concat(" Target"), rost[0], thisLoc.concat(" Target"));
    var qual = findLabel(locs[6][l].concat(" Cert"), rost[0], locs[6][l].concat(" Cert"));
    var thisEntry = [l, given, targ, min, max, qual];
    locLabelTable.push(thisEntry);
  }
  
  //Set up schedTable
  for(var t=2; t<timeColLength; t++){
    for(var l=2; l<timeRowLength; l++){
      if(sched[t][l]=="--"){
        sched[t][l]=[sched[t][l],-1];
      }
      else{
        sched[t][l]=[sched[t][l],0];
      }
    }
  }
  
  
  //HARDCODED!!!! Make table of just tto
  var firstBatch = [];
  for(var t=0; t<timeColLength; t++){
    var row =[];
    for(var l=0; l<5; l++){ //SO HARDCODED
      if(l!=2){
        row.push(times[t][l]);
      }
    }
    firstBatch.push(row);
  }
  
  var freeList = makeFreeList(rost, rosterIndexTable, times, avail, locLabelTable);
  var startList = makeStartList(firstBatch, locs, rost, freeList, mealTable, slotSize, rosterIndexTable, locLabelTable, sched, avail, 1, 1);
  var startListLength = startList.length;
  
  for(var s=0; s<startListLength; s++){
    var t = startList[s][0];
    var l = startList[s][1];
    if(t>1 && sched[t][l][1]==0){
      if(startList[s][2]==0){
        sched[t][l]=["NO ONE CAN START", 0];
      }
      if(startList[s][2]>0){
        var studentList = [0];
        if(startList[s][2]==1){
          studentList = [[startList[s][3],0]];
        }
        else{
          studentList = pickStudent(startList[s], locLabelTable, rost, rosterIndexTable, hrsOfferedLabel, hrsGivenLabel);
        }
        var scheduledSomeone = 0;
        for(var ss=0; ss<studentList.length; ss++){
          var p = studentList[ss][0];
          if(sched[t][l][1]==0){
            var rostIndex = rosterIndexTable.indexOf(p);
            var minTime = rost[rostIndex][minTimeLabel];
            var maxTime = rost[rostIndex][maxTimeLabel];
            var givenHereLabel = locLabelTable[l][1];
            var maxHereLabel = locLabelTable[l][4];
            var givenHere = rost[rostIndex][givenHereLabel];
            var givenTtl = rost[rostIndex][hrsGivenLabel];
            var maxHrs = rost[rostIndex][maxHrsLabel];
            var maxHere = rost[rostIndex][maxHereLabel];
            //Logger.log("----Testing p %s t %s l %s----",p,t,l);
            
            var shiftSize = canStartHere(p, t, l, locs, rost, sched, freeList, mealTable, slotSize, rostIndex, minTime, maxTime, givenHere, givenTtl, maxHrs, maxHere, mealLabel, 1, avail, locLabelTable, 1);
            
            if(shiftSize>0){ //if we can schedule here
              scheduledSomeone = 1;
              ss=studentList.length;
              var shiftSizeSlots = shiftSize * slotSize;
              Logger.log("SUCCESS: p %s t %s l %s size %s", p,t,l,shiftSizeSlots);
              for(var w = 0; w<shiftSizeSlots; w++){ //update sched and mealTable
                sched[t+w][l][1]=p;
                sched[t+w][l][0]=rost[rostIndex][nameLabel];
                mealTable[t+w][2]=p;
                for(var ll = 2; ll<locRowLength; ll++){
                  if(sched[0][l]==sched[0][ll]){
                    freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p);
                  }
                }
              }
              for(var w=-slotSize; w<shiftSizeSlots+slotSize; w++){ //block out from list
                for(var ll = 2; ll<locRowLength; ll++){
                  if(t+w>=2 && t+w<timeColLength){
                    if(sched[0][l]!=sched[0][ll]){
                      freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p);
                    }
                  }
                }
              }
              rost[rostIndex][givenHereLabel]=givenHere+shiftSize;
              rost[rostIndex][hrsGivenLabel]=givenTtl+shiftSize;
            }
          }
        }
        if(scheduledSomeone == 0){
          sched[t][l]=["NO ONE WORKED", 0];
        }
      }
    }
  }
  
  
  //////// SCHEDULE HW SPECS ////////
  
  var hwSpecBatch = [];
  for(var t=0; t<timeColLength; t++){
    var row =[];
    for(var l=0; l<3; l++){
      row.push(times[t][l]);
    }
    hwSpecBatch.push(row);
  }
 
  var hwSpecRost = [rost[0]];
  for(var h = 2; h<=rostColLength; h++){
    var rostIndex = rosterIndexTable.indexOf(h);
    if(rost[rostIndex][hwCertLabel]==1){
      hwSpecRost.push(rost[rostIndex]);
    }
  }
  
  Logger.log("----Make hwSpecList----");
  var hwSpecList = makeStartList(hwSpecBatch, locs, rost, freeList, mealTable, slotSize, rosterIndexTable, locLabelTable, sched, avail, 3, 2);
  //var ioListLength = preIoList.length;
  var ioListLength = hwSpecList.length;
  
  Logger.log("----Schedule hwSpecs----");
  
  for(var s=0; s<ioListLength; s++){ // for each slot in ioList
    //var t = preIoList[s][0]; //time is preIoList[0]
    var t = hwSpecList[s][0];
    var l = hwSpecList[t][1]; //loc is hwSpecList[1] at t
    if(t>1 && sched[t][l][1]==0){ //if slot is free
    
      /*if(hwSpecList[t][2]==0){
        sched[t][l]=["NO ONE CAN START", 0];
      }*/
  
      if(hwSpecList[t][2]>0){ //if someone free
        var studentList = [0]; //start list of free people
        if(hwSpecList[t][2]==1){ //if one person free
          studentList = [[hwSpecList[t][3],0]]; //list is that person
        }
        else{
          studentList = pickStudent(hwSpecList[t], locLabelTable, rost, rosterIndexTable, hrsOfferedLabel, hrsGivenLabel); //make the student list
        }
        var scheduledSomeone = 0;
        for(var ss=0; ss<studentList.length; ss++){ //for each free person
          var p = studentList[ss][0];
          if(sched[t][l][1]==0){  //if slot is free
            var rostIndex = rosterIndexTable.indexOf(p);
            var minTime = rost[rostIndex][minTimeLabel];
            var maxTime = rost[rostIndex][maxTimeLabel];
            var givenHereLabel = locLabelTable[l][1];
            var maxHereLabel = locLabelTable[l][4];
            var givenHere = rost[rostIndex][givenHereLabel];
            var givenTtl = rost[rostIndex][hrsGivenLabel];
            var maxHrs = rost[rostIndex][maxHrsLabel];
            var maxHere = rost[rostIndex][maxHereLabel];
            //Logger.log("----Testing p %s t %s l %s----",p,t,l);
            
            var shiftSize = canStartHere(p, t, l, locs, rost, sched, freeList, mealTable, slotSize, rostIndex, minTime, maxTime, givenHere, givenTtl, maxHrs, maxHere, mealLabel, 1, avail, locLabelTable, 2);
            if(shiftSize>0){ //if we can schedule here
              scheduledSomeone = 1;
              ss=studentList.length; //end for loop
              var shiftSizeSlots = shiftSize * slotSize;
              Logger.log("SUCCESS: p %s t %s l %s size %s", p,t,l,shiftSizeSlots);
              for(var w = 0; w<shiftSizeSlots; w++){ //for each slot of the shift
                sched[t+w][l][1]=p; //make sched entry
                sched[t+w][l][0]=rost[rostIndex][nameLabel];
                mealTable[t+w][2]=p; //update mealTable
                for(var ll = 2; ll<locRowLength; ll++){ //for each location
                  if(sched[0][l]==sched[0][ll]){ //if it's HWSpec
                    freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p); //remove person from list
                  }
                }
              }
              for(var w=-slotSize; w<shiftSizeSlots+slotSize; w++){ //for each slot plus an hr on either side
                for(var ll = 2; ll<locRowLength; ll++){ //for each location
                  if(t+w>=2 && t+w<timeColLength){ //if not running off the ends
                    if(sched[0][l]!=sched[0][ll]){ //if not HWSpec
                      freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p); //remove person from list
                    }
                  }
                }
              }
              rost[rostIndex][givenHereLabel]=givenHere+shiftSize;
              rost[rostIndex][hrsGivenLabel]=givenTtl+shiftSize;
            }
          }
        }
      }
    }
    
  }
  
  /*
  for(var t=2; t<timeColLength; t++){
    for(var l=2; l<locRowLength; l++){
      if(l>=6){
        if(sched[t][2][1]>0){
          sched[t][l] = ["IO", 0];
        }
        else{
          sched[t][l] = ["--", -1];
        }
      }
      if(l==2){
        if(sched[t][2][1]>0){
          //do nothing
        }
        else{
          sched[t][l] = ["--", -1];
        }
      }
    }
  }
  */
  
  
  
  
  
  
  //HARDCODED
  var ioBatch = [];
  for(var t=0; t<timeColLength; t++){
    var row =[];
    for(var l=0; l<6; l++){
      if(l<2 || l>4){ //SO HARDCODED UGH
        row.push(times[t][l]);
      }
    }
    ioBatch.push(row);
  }
  
  
  //schedRange.setValues(sched);
  //rosterRange.setValues(rost);
  
  Logger.log("----Make preIoList----");
  var preIoList = makeStartList(ioBatch, locs, rost, freeList, mealTable, slotSize, rosterIndexTable, locLabelTable, sched, avail, 2, 0);
  
  
  /*var ioBatch2 = [times[0],times[1]];
  for(var t=2; t<timeColLength; t++){
    var row =[];
    for(var l=0; l<locRowLength; l++){
      if(l<2){
        row.push(times[t][l]);
      }
      if(l>=2 && l <=5){
        row.push("--");
      }
      if(l>5){
        row.push(sched[t][l][0]);
      }
    }
    ioBatch2.push(row);
  }*/
  
  Logger.log("----make ioList----");
  var ioList = makeStartList(ioBatch, locs, rost, freeList, mealTable, slotSize, rosterIndexTable, locLabelTable, sched, avail, 1, 1);
  var ioListLength2 = ioList.length;
  
  for(var s=0; s<ioListLength2; s++){ // for each slot in ioList
    var t = ioList[s][0]; //time is ioList[0]
    var l = ioList[s][1]; //loc is ioList[1] at t
    if(t>1 && sched[t][l][1]==0){ //if slot is free
    
      /*if(hwSpecList[t][2]==0){
        sched[t][l]=["NO ONE CAN START", 0];
      }*/
  
      if(ioList[s][2]>0){ //if someone free
        var studentList = [0]; //start list of free people
        if(ioList[s][2]==1){ //if one person free
          studentList = [[ioList[s][3],0]]; //list is that person
        }
        else{
          studentList = pickStudent(ioList[s], locLabelTable, rost, rosterIndexTable, hrsOfferedLabel, hrsGivenLabel); //make the student list
        }
        
        var scheduledSomeone = 0;
        for(var ss=0; ss<studentList.length; ss++){ //for each free person
          var p = studentList[ss][0];
          if(sched[t][l][1]==0){  //if slot is free
            var rostIndex = rosterIndexTable.indexOf(p);
            var minTime = rost[rostIndex][minTimeLabel];
            var maxTime = rost[rostIndex][maxTimeLabel];
            var givenHereLabel = locLabelTable[l][1];
            var maxHereLabel = locLabelTable[l][4];
            var givenHere = rost[rostIndex][givenHereLabel];
            var givenTtl = rost[rostIndex][hrsGivenLabel];
            var maxHrs = rost[rostIndex][maxHrsLabel];
            var maxHere = rost[rostIndex][maxHereLabel];
            //Logger.log("----Testing p %s t %s l %s----",p,t,l);
            
            var shiftSize = canStartHere(p, t, l, locs, rost, sched, freeList, mealTable, slotSize, rostIndex, minTime, maxTime, givenHere, givenTtl, maxHrs, maxHere, mealLabel, 1, avail, locLabelTable, 1);
            if(shiftSize>0){ //if we can schedule here
              scheduledSomeone = 1;
              ss=studentList.length; //end for loop
              var shiftSizeSlots = shiftSize * slotSize;
              Logger.log("SUCCESS: p %s t %s l %s size %s", p,t,l,shiftSizeSlots);
              for(var w = 0; w<shiftSizeSlots; w++){ //for each slot of the shift
                sched[t+w][l][1]=p; //make sched entry
                sched[t+w][l][0]=rost[rostIndex][nameLabel];
                mealTable[t+w][2]=p; //update mealTable
                for(var ll = 2; ll<locRowLength; ll++){ //for each location
                  if(sched[0][l]==sched[0][ll]){ //if it's HWSpec
                    freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p); //remove person from list
                  }
                }
              }
              for(var w=-slotSize; w<shiftSizeSlots+slotSize; w++){ //for each slot plus an hr on either side
                for(var ll = 2; ll<locRowLength; ll++){ //for each location
                  if(t+w>=2 && t+w<timeColLength){ //if not running off the ends
                    if(sched[0][l]!=sched[0][ll]){ //if not HWSpec
                      freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p); //remove person from list
                    }
                  }
                }
              }
              rost[rostIndex][givenHereLabel]=givenHere+shiftSize;
              rost[rostIndex][hrsGivenLabel]=givenTtl+shiftSize;
            }
          }
        }
      }
    }
  }
  
  schedRange.setValues(sched);
  rosterRange.setValues(rost);
}




//############################################################################################ Make MT Schedule ############################################################################################//
function runScheduleMT(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  
  var Availability = spreadsheet.getSheetByName("Availability");
  var availRange = Availability.getDataRange();
  var avail = availRange.getValues();
  var availRowLength = avail[0].length;
  var availColLength = avail.length;
  
  var Roster = spreadsheet.getSheetByName("Roster");
  var rosterRange = Roster.getDataRange();
  var rost = rosterRange.getValues();
  var rostRowLength = rost[0].length;
  var rostColLength = rost.length;
  
  var Timeslots = spreadsheet.getSheetByName("Timeslots");
  var timeRange = Timeslots.getDataRange();
  var times = timeRange.getValues();
  var timeRowLength = times[0].length;
  var timeColLength = times.length;
  
  var Schedule = spreadsheet.getSheetByName("Schedule");
  var schedRange = Schedule.getDataRange();
  var sched = schedRange.getValues();
  
  var locs = spreadsheet.getSheetByName("Locations").getDataRange().getValues();
  var locRowLength = locs[0].length;
  
  var meals = spreadsheet.getSheetByName("Meals").getDataRange().getValues();
  var mealColLength = meals.length;
  
  var availSchedNonstart = timeRange.getValues();
  var availSchedStart = timeRange.getValues();
  
  var nameIndexLabel = findLabel("Name Index", rost[0], "Name Index");
  var nameLabel = findLabel("Name", rost[0], "Name");
  var hrsOfferedLabel = findLabel("Hours Offered", rost[0], "Hours Offered");
  var targetLabel = findLabel("Target", rost[0], "Target");
  var hrsGivenLabel = findLabel("Hrs Given", rost[0], "Hrs Given");
  var minTimeLabel = findLabel("Shortest Shift", rost[0], "Shortest Shift");
  var maxTimeLabel = findLabel("Longest Shift", rost[0], "Longest Shift");
  var maxHrsLabel = findLabel("Max Hrs", rost[0], "Max Hrs");
  var mealLabel = findLabel("Thru Meals", rost[0], "Thru Meals");
  var ccamCertLabel = findLabel("CCAM Cert", rost[0], "CCAM Cert");
  var eqCertLabel = findLabel("EqSpec Cert", rost[0], "EqSpec Cert");
  var eloCertLabel = findLabel("ELO Cert", rost[0], "ELO Cert");
  
  var slotSize = 4;
  
  //Set up mealTable
  var mealTable = [];
  for(var c=0; c<mealColLength; c++){
    mealTable.push([meals[c][1],meals[c][3],0]);
  }
 
  //Make rosterIndexTable
  //To use: rosterIndexTable.indexOf(name)
  var rosterIndexTable = []; //row of nameindex on roster
  for(var rosterName = 1; rosterName < rostColLength; rosterName++){
    rosterIndexTable[rosterName] = rost[rosterName][3];
  }
  
  //Make locLabelTable: [location, loc hrs given, loc target, loc min, loc max, loc qualified]
  var locLabelTable = [];
  for(var l=0; l<locRowLength; l++){
    var thisLoc = locs[0][l];
    var given = findLabel(thisLoc.concat(" Hrs"),rost[0],thisLoc.concat(" Hrs"));
    var min = findLabel(thisLoc.concat(" Min"),rost[0],thisLoc.concat(" Min"));
    var max = findLabel(thisLoc.concat(" Max"), rost[0],thisLoc.concat(" Max"));
    var targ = findLabel(thisLoc.concat(" Target"), rost[0], thisLoc.concat(" Target"));
    var qual = findLabel(locs[6][l].concat(" Cert"), rost[0], locs[6][l].concat(" Cert"));
    var thisEntry = [l, given, targ, min, max, qual];
    locLabelTable.push(thisEntry);
  }
  
  //Set up schedTable
  for(var t=2; t<timeColLength; t++){
    for(var l=2; l<timeRowLength; l++){
      if(sched[t][l]=="--"){
        sched[t][l]=[sched[t][l],-1];
      }
      else{
        sched[t][l]=[sched[t][l],0];
      }
    }
  }
  
  //HARDCODED!!!! Make table of just elo and ccam
  /*var firstBatch = [];
  for(var t=0; t<timeColLength; t++){
    var row =[];
    for(var l=0; l<6; l++){
      row.push(times[t][l]);
    }
    firstBatch.push(row);
  }*/
  
  var freeList = makeFreeList(rost, rosterIndexTable, times, avail, locLabelTable);
  //the line below this said firstBatch instead of times
  var startList = makeStartList(times, locs, rost, freeList, mealTable, slotSize, rosterIndexTable, locLabelTable, sched, avail, 1, 1);
  var startListLength = startList.length;
  
  for(var s=0; s<startListLength; s++){
    var t = startList[s][0];
    var l = startList[s][1];
    if(t>1 && sched[t][l][1]==0){
      if(startList[s][2]==0){
        sched[t][l]=["NO ONE CAN START", 0];
      }
      if(startList[s][2]>0){
        var studentList = [0];
        if(startList[s][2]==1){
          studentList = [[startList[s][3],0]];
        }
        else{
          studentList = pickStudent(startList[s], locLabelTable, rost, rosterIndexTable, hrsOfferedLabel, hrsGivenLabel);
        }
        var scheduledSomeone = 0;
        for(var ss=0; ss<studentList.length; ss++){
          var p = studentList[ss][0];
          if(sched[t][l][1]==0){
            var rostIndex = rosterIndexTable.indexOf(p);
            var minTime = rost[rostIndex][minTimeLabel];
            var maxTime = rost[rostIndex][maxTimeLabel];
            var givenHereLabel = locLabelTable[l][1];
            var maxHereLabel = locLabelTable[l][4];
            var givenHere = rost[rostIndex][givenHereLabel];
            var givenTtl = rost[rostIndex][hrsGivenLabel];
            var maxHrs = rost[rostIndex][maxHrsLabel];
            var maxHere = rost[rostIndex][maxHereLabel];
            //Logger.log("----Testing p %s t %s l %s----",p,t,l);
            
            var shiftSize = canStartHere(p, t, l, locs, rost, sched, freeList, mealTable, slotSize, rostIndex, minTime, maxTime, givenHere, givenTtl, maxHrs, maxHere, mealLabel, 1, avail, locLabelTable, 1);
            if(shiftSize>0){ //if we can schedule here
              scheduledSomeone = 1;
              ss=studentList.length;
              var shiftSizeSlots = shiftSize * slotSize;
              Logger.log("SUCCESS: p %s t %s l %s size %s", p,t,l,shiftSizeSlots);
              for(var w = 0; w<shiftSizeSlots; w++){ //update sched and mealTable
                sched[t+w][l][1]=p;
                sched[t+w][l][0]=rost[rostIndex][nameLabel];
                mealTable[t+w][2]=p;
                for(var ll = 2; ll<locRowLength; ll++){
                  if(sched[0][l]==sched[0][ll]){
                    freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p);
                  }
                }
              }
              for(var w=-slotSize; w<shiftSizeSlots+slotSize; w++){ //block out from list
                for(var ll = 2; ll<locRowLength; ll++){
                  if(t+w>=2 && t+w<timeColLength){
                    if(sched[0][l]!=sched[0][ll]){
                      freeList[t+w][ll] = removeFromFreeList(freeList[t+w][ll],p);
                    }
                  }
                }
              }
              rost[rostIndex][givenHereLabel]=givenHere+shiftSize;
              rost[rostIndex][hrsGivenLabel]=givenTtl+shiftSize;
            }
          }
        }
        if(scheduledSomeone == 0){
          sched[t][l]=["NO ONE WORKED", 0];
        }
      }
    }
  }
  
  schedRange.setValues(sched);
  rosterRange.setValues(rost);
  jkl
}


//########################################################################################## Scheduling Functions ############################################################################################//
function pickStudent(startList, locLabelTable, rost, rostIndexTable, offeredLabel, givenLabel){ //of a given list of students who are free, which one should we try first?
  var numStudents = startList[2];
  var t = startList[0];
  var l = startList[1];
  var studentList = [];
  var givenHereLabel = locLabelTable[l][1];
  var targetHereLabel = locLabelTable[l][2];
  
  for(var p = 0; p<numStudents; p++){
    var rostIndex = rostIndexTable.indexOf(startList[3+p]);
    var givenHere = rost[rostIndex][givenHereLabel];
    var targetHere = rost[rostIndex][targetHereLabel];
    var offered = rost[rostIndex][offeredLabel];
    var given = rost[rostIndex][givenLabel];
    var stat = (targetHere-givenHere)/(offered-given);
    var entry = [startList[3+p],stat];
    studentList.push(entry);
  }
  
  studentList.sort(function(a,b){ //sort by difficulty
    return b[1] - a[1];
  });
  
  return studentList;
}
    

function makeStartList(times, locs, rost, freeList, mealTable, slotSize, rostIndexTable, locLabelTable, masterSched, avail, direction, gapChecker){
  var listLength = times.length;
  var locLength = times[0].length;
  //var locMult = locLength-2;
  var rostLength = rost.length;
  var minShiftLabel = findLabel("Shortest Shift", rost[0], "Shortest Shift");
  var givenLabel = findLabel("Hrs Given", rost[0], "Hrs Given");
  var maxShiftLabel = findLabel("Longest Shift", rost[0], "Longest Shift");
  var mealLabel = findLabel("Thru Meals", rost[0], "Thru Meals"); 
  var maxHrsLabel = findLabel("Max Hrs", rost[0], "Max Hrs");
  var startList=[];
  
  for(var l=2; l<locLength; l++){ //for each location
    var locNumber = times[1][l];
    var givenHereLabel = locLabelTable[locNumber][1];
    var maxHereLabel = locLabelTable[locNumber][4];
    for(var s=0; s<listLength; s++){ //for each timeslot
      startList.push([s,locNumber,-1]); //add initial value set
      if(s>=2){
        if(times[s][l]!="--"){
          var count = 0;
          for(var p=2; p<=rostLength; p++){ //for each person
            var rostIndex = rostIndexTable.indexOf(p);
            
            if(rostIndex != -1){ //if they exist
              var minShift = rost[rostIndex][minShiftLabel];
              var maxShift = rost[rostIndex][maxShiftLabel];
              var givenHere = rost[rostIndex][givenHereLabel];
              var givenTtl = rost[rostIndex][givenLabel];
              var maxHrs = rost[rostIndex][maxHrsLabel];
              var maxHere = rost[rostIndex][maxHereLabel];
              
              if(canStartHere(p, s, locNumber, locs, rost, masterSched, freeList, mealTable, slotSize, rostIndex, minShift, maxShift, givenHere, givenTtl, maxHrs, maxHere, mealLabel, 0, avail, locLabelTable, gapChecker)>0){ //if they can start here
                startList[((l-2)*listLength)+s].push(p); //add to the list if they can start here
                count++;
              }
              
            }
          }
          startList[((l-2)*listLength)+s][2]=count;
        }
      }
      if(startList[((l-2)*listLength)+s][2]>=0){
        Logger.log("t %s l %s startList row %s", s, l, startList[((l-2)*listLength)+s]);
      }
    }
  }
  
  if(direction==1){
    startList.sort(function(a,b){ //sort by difficulty
      return a[2] - b[2];
    });
    return startList;
  }
  if(direction==2){
    startList.sort(function(a,b){ //sort by difficulty
      return b[2] - a[2];
    });
    return startList;
  }
  if(direction==3){
    return startList;
  }
}


function makeFreeList(rost, rostIndexTable, times, avail, locLabelTable){
  var listLength = times.length;
  var locLength = times[0].length;
  var rostLength = rost.length;
  var freeList=[0,0];          
  for(var t=2; t<listLength; t++){
    freeList.push([0,t]);
    for(var l=2; l<locLength; l++){
      freeList[t].push([l,0]);
      var count = 0;
      for(var p=2; p<=rostLength; p++){
        var skip = 0;
        if(rostIndexTable.indexOf(p)==-1){
          skip = 1;
        }
        if(skip ==0){
          if(isFree(p,t,l, rost, avail, times, rostIndexTable.indexOf(p),locLabelTable)==1){
            freeList[t][l].push(p);
            count++;
          }
        }
      }
      freeList[t][l][1]=count;
    }  
  }
  return freeList;
}


function isFree(person, time, location, rost, avail, sched, rostIndex, locLabelTable){
  var locName = sched[0][location]; //was times[time][location]
  var ok = 0;
  if(locName != "--"){ //open
    if(avail[time][person] == "YES" || avail[time][person].indexOf(locName) != -1){ //free
      
      if(locLabelTable[location][5] == -1 || rost[rostIndex][locLabelTable[location][5]] == 1){ //qualified
        if(rost[rostIndex][locLabelTable[location][4]] > 0){ //wants to work here
          ok = 1;
        }
      }
    }
  }
  return ok;
}


function isFreeForShift(person, time, location, freeList, shiftLength, slotSize){
  var shiftLengthSlots = shiftLength*slotSize;
  var ok = 1;
  for(var t=0; t<shiftLengthSlots; t++){ //for each time at this loc
    var thisSlot = freeList[time+t][location];
    var proceed = 0;
    if(thisSlot[1]>0){ //if anyone free at this time
      for(var s=2; s<thisSlot[1]+2; s++){ //look through list for person
        if(thisSlot[s]==person){ //if you find the person
          proceed = 1; //move on
          s = thisSlot[1]+2; //stop looking
        }
      }
    }
    if(proceed==0){ //if don't move on
      ok = 0;
      t = shiftLengthSlots; //end for loop
    }
  }
  return ok;
}

function canStartHere(person, time, location, locs, rost, masterSched, freeList, mealTable, slotSize, rostIndex, minShift, maxShift, givenHere, givenTtl, maxHrs, maxHere, mealLabel, schedTag, avail, locLabelTable, gapChecker){
  if(locs[5][location] == "" || locs[5][location].search(time.toString()) == -1){
    if(isFree(person, time, location, rost, avail, masterSched, rostIndex, locLabelTable)==1){ //if they were originally free
      if(schedTag==0 || isFreeForShift(person, time, location, freeList, 1/slotSize, slotSize)==1){ //if they're still free
        var timeLeft = maxHere-givenHere;
        var minLoc = locs[2][location];
        var oneSlot = 1/slotSize;
        
        if(minLoc>minShift){ //set shortest shift
          minShift=minLoc;
        }
        var startMinShift = minShift;
        if(maxHrs-givenTtl<timeLeft){ //set time left
          timeLeft = maxHrs-givenTtl;
        } 
        var ok = 0;
        
        if(gapChecker==1){
          minShift = checkGaps(person, time, location, minShift, maxShift, locs, rost, masterSched, slotSize, rostIndex, schedTag, gapChecker, freeList, 0); //check gaps, extend shift length if necessary
        }
        if(gapChecker==2){
          minShift = checkGaps(person, time, location, minShift, maxShift, locs, rost, masterSched, slotSize, rostIndex, schedTag, gapChecker, freeList, 1); //check gaps, extend shift length if necessary
        }
        if(gapChecker==0){
          var endTime = time + (minShift*slotSize) -1;
          if(endTime>=masterSched.length){
            minShift = 0;
          }
        }
        //if(schedTag==1 && minTime==0){
        //Logger.log("p %s t %s l %s checkGaps Failed",person, time, location);
        //}      
        if(schedTag==1 && minShift > timeLeft+oneSlot){
          Logger.log("p %s t %s l %s minTime %s > timeLeft %s",person, time, location, minShift, timeLeft);
        }
        if(schedTag==1 && minShift > maxShift+oneSlot){
          Logger.log("p %s t %s l %s minTime %s > maxTime %s",person, time, location, minShift, maxShift);
        }
        
        if(minShift>0 && minShift<=timeLeft+oneSlot && minShift<=maxShift+oneSlot){//if we're still ok and resulting shift isnt too long
          if(rost[rostIndex][mealLabel]==1 || checkMeals(person, time, minShift, mealTable, slotSize)==1){ //if meals are ok
            ok = isFreeForShift(person, time, location, freeList, minShift, slotSize); //check that they're free for the whole time
            if(schedTag==1 && ok == 0){
              Logger.log("p %s t %s l %s isFreeForShift Failed",person, time, location);
            }
          }
          else{
            if(schedTag==1){
              Logger.log("p %s t %s l %s checkMeals Failed",person, time, location);
            }
          }
        }
        
        if(ok==1){
          return minShift;
        }
        else{
          if(gapChecker<2){
            return 0;
          }
          else{
            minShift = checkGaps(person, time, location, startMinShift, maxShift, locs, rost, masterSched, slotSize, rostIndex, schedTag, gapChecker, freeList, 0);
            if(minShift > 0 && minShift<=timeLeft+oneSlot && minShift<=maxShift+oneSlot){
              if(rost[rostIndex][mealLabel]==1 || checkMeals(person, time, minShift, mealTable, slotSize)==1){
                ok = isFreeForShift(person, time, location, freeList, minShift, slotSize);
              }
            }
            if(ok==1){
              return minShift;
            }
            else{
              return 0;
            }
          }
        }
        
      }
      else{
        if(schedTag==1){
          Logger.log("p %s t %s l %s first pass isFreeForShift Failed",person, time, location);
        }
      }
    }
    else{
      if(schedTag==1){
        //Logger.log("p %s t %s l %s basic isFree Failed",person, time, location);
      }
      return 0;
    }
  }
}


function checkGaps(person, time, location, shiftLength, maxShift, locs, rost, masterSched, slotSize, rostIndex, schedTag, gapMode, freeList, extend){
  var locMinSlots = locs[2][location]*slotSize;
  var shiftLengthSlots = shiftLength*slotSize;
  var ok = 1;
  var connectAfter = 0;
  var hitDuring = 0;
  var mustExtend = 0;
  
  if(masterSched[time-1][location][1] == 0 && gapMode == 1){ //if not directly touching prev slot
    for(var b=1; b<=locMinSlots; b++){ //look before
      if(masterSched[time-b][location][1] != 0){ //if not open
        ok=0;
        shiftLengthSlots=0;
        b = locMinSlots+1;
        if(schedTag==1){
        Logger.log("p %s t %s l %s checkGaps failed lookBefore",person,time,location);
        }
      }
    }
  }
  if(ok==1){ //if look before was okay, look during AND AFTER
    for(var d=1; d<shiftLengthSlots+locMinSlots; d++){ 
      if(masterSched[time+d][location][1] != 0){ //if something there
        if(masterSched[time+d][location][1] == person){ //if same person
          connectAfter = d;
          d = locMinSlots+shiftLengthSlots;
        }
        else{ //if different person
          if(gapMode==1){
            if(d>=shiftLengthSlots){
              shiftLengthSlots = d;
            }
            else{
              shiftLengthSlots = 0;
              if(schedTag==1){
                Logger.log("p %s t %s l %s checkGaps failed lookAfterDiffPerson",person,time,location);
              }
            }
          }
          if(gapMode==2){
            if(d<shiftLengthSlots){
              shiftLengthSlots=0;
              if(schedTag==1){
                Logger.log("p %s t %s l %s checkGaps failed lookAfterDiffPerson",person,time,location);
              }
            }
            else{
              if(extend==1){//extend if you can, ignore if not
                for(var c=1; c<=d-shiftLengthSlots+1; c++){
                  var check = 0;
                  if(time+shiftLengthSlots+c-1<masterSched.length){
                    check = isFreeForShift(person, time+shiftLengthSlots, location, freeList, c/slotSize, slotSize);
                  }
                  if(check==0){
                    if(c>1){
                      shiftLengthSlots = shiftLengthSlots+c-1;
                      c=d-shiftLengthSlots+2;
                    }
                  }
                }
              }
            }   
          }
          d = locMinSlots+shiftLengthSlots;
        }
      }
      else{//if nothing there
        if(extend==1){
          for(var c=1; c<(maxShift*slotSize)-shiftLengthSlots; c++){
            var check = 0;
            if(time+shiftLengthSlots+c-1<masterSched.length){
              check=isFreeForShift(person, time+shiftLengthSlots, location, freeList, c/slotSize, slotSize);
            }
            if(check==0){
              if(c>1){
                shiftLengthSlots = shiftLengthSlots+c-1;
                c = (maxShift*slotSize)-shiftLengthSlots;
              }
            }
          }
        }
      }
    }
  }    
  if(connectAfter>0){ //if next slot is same person
    var persMaxSlots = maxShift*slotSize;
    for(var a=0; a<=persMaxSlots; a++){
      if(a==persMaxSlots){//if next shift is max length
        shiftLengthSlots=0;//no pass
        if(schedTag==1){
        Logger.log("p %s t %s l %s checkGaps failed connectTooLong",person,time,location);
        }
      }
      else{ //if next shift is not at max length
        if(masterSched[time+connectAfter+a][location][1] != person){ //look thru next shift until it stops being that person
          var combinedSlots = connectAfter+a;//when it's not that person anymore, set the new shift length with the gap and the next shift
          if(combinedSlots>persMaxSlots+1){ //if connecting to the next shift makes too long a shift
            shiftLengthSlots=0; //no pass
            if(schedTag==1){
            Logger.log("p %s t %s l %s checkGaps failed connectTooLong2",person,time,location);
            }
          }
          else{ //if not too long
            shiftLengthSlots = combinedSlots; //otherwise, adjust shift length to include gap
          }
          a = persMaxSlots+1;
        }
      }
    }
  }
  return shiftLengthSlots/slotSize;
}


function checkMeals(person, startTime, shiftLength, mealTable, slotSize){
  var mealIndex = [];
  var mealCounter = -1;
  var mealSize = slotSize/2;
  var shiftLengthSlots = shiftLength*slotSize;
    
  //Look thru shift, gather info on any meals it touches
  for(var c=0; c<shiftLengthSlots; c++){ //for each slot of this shift
    if(mealTable[startTime+c][1]>0){ //if it's during a meal
      if(mealCounter == -1 || mealIndex[mealCounter] != mealTable[startTime+c][1]){ //if this meal isn't on the list yet
        mealIndex.push(mealTable[startTime+c][1]); //add this meal to the list
        mealCounter++;
      }
    }
  }
  //Look thru those meals, find out if there's time in the meals outside the shift
  var ok = 1;
  if(mealCounter>-1){
    for(var m=0; m<mealCounter+1; m++){ //for each meal
      var ok1 = 1;
      for(var a=0; a<mealSize; a++){ //for each slot after the end of the shift
        if(startTime+shiftLengthSlots+a<mealTable.length){
          if(mealTable[startTime+shiftLengthSlots+a][1]!=mealIndex[m]){ //if not the same meal
            ok1 = 0;
          }
        }
      }
      if(ok1 == 0){//if no time after meal
        var ok2 = 1;
        for(var b=1; b<mealSize+1; b++){
          if(mealTable[startTime-b][1]!=mealIndex[m]){ //if not the same meal
             ok2 = 0;
          }
        }
        if(ok2==1){
          ok1 = 1;
        }
      }
      if(ok1==0){
        ok = 0;
        m = mealCounter+1;
      }
    }
  }
  return ok;
}
        
               


//########################################################################################## Deputy Conversion ############################################################################################//
function deputyConvert(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();  
  var Schedule = spreadsheet.getSheetByName("Schedule");
  var oldDeputy = spreadsheet.getSheetByName("Deputy Schedule");
  var sched = Schedule.getDataRange().getValues();
  var schedRows = sched.length;
  var schedCols = sched[0].length;
  
  //Make Deputy tab
  if(oldDeputy != null){
    spreadsheet.deleteSheet(oldDeputy);
  }
  
  var depSched = [["Start Time", "End Time", "Area", "Location", "Employee", "Open"]];
  for(var l=2; l<schedCols; l++){ //for each location
    var lastName = "--";
    var shiftLength = 0;
    var onShift = 0;
    var thisEntry = [0,0,0,0,0]; //each entry: timeindex, locindex, starttime, name, length
    for(var t=2; t<schedRows; t++){ //for each time 
      var thisName = sched[t][l];
      if(thisName != lastName){ //if starting or ending a shift
        if(thisName != "--" && thisName != ""){ //if starting a shift
          if(onShift==1){//if right after another shift
            thisEntry[4]=shiftLength;
            depSched.push(thisEntry);
            thisEntry = [0,0,0,0,0];
            shiftLength = 0;
            onShift = 0;
            lastName = thisName;
          }
          shiftLength = 1;
          onShift = 1;
          thisEntry[0] = t;
          thisEntry[1] = l;
          thisEntry[2] = sched[t][0];
          thisEntry[3] = thisName;
          lastName = thisName;
        }
        else{ //if ending a shift
          if(onShift==1){ //make sure ending a shift and not just switching from "--" to ""
            thisEntry[4]=shiftLength;
            depSched.push(thisEntry);
            thisEntry = [0,0,0,0,0];
            shiftLength = 0;
            onShift = 0;
            lastName = thisName;
          }
          else{ //if fake ending
            thisEntry = [0,0,0,0,0];
            shiftLength = 0;
            onShift = 0;
            lastName = thisName;
          }
        }
      }
      else{ //if not starting or ending
        if(onShift==1){ //if on a shift
          shiftLength++; //shift gets longer
          lastName = thisName;
        }
      }
    }
  }
  
  //now we have a list of shifts in depSched, create actual Deputy spreadsheet
  
  //----make actual spreadsheet----//
  var numShifts = depSched.length;
  var finalDeputy = [depSched[0]];
  for(var s=1; s<numShifts; s++){
    var thisShift = depSched[s];
    
    if(thisShift[3]==depSched[0][thisShift[1]]){
      //if it's just IO, skip this row
    }
    else{
    
      //----calculate start and end times----// 
      var shiftStart = thisShift[2];
      var startTime = 0;
      var day = 0;
      var mon = shiftStart.replace("Mon", "2018-10-15");
      if(mon != shiftStart){
        startTime = mon.concat(":00");
      }
      var tue = shiftStart.replace("Tue", "2018-10-16");
      if(tue != shiftStart){
        startTime = tue.concat(":00");
      }
      var wed = shiftStart.replace("Wed", "2018-10-17");
      if(wed != shiftStart){
        startTime = wed.concat(":00");
      }
      var thu = shiftStart.replace("Thu", "2018-10-18");
      if(thu != shiftStart){
        startTime = thu.concat(":00");
      }
      var fri = shiftStart.replace("Fri", "2018-10-19");
      if(fri != shiftStart){
        startTime = fri.concat(":00");
      }
      var sat = shiftStart.replace("Sat", "2018-10-20");
      if(sat != shiftStart){
        startTime = sat.concat(":00");
      }
      var sun = shiftStart.replace("Sun", "2018-10-21");
      if(sun != shiftStart){
        startTime = sun.concat(":00");
      }
      
      var startMinutes = parseInt(startTime.substring(14,16),10);
      var startHours = parseInt(startTime.substring(11,13),10);
      var addMinutes = thisShift[4]*15;
      var endMinutes = (startMinutes+addMinutes) % 60;
      var addHours = ((startMinutes+addMinutes)/60)-(endMinutes/60);
      var endHours = String(startHours+addHours);
      if(endHours.length<2){
        endHours="0".concat(endHours);
      }
      var endTime = startTime.substring(0,11).concat(endHours,":",endMinutes,":00");
      
      //----Open or Name----//
      var open = "n";
      var name = "";
      
      if(thisShift[3]=="OPEN"){
        open = "y";
      }
      else{
        name = thisShift[3];
      }
      
      //----Location or Area----//
      var location = "";
      var area = "";
      
      if(thisShift[1]==2){
         location = "Hardware Office";
         area = "HWSpec";
      }
      if(thisShift[1]==3 || thisShift[1]==4){
        location = "Tech Troubleshooting Office";
        area = "TTO";
      }
      if(thisShift[1]==5){
        location = "Sage Walk-In Center";
        area = "Sage";
      }
      if(thisShift[1]>5){
        location = "Hardware Office";
        area = "io worker";
      }
      //look up names of locations and areas in deputy
      //make a translation table to take loc > location and area
      
      finalDeputy.push([startTime, endTime, area, location, name, open]);
    }
    
  }
  
  var depRows = finalDeputy.length;
  var Deputy = spreadsheet.insertSheet();
  Deputy.setName("Deputy Schedule");
  var depRange = Deputy.getRange(1,1,depRows,6);
  depRange.setValues(finalDeputy);
}





//############################################################################################ Utility Functions ############################################################################################//
function removeFromFreeList(list, value){
  var listLength = list.length;
  var newList = [list[0],0];
  var counter = 0;
  for(var v=2; v<listLength; v++){
    if(list[v] != value){
      newList.push(list[v]);
      counter++;
    }
  }
  newList[1]=counter;
  return newList;
}


function findLabel(label, row, rename){
  //Searches 'row' for a cell containing 'label'.
  //If found, it renames the cell 'rename' and returns the cell's position.
  //If not found, it returns 0.
  var toReturn = -1;
  var rowLength = row.length;
  for(var labelCell = 0; labelCell<rowLength; labelCell++){
    if(row[labelCell] == label){
      row[labelCell] = rename;
      toReturn = labelCell;
      labelCell = rowLength;      
    }
  }
  return toReturn;
}


function colToArray(data, colNum){
  var cLength = data.length;
  var array = [];
  for(var c = 0; c<cLength; c++){
    array.push(data[c][colNum]);
  } 
  return array;
}


function deleteCol(sheet, data, label){
  var colLabel = -1;
  colLabel = findLabel(label, data[0], label); 
  if(colLabel>=0){
    sheet = sheet.deleteColumn(colLabel+1);
    data = sheet.getDataRange().getValues();
  } 
  return data;
}


function addCol(sheet, data, label, col){
  var colLabel = -1;
  colLabel = findLabel(label, data[0], label);
  if(colLabel == -1){
    sheet = sheet.insertColumnBefore(col);
    var dataRange = sheet.getDataRange();
    data = dataRange.getValues();
    data[0][col-1] = label;
    dataRange.setValues(data);
  }
  return data;
}


function addRow(sheet, data, label, row){
  var rowLabel = -1;
  var column = colToArray(data, 0);
  rowLabel = findLabel(label, column, label);
  if(rowLabel == -1){
    sheet = sheet.insertRowBefore(row);
    var dataRange = sheet.getDataRange();
    data = dataRange.getValues();
    data[row-1][0] = label;
    dataRange.setValues(data);
  }
  return data;
}
  

function addIndexCol(sheet, data, label, col, startAt, startWith){
  var colLabel = -1;
  colLabel = findLabel(label, data[0], label);
  if(colLabel == -1){
    sheet = sheet.insertColumnBefore(col);
    var dataRange = sheet.getDataRange();
    data = dataRange.getValues();
    data[0][col-1] = label;
    var dataCLength = data.length;
    var counter = 0;
    for(var r = startAt-1; r<dataCLength; r++){
      data[r][col-1] = startWith+counter;
      counter++;
    }
    dataRange.setValues(data);
  }
  return data;
}


function addIndexRow(sheet, data, label, row, startAt, startWith){ //2, 3
  var rowLabel = -1;
  var column = colToArray(data, 0);
  rowLabel = findLabel(label, column, label);
  if(rowLabel == -1){
    sheet = sheet.insertRowBefore(row);
    var dataRange = sheet.getDataRange();
    data = dataRange.getValues();
    data[row-1][0] = label;
    var dataRLength = data[0].length;
    var counter = 0;
    for(var c = startAt-1; c<dataRLength; c++){
      data[row-1][c] = startWith+counter;
      counter++;
    }
    dataRange.setValues(data);
  }
  return data;
}