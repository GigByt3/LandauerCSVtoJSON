//**************************************USER Controls and Main Spreadsheet Generation Function!**************************************
/* MakeSpreadsheet is the call point by which generating each new report starts. User and KickStart calls will start there until reaching project completion and sending an email to the current user alerting them that their current spreadsheet is finished.
* DEPENDANCIES:
* CheckBoxes: Each Analysis Element type must have the same number of boxes. If you change rthe number of analysis elements in each section you will have to alter this script accordingly.
* Make Spreadsheet: If you change the number or order of regions/staff you'll need to correct the cascading use of "regionsToActivate".
* If you change the positions of "Baller Name" the name entry box, the starting year, or review length boxes, you will need to correct the checks, however these changes should be confined to this function alone.
*KNOWN BUGS:
*If you don't hit enter after typing a new name you can get the sheet to make two reports with the same name. Easy fix I just, need to finish this project at this point.
*Must be analyzing at least 1 regional analysis (I think)
*Don't attach things to emails for the last 30 seconds of the script executing? Don't ask me, might not still be a bug.
*XXX DO NOT PUT SpreadsheetApp.getUi().alert() calls in the running of the Insert Elements() script! Will cause the script to fail when the kickstart routine hits it!!
*/
//===================================================================================================================================

/*
*   Sets up the Analysis Menu in the Menu Bar.
* Send Full Data Pdf: Sends an xlxs with all data from the current open report in an email to the current user.
* Send Analysis Pdf: Runs auto-analysis and generates charts based on the data. Compiles these into a pdf and sends it in an email to the current user.
* Generate Graphic Analysis: Generates Graphs and Charts for the open report and writes them to the analysis sheet.
* Generate Graphic Analysis: Runs auto-analysis for the open report and writes it too the analysis sheet.
*/
function onOpen(e) 
{
  const menu = SpreadsheetApp.getUi().createMenu("Analysis")
  menu
    .addItem('Port Data', 'ShowPicker')
    .addItem('Generate Report', 'MakeSpreadsheet')
    .addSeparator()
    .addItem('Check All Boxes', 'CheckBoxesAll')
    .addItem('Check Boxes Same As First', 'CheckBoxesSameAsFirst')
    .addItem('Check Box Totals', 'CheckBoxesTotals')
    .addItem('Uncheck All Boxes', 'CheckBoxesUncheck')
    .addToUi();
  InitializeRegions();
}

//Assigned to a button on Analysis Spreadsheet page. Checks all Analysis Type Boxes (not Totals or Region/Staff)
function CheckBoxesAll() 
{
    CheckBoxes("All");
}

//Assigned to a button on Analysis Spreadsheet page. Checks all Analysis Type Boxes in the same pattern as the DDE mR section. Will overwrite previous selections accordingly.
function CheckBoxesSameAsFirst() 
{
    CheckBoxes("Same As First");
}

//Checks all Totals.
function CheckBoxesTotals() 
{
    CheckBoxes("Totals");
}

//Uncheck All Boxes (not Region/Staff)
function CheckBoxesUncheck() 
{
    CheckBoxes("Uncheck");
}

//Ports Data into Spreadsheet from a .json file uploaded to your google drive
function ShowPicker(pickUp) {

  const runtimeStart = new Date();

  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var dataSheet = sheetActive.getSheetByName("Data Entry Sheet");
  var loadingAction = dataSheet.getRange(1, 1);

  if(pickUp != undefined && pickUp[0] == "Two")
  {
    var fileName = pickUp[1];
  } else {
    var year = ui.prompt("Year", "Please enter the year, eg. '2024'", ui.ButtonSet.YES_NO);
    var quarter = ui.prompt("Quarter", "Please enter the quarter, eg. '3'", ui.ButtonSet.YES_NO);
    var fileName = ui.prompt("File Name", "Please enter the name of the .json file you uploaded to google drive, eg. 'landauerdata2024quarter3.json'", ui.ButtonSet.YES_NO);
  }

  if((year.getSelectedButton() == ui.Button.YES && quarter.getSelectedButton() == ui.Button.YES && fileName.getSelectedButton() == ui.Button.YES) || (pickUp != undefined && pickUp[0] == "Two"))
  {

    var jsonFileName = fileName.getResponseText();
    jsonFiles = DriveApp.getFilesByName(jsonFileName)

    if(jsonFiles.hasNext())
    {

      // Get JSON Data
      var jsonFile = jsonFiles.next()
      var jsonfilestring = jsonFile.getBlob().getDataAsString();
      var jsonobject = JSON.parse(jsonfilestring);
      var regionDictionary = jsonobject["personel"];
      var regionIndex = {};

      for(var i in regionDictionary)
      {
        regionIndex[regionDictionary[i]["number"]] = i;
      }
      
      if(pickUp != undefined && pickUp[0] == "Two")
      {
        var quarterCol = pickUp[2];
      } else {
        //Get Year Columns
        var yearRange = dataSheet.getRange("1:1");
        var yearCol = yearRange.getValues()[0].indexOf(year.getResponseText());

        //Get quarter Column
        var quarterCol = Number(yearCol) + Number(quarter.getResponseText());
      }
      
      //Enter Data
      var regionWriteRange = dataSheet.getRange(1, quarterCol, dataSheet.getMaxRows(), 1);
      var regionsPersonelRange = dataSheet.getRange("A:A").getValues();
      var regionsTypeRange = dataSheet.getRange("B:B").getValues();

      if(pickUp != undefined && pickUp[0] == "Two")
      {
        var currentRegion = pickUp[3];
      } else {
        var currentRegion = "pangalacticgargleblaster";
      }

      for(i = 0; i < regionsPersonelRange.length; i++)
      {
        CheckPoint(runtimeStart, "Two", [fileName, quarterCol, currentRegion]);

        if(i % 4 == 0)
        {
          loadingAction.setValue("DATA ENTRY SHEET |");
        } else if(i % 4 == 1)
        {
          loadingAction.setValue("DATA ENTRY SHEET /");
        } else if(i % 4 == 2)
        {
          loadingAction.setValue("DATA ENTRY SHEET -");
        } else if(i % 4 == 3)
        {
          loadingAction.setValue("DATA ENTRY SHEET \\");
        }

        var entryNumber = -1;
        var type;
        var result;
        var monitor;
        var hasHands = false;
        //Check if the region is indexed to a number and thus is a region/environemtnal dosimeter
        if(regionIndex[regionsPersonelRange[i][0]] != undefined)
        {
          entryNumber = regionIndex[regionsPersonelRange[i][0]];
          currentRegion = entryNumber;
        } else if(regionIndex[getNumOfName(regionsPersonelRange[i][0])] != undefined){
          entryNumber = regionIndex[getNumOfName(regionsPersonelRange[i][0])];
          currentRegion = entryNumber;
        } else if(currentRegion != "pangalacticgargleblaster" && regionsPersonelRange[i][0] == "") {
          entryNumber = currentRegion;
        } else if(regionsPersonelRange[i][0] != "") {
          currentRegion = "pangalacticgargleblaster";
          continue;
        } else {
          continue;
        }

        //We hate entry 222. Everybody hates 222.
        if(entryNumber == 222)
        {
          continue;
        }
        
        type = regionsTypeRange[i][0];
        if(regionDictionary[entryNumber]["AREA"] != undefined)
        {
          monitor = "AREA"
        } else {
          monitor = "CHEST"
        }
        if(Object.hasOwn(regionDictionary[entryNumber], "LFINGR"))
        {
          hasHands = true;
        }

        switch(type)
        {
          case "Deep Dose Equivalent (DDE) mR":
            result = regionDictionary[entryNumber][monitor]["DDE"];
            break;
          case "Lens (Eye) Dose Equivalent (LDE) mR":
            result = regionDictionary[entryNumber][monitor]["LDE"];
            break;
          case "Shallow Dose Equivalent (Whole Body) (SDE) mR":
            result = regionDictionary[entryNumber][monitor]["SDE"];
            break;
          case "Shallow Dose Equiavlent (Max Extremity) (SDE) mR":
            if(hasHands)
            {
              var lfinger = regionDictionary[entryNumber]["LFINGR"]["SDE"]
              var rfinger = regionDictionary[entryNumber]["RFINGR"]["SDE"]
              if(lfinger == "M" || lfinger == "A")
              {
                lfinger = 0
              }
              if(rfinger == "M" || rfinger == "A")
              {
                rfinger = 0
              }
              if(Number(lfinger) > Number(rfinger))
              {
                result = regionDictionary[entryNumber]["LFINGR"]["SDE"];
              } else {
                result = regionDictionary[entryNumber]["LFINGR"]["SDE"];
              }
            } else {
              result = 0;
            }
            break;
          case "Total Organ Dose Equivalent (max organ) [DDE + CDE]":
            //Change the CDE (0 here) if you expect some commited CDE!
            result = String(Number(regionDictionary[entryNumber][monitor]["DDE"]) + 0);
            if(result = "NaN")
            {
              result = "0"
            }
            break;
          case "Quarters Since Inception":
            result = "=1 +" + indexToCol(quarterCol-1) + String(i + 1);
            break;
          case "Highest Historical Dose (TDE)":
            result = "=MAX(C" + String(i + 1) + ", D" + String(i - 5) + ":" + indexToCol(quarterCol) + String(i - 5) + ")";
            break;
          default:
            result = "unrecognized type";
        }
        if(result == "M")
        {
          regionWriteRange.getCell(i+1, 1).setValue(0);
          regionWriteRange.getCell(i+1, 1).setNote("M");
        } else {
          regionWriteRange.getCell(i+1, 1).setValue(result);
        }
      }
      loadingAction.setValue("DATA ENTRY SHEET");
    } else {
      ui.alert("No such File in your drive.");
    }
  } else {
    ui.alert("Please answer all the dialog questions!")
  }
}

//Checks Boxes based on Selection.
function CheckBoxes(type) 
{
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = sheetActive.getSheets();
  var range = sheetActive.getSheetByName("Analysis Elements").getRange("D3:D62");
  var namerange = sheetActive.getSheetByName("Analysis Elements").getRange("C3:C62");
  var firstRange = sheetActive.getSheetByName("Analysis Elements").getRange("D3:D14");
  var totalRange = sheetActive.getSheetByName("Analysis Elements").getRange("D63:D74");

  if(type == "All") 
  {
    range.check();
  } else if(type == "Same As First") 
  {
    var first = firstRange;
    var firstLen = first.getNumRows();
    for(i=1; i <= range.getNumRows()-firstLen; i++)
    {
      if(range.getCell(i+firstLen, 1).isChecked() && !(firstRange.getCell((1+((i-1)%firstLen)), 1).isChecked()))
      {
        var result = SpreadsheetApp.getUi().alert("Cell E" + i + firstLen +" (" + namerange.getCell(i+firstLen, 1).getValue() + ") is checked but not in the First Range, Uncheck?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
        if(result == SpreadsheetApp.getUi().Button.YES) { range.getCell(i+firstLen, 1).uncheck(); }
      }
      if(firstRange.getCell((1+((i-1)%firstLen)), 1).isChecked())
      {
        range.getCell(i+firstLen, 1).check();
      }
    }
  } else if(type == "Totals") 
  {
    totalRange.check();
  } else if(type == "Uncheck") 
  {
    range.uncheck();
    totalRange.uncheck();
  }
}

//Makes a New Spreadsheet based on selected sections.
function MakeSpreadsheet(pickUp) 
{
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();

  //Runtime Start used for Checkpoint and KickStart Data Loader.
  const runtimeStart = new Date();
  //Get Regions/Staff selections.
  var regionsToActivate = [
      sheetActive.getSheetByName("Analysis Elements").getRange("D75").getValue(),
      sheetActive.getSheetByName("Analysis Elements").getRange("D76").getValue(),
      sheetActive.getSheetByName("Analysis Elements").getRange("D77").getValue(),
      sheetActive.getSheetByName("Analysis Elements").getRange("D78").getValue()
  ];

  //pickUp will have data assigned too it iff MakeSpreadsheet is being called by the DataLoader. When called by the user, it will be undefined.
  if(pickUp != undefined)
  {
    console.log(pickUp[0]);
    newSheetName = pickUp[0];
    //Scrape Pick up Data
    var newSheet = sheetActive.getSheetByName(pickUp[0]);
    loadingAction = newSheet.getRange(1, 1);
    var year = pickUp[1];
    var periodLength = pickUp[2];
    //Call Insert Values based on pickUp Values. Checkpoint will *always* be after insert elements.
    InsertValues(newSheet, year, periodLength, regionsToActivate, runtimeStart, pickUp[3]);
  } else 
  {
    //clears Appscript Latencies? Not sure why but .flush(); helped stop a number of errors. Thanks StackExchange.
    SpreadsheetApp.flush();

    //Check if the user name is identical to existing sheet name.
    if(sheetActive.getSheetByName("Analysis Elements").getRange("C81").getValue() != "Baller Name!")
    {
      SpreadsheetApp.getUi().alert(sheetActive.getSheetByName("Analysis Elements").getRange("C81").getValue());
      return false;
    }

    //Check is year less than the most recent data entry year
    var year = sheetActive.getSheetByName("Analysis Elements").getRange("D80").getValue();
    var maxYear = sheetActive.getSheetByName("Data Entry Sheet").getRange(1, sheetActive.getSheetByName("Data Entry Sheet").getLastColumn()-3).getCell(1,1).getValue()
    if(maxYear < year)
    {
      SpreadsheetApp.getUi().alert("Please Enter a Year less than our Max!");
      return false;
    }

    //Check is user defined an allowable period length.
    var periodLength = parseInt(sheetActive.getSheetByName("Analysis Elements").getRange("D82").getValue());
    if(periodLength == "NaN")
    {
      SpreadsheetApp.getUi().alert("Please Enter Only an Integer for the Review Length");
      return false;
    }
    if(periodLength > maxYear-year+1)
    {
      SpreadsheetApp.getUi().alert("Review Period is too long given the entered Data! Please Enter an allowable Review Period");
      return false;
    }
    

    //Make a New Sheet
    sheetActive.setActiveSheet(sheetActive.getSheetByName("Template"));
    newSheet = sheetActive.duplicateActiveSheet();
    newSheetName = sheetActive.getSheetByName("Analysis Elements").getRange("C80").getValue();
    newSheet.setName(newSheetName);
    loadingAction = newSheet.getRange(1, 1);

    //Insert Elements and Values
    InsertElements(newSheet, year, periodLength, regionsToActivate);
    InsertValues(newSheet, year, periodLength, regionsToActivate, runtimeStart);
  }

  //Name sheet and send update that the doese report has finished compiling.
  //DEV NOTE: The function (MailApp.sendEmail) is weird? Please dont- try to add attachments to emails while this proscess is ongoing. Or- anything. Or download files. I dunno. It crashes chrome. Maybe other browsers.... fixing this is almost certainly beyond my power and possibly gods. glhf -Finlay
  loadingAction.setValue(newSheetName);
  MailApp.sendEmail("#########@reed.edu", "Dose Report: " + newSheetName + " is finished!", "Enjoy!");
}
