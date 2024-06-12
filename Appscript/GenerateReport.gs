/*
SCRIPT NOTES:
  Must Have at least 1 regional analysis
  Totals will always be over all the data
*/
//region dictionary
var regionDictionary = {};
//facility
var regionsFacility = [];
//non-facility
var regionsOS = [];
//leave regions empty- will be populated based on your selections in runtime with all facility and staff to be analyzed.
var regions = [];

function InitializeRegions()
{
  var ui = SpreadsheetApp.getUi();
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var regionsRange = sheetActive.getSheetByName("Facility and Environmental").getRange(1, 1, sheetActive.getSheetByName("Facility and Environmental").getMaxRows(), 3);
  //ui.alert(sheetActive.getSheetByName("Facility and Environmental").getMaxRows());
  for(i = 0; i < regionsRange.getValues().length-1; i++)
  {
    if(regionsRange.getValues()[i][0] != "Number" && regionsRange.getValues()[i][0] != "")
    {
      //ui.alert("At position" + String(i) + "Key is " + String(regionsRange.getValues()[i][1]) + " with value " + String(regionsRange.getValues()[i][0]));
      Object.defineProperty(regionDictionary, String(regionsRange.getValues()[i][1]), {value: String(regionsRange.getValues()[i][0])});
      if(regionsRange.getValues()[i][2])
      {
        regionsFacility.push(regionsRange.getValues()[i][1]);
      } else {
        regionsOS.push(regionsRange.getValues()[i][1]);
      }
    }
  }
}

/*
/*    INSERT ELEMENTS
/* This function will template your spreadsheet according to your selections. INSERTVALUES will then insert values into the data columns *acording to this functions formatting of the regions and analysis-types*
InsertElements Dependencies:
  Template has exactly 16 years and goes to column X
  Each Region has all of the following in the following Order:
      //Deep Dose Equivalent (DDE) mR
      //Lens (Eye) Dose Equivalent (LDE) mR
      //Shallow Dose Equivalent (Whole Body) (SDE) mR
      //Shallow Dose Equiavlent (Max Extremity) (SDE) mR
      //Total Organ Dose Equivalent (max organ) [DDE + CDE]
      //Quarters Since Inception
      //Highest Historical Dose (TDE)
  The Totals, and Regions section maintain current formatting.
*/
function InsertElements(newSheet, startingYear, reviewLength, regionsToActivate) 
{
  loadingAction = newSheet.getRange(1, 1);

  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var allSheets = sheetActive.getSheets();

  regions = setRegions(regionsToActivate);

  newSheet.deleteColumns(4*(reviewLength+1) + 1, (4*17)-4*(reviewLength+1));
  
  var firstRange = sheetActive.getSheetByName("Analysis Elements").getRange("D3:D14");

  var range = sheetActive.getSheetByName("Analysis Elements").getRange("D3:D62");
  var rangeNames = sheetActive.getSheetByName("Analysis Elements").getRange("C3:C62");
  //var rangeValues = sheetActive.getSheetByName("Analysis Elements").getRange("E3:E62");
  var rangeCategories = sheetActive.getSheetByName("Analysis Elements").getRange("B3:B62");

  var totalRange = sheetActive.getSheetByName("Analysis Elements").getRange("D63:D74");
  var totalRangeNames = sheetActive.getSheetByName("Analysis Elements").getRange("C63:C74");
  //var totalRangeValues = sheetActive.getSheetByName("Analysis Elements").getRange("E63:E74");

  var elements = [];
  var totalElements = [];

  loadingAction.setValue("Retrieving Your Selections ...");
  for(i=1; i < range.getNumRows()+1; i++)
  {
    if(loadingAction.getValue() != "Retrieving Your Selections ... " + Math.round((i / (range.getNumRows()+1)) * 100) + "%")
      {
        loadingAction.setValue("Retrieving Your Selections ... " + Math.round((i / (range.getNumRows()+1)) * 100) + "%");
      }
    if(range.getCell(i, 1).isChecked())
    {
      var category = rangeCategories.getCell(i, 1).getMergedRanges()[0].getCell(1, 1).getValue();
      elements.push([rangeNames.getCell(i, 1).getValue(), range.getCell(i, 1).getValue(), category]);
    }
  }

  totalCounter = 0;
  for(i=1; i < totalRange.getNumRows()+1; i++)
  {
    if(totalRange.getCell(i, 1).isChecked())
    {
      totalElements.push([totalRangeNames.getCell(i, 1).getValue(), totalRange.getCell(i, 1).getValue(), "Total"]);
      totalCounter = totalCounter + 1;
    }
  }

  loadingAction.setValue("Making Rows...");

  //regions- is like, facilities, personel- elements is ~all~ the analysis options!
  if((regions.length-1)*(elements.length-1) > 0)
  {
    newSheet.insertRowsAfter(3, (regions.length-1)*(elements.length-1));
  } 
  if(totalElements.length > 0)
  {
    newSheet.insertRowsAfter(3, (totalElements.length));
  } 
  if((totalElements.length + (regions.length-1)*(elements.length-1)) <= 0)
  {
    loadingAction.setValue("Null Sheet - No Analysis");
    throw new Error("running this sheet with no analysis? This is an afront to me, presonally. And everything I stand for. I'm offended you've even asked. Flaberghasted. Bemused and Amazed. Going to terminate the script.")
  }


  var lock = LockService.getScriptLock(); lock.waitLock(300000); 
  SpreadsheetApp.flush(); lock.releaseLock();


  var categories = [["DDE mR", []], ["LDE mR", []], ["Whole Body SDE mR", []], ["Max Extremity SDE mR", []], ["Total Organ Dose Equivalent", []]];
  var totalData = [];

  loadingAction.setValue("Compiling Categories...");
  //Load Region Categories
  for(i=0; i < categories.length; i++)
  {
    var check = 0;
    for(j=0; j < elements.length; j++)
    {
      if(categories[i][0] == elements[j][2])
      {
        categories[i][1].push([elements[j][0], elements[j][1]]); // ["Name", ~Opperation~]
      }
    }
  }
  //Load Total Categories
  for(j=0; j < totalElements.length; j++)
  {
    if("Total" == totalElements[j][2])
    {
      totalData.push([totalElements[j][0], totalElements[j][1]]); // ["Name", ~Opperation~]
    }
  }

  finalPosition = 2;
  catCount = 0;

  loadingAction.setValue("Naming Regions...");
  for(r = 0; r < regions.length; r++)
  {
    if(loadingAction.getValue() != "Naming Regions... " + Math.round((r / regions.length) * 100) + "%")
      {
        loadingAction.setValue("Naming Regions... " + Math.round((r / regions.length) * 100) + "%");
      }
    marker = finalPosition+1;
    countMarker = 0;
    
    for(k = 0; k < categories.length; k++)
    {
      
      if(categories[k][1].length != 0)
      {
        countMarker = countMarker + categories[k][1].length;
        catCount++;
        newSheet.getRange(finalPosition+1, 2).setValue(categories[k][0]);
        newSheet.getRange(finalPosition+1, 2, categories[k][1].length, 1).merge();

        if((catCount % 2) != 0)
        {
          newSheet.getRange(finalPosition+1, 2, categories[k][1].length, newSheet.getMaxColumns()-1).setBackgroundColor("#C9C9C9");
        } else {
          newSheet.getRange(finalPosition+1, 2, categories[k][1].length, newSheet.getMaxColumns()-1).setBackgroundColor("#FFFFFF");
        }
        for(h = 0; h < categories[k][1].length; h++)
        {
          newSheet.getRange(finalPosition+1+h, 3).setValue(categories[k][1][h][0]);
        }
        finalPosition = finalPosition + categories[k][1].length;
      }
    }
    newSheet.getRange(marker, 1, countMarker).merge();
    newSheet.getRange(marker, 1).setValue(regions[r]);
  }

  loadingAction.setValue("Naming Totals...");
  if(totalCounter != 0)
  {
    //SpreadsheetApp.getUi().alert(finalPosition+1 + ", is the final position")
    newSheet.getRange(finalPosition+1, 2).setValue("Totals");
    newSheet.getRange(finalPosition+1, 2, totalCounter).merge();
    //SpreadsheetApp.getUi().alert("Final Position is " + (finalPosition+1) + " total counter is " + totalCounter);

    if((catCount % 2) == 0)
    {
      newSheet.getRange(finalPosition+1, 2, totalCounter, newSheet.getLastColumn()).setBackgroundColor("#C9C9C9");
    }
    for(h = 0; h < totalCounter; h++)
    {
      newSheet.getRange(finalPosition+1+h, 3).setValue(totalData[h][0]);
    }
  }

  //loadingAction.setValue("Trimming Rows...");
  //newSheet.deleteRows(finalPosition+totalCounter - 1, newSheet.getMaxRows() - finalPosition - totalCounter + 1);

  newSheet.autoResizeColumn(3.4);
  //newSheet.deleteRows(finalPosition+totalCounter, newSheet.getLastRow()-finalPosition-totalCounter-3);
}

/*
/*    INSERT VALUES
/* This function will insert values into the data columns *acording to this functions formatting of the regions and analysis-types*. Gets the analysis formulae for each row from getCellFormula, getHistoricalFormula and getTotalFormula.
InsertValues Dependencies:
  Data is entered into Data Entry Sheet
  Data starts on the third line.
  Data types are in Column B
  Each Region has all of the following in the following Order:
      //Deep Dose Equivalent (DDE) mR
      //Lens (Eye) Dose Equivalent (LDE) mR
      //Shallow Dose Equivalent (Whole Body) (SDE) mR
      //Shallow Dose Equiavlent (Max Extremity) (SDE) mR
      //Total Organ Dose Equivalent (max organ) [DDE + CDE]
      //Quarters Since Inception
      //Highest Historical Dose (TDE)
  XXX DO NOT LEAVE SpreadsheetApp.getUi().alert() calls in the running of the Insert Elements() script! Will cause the script to fail when the kickstart routine hits it!!
*/
function InsertValues(newSheet, startingYear, reviewLength, regionsToActivate, runtimeStart, pickUp)
{
  //      GLOBAL VARRIABLE SET
  regions = setRegions(regionsToActivate);

  //      LOCAL VARIABLE SET
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var loadingAction = newSheet.getRange(1, 1);
  var numRowsInDataEntry = sheetActive.getSheetByName("Data Entry Sheet").getLastRow()-3;
  //regionRange encapsulates the A column with all regions/personel in the Data Entry Sheet.
  var regionRange = sheetActive.getSheetByName("Data Entry Sheet").getRange(3, 1, numRowsInDataEntry);
  //regionLocations will store a map of the row in which each region/type pairing is listed in format:
  //  Map["RegionName": Map["Type 1": rowOfType1, "Type 2": rowOfType2, etc...], etc...]
  var regionLocations = new Map();

  //      GET REGION/DOSETYPE PAIR LOCATIONS FROM DATA ENTRY SHEET
  loadingAction.setValue("Retrieving Your Data...");
  var regionTracker = "";
  for(i=1; i < regionRange.getNumRows(); i++)
  {
    if(loadingAction.getValue() != "Retrieving Your Data... " + Math.round((i / regionRange.getNumRows()) * 100) + "%")
    {
      loadingAction.setValue("Retrieving Your Data... " + Math.round((i / regionRange.getNumRows()) * 100) + "%");
    }
    if(regionRange.getCell(i, 1).getValue() != "")
    {
      regionTracker = regionRange.getCell(i, 1).getValue();
      
      regionLocations.set(regionTracker, new Map([
        ["DDE mR", 2 + i],
        ["LDE mR", 3 + i],
        ["Whole Body SDE mR", 4 + i],
        ["Max Extremity SDE mR", 5 + i],
        ["Total Organ Dose Equivalent", 6 + i],
        ["Quarters Since Inception", 7 + i],
        ["Highest Historical Dose (TDE)", 8 + i]
      ]));
    }
  }

  //      DEFINE VARIABLES FOR ITTERATION
  // If picking up from an earlier execution will define varriables based on that execution- if this is the first execution will
  //initialize the varriables.
  if(pickUp != undefined && pickUp[0] == "One")
  {
    console.log(pickUp[0]);
    //      LOCAL VARIABLES FOR PLUGGING CUSTOM FORMULAE INTO NEW SHEET
    // Defines ranges to be set with formulae in the new sheet.
    var newRegionRange = newSheet.getRange(3, 1, newSheet.getMaxRows()-2);
    var newTypeRange = newSheet.getRange(3, 2, newSheet.getMaxRows()-2);
    var newAnalysisRange = newSheet.getRange(3, 3, newSheet.getMaxRows()-2);
    var regionTracker = pickUp[3];
    var typeTracker = pickUp[2];
    var rowTracker = pickUp[1];
  } else
  {
    //      LOCAL VARIABLES FOR PLUGGING CUSTOM FORMULAE INTO NEW SHEET
    // Defines ranges to be set with formulae in the new sheet.
    var newRegionRange = newSheet.getRange(3, 1, newSheet.getMaxRows()-2);
    var newTypeRange = newSheet.getRange(3, 2, newSheet.getMaxRows()-2);
    var newAnalysisRange = newSheet.getRange(3, 3, newSheet.getMaxRows()-2);
    var regionTracker = "";
    var typeTracker = "";
    var rowTracker = 1;
  }
  

  //      INSERTING CUSTOM FORUMAE INTO NEW SHEET
  //Will iterate through the E column until there is a formula in each cell.
  while(newSheet.getRange(newSheet.getMaxRows(), 5).getCell(1, 1).getValue() == "")
  {
    //Checkpoint One! Save Data! This bad boy takes a long time to run!
    CheckPoint(runtimeStart, "One", [newSheet.getSheetName(), startingYear, reviewLength, rowTracker, typeTracker, regionTracker]);
    var loadingAction = newSheet.getRange(1, 1);

    //Updates the Region for which its compiling new formula
    if(newRegionRange.getCell(rowTracker, 1).getValue() != "" && regionTracker != newRegionRange.getCell(rowTracker, 1).getValue())
    {
      regionTracker = newRegionRange.getCell(rowTracker, 1).getValue();
    }

    //Updates the Type of Dose for which its compiling new formula
    if(newTypeRange.getCell(rowTracker, 1).getValue() != "" && typeTracker != newTypeRange.getCell(rowTracker, 1).getValue())
    {
      typeTracker = newTypeRange.getCell(rowTracker, 1).getValue();
    }

    //Updates the Progress Bar
    if(loadingAction.getValue() != "Compiling Cells for: " + regionTracker + "..." + Math.round((rowTracker / newSheet.getMaxRows()) * 100) + "%")
    {
      loadingAction.setValue("Compiling Cells for: " + regionTracker + "..." + Math.round((rowTracker / newSheet.getMaxRows()) * 100) + "%");
    }

    //Inserts the Row Values
    if(typeTracker == "Totals" || regionTracker == "Totals")
    {
      //Sets Values for Totals
      regionTracker = "Totals";
      var thisRowValues = [Array.from(Array.from({length: (newSheet.getMaxColumns() - 4)}, (_, i) => i + 1), function (num) {
        return getTotalFormula(4 + num, 2 + rowTracker, newAnalysisRange.getCell(rowTracker, 1).getValue(), startingYear, reviewLength, newSheet)
      })];
      newSheet.getRange(2 + rowTracker, 5, 1, newSheet.getMaxColumns() - 4).setValues(thisRowValues);
      rowTracker++;
    } 
    else 
    {
      // Sets Values for a Row
      newSheet.getRange(2 + rowTracker, 4).setValue(getHistoricalFormula(regionTracker, typeTracker, newAnalysisRange.getCell(rowTracker, 1).getValue(), regionLocations, startingYear))
      var thisRowValues = [Array.from(Array.from({length: (newSheet.getMaxColumns() - 4)}, (_, i) => i + 1), function (num) {
        return getCellFormula(4 + num, 2 + rowTracker, regionTracker, typeTracker, newAnalysisRange.getCell(rowTracker, 1).getValue(), regionLocations, startingYear, reviewLength)
      })];
      newSheet.getRange(2 + rowTracker, 5, 1, newSheet.getMaxColumns() - 4).setValues(thisRowValues);
      rowTracker++;
    }
  }    
}

/*
/*    SET REGIONS
/* Returns a list of regions/staff to analyze from the Data Entry Sheet based on the selections made in Analysis Elements
setRegions Dependancies:
  Must have a Staff Sheet! Formatted exactly as is! 
  regionsFacility and regionsOS on lines 8 and 10 must be accurate!
*/
function setRegions(regionsToActivate)
{
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();

  if(regionsToActivate[0] == true)
  {
    regions = regions.concat(regionsFacility);
  }
  if(regionsToActivate[1] == true)
  {
    regions = regions.concat(regionsOS);
  }
  if(regionsToActivate[2] == true)
  {
    // current staff
    curentStaff = [];
    employeeRange = sheetActive.getSheetByName("Staff").getDataRange();
    for(i = 2; i<=employeeRange.getNumRows(); i++)
    {
      //SpreadsheetApp.getUi().alert(i);
      if(employeeRange.getCell(i, 2).getValue() == "Current")
      {
        curentStaff.push(employeeRange.getCell(i, 1).getValue());
      }
    }
    regions = regions.concat(curentStaff);
  }
  if(regionsToActivate[3] == true)
  {
    // former staff
    formerStaff = [];
    employeeRange = sheetActive.getSheetByName("Staff").getDataRange();
    for(i = 1; i<=employeeRange.getNumRows(); i++)
    {
      if(employeeRange.getCell(i, 2).getValue() == "Former")
      {
        formerStaff.push(employeeRange.getCell(i, 1).getValue());
      }
    }
    regions = regions.concat(formerStaff);
  }

  return regions;
}

/*
/*    GET CELL FORMULA
/* Returns an array of strings for the formula of each column of a row of analysis- based on the type and region of analysis. 
get____Formula Dependencies:
  The Historical Data Column *must* be the same in the Data Entry Sheet and Review - Template Sheet. The First Quarter in the Review and Data Entry sheet must also be the same. Basically, you are allowed to add years onto the end of the Data Sheet- but DO NOT FUCK AROUND WITH THE YEARS OR NUMBER OF COLUMNS.
*/
function getCellFormula(column, row, region, type, calculation, dataEntryRowMap, startingYear, reviewLength)
{
  //      LOCAL VARRIABLES
  var today = new Date();
  var currentYear = today.getFullYear();
  var diffFromDataEntry = Math.abs(2022 - startingYear);
  //Column CF will be parsed from Ints to Strings with Apropriate Column Letter (4 -> "D"), Rows will be correct for region/type pairing
  var columnIndexCF = (column - 1) + 4*(diffFromDataEntry);
  var columnCF = indexToCol(columnIndexCF);
  var rowCF = dataEntryRowMap.get(region).get(type);
  //"RAW" Values are just the different types of analysis with open varriables for the specifics relevant to this particular formula.
  var startingYearCol;

  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var dataEntrySheet = sheetActive.getSheetByName("Data Entry Sheet");
  var tf = dataEntrySheet.createTextFinder(startingYear);
  var all = tf.findAll();
  for(let cell in all)
  {
    if(all[cell].getRow() == 1)
    {
      startingYearCol = all[cell].getColumn();
    }
  }
  var startingYearColCF = indexToCol(startingYearCol);


  var valuesRAW = new Map([
    /*
    Total Dose Since Inception
    Mean Quarterly Dose (1) - Since Inception
    Mean Quarterly Dose (2) - Over Review Period
    Median Quarterly Dose - Since Inception
    Median Quarterly Dose - Over Review Period
    Standard Deviation from (1)
    Standard Deviation from (2)
    Differential from Highest Historical Dose
    Percent of total represented by this quarter
    Percent of review period represented by this quarter
    Quarters Since Inception
    Percent of time represented by this quarter
    */
    [
      `Total Dose Since Inception`,
      `='Data Entry Sheet'!C`+ rowCF +` + SUM('Data Entry Sheet'!D` + rowCF + `:` + columnCF + `` + rowCF + `)`
    ],
    [ 
      `Mean Quarterly Dose (1) - Since Inception`, 
      `=ROUND(('Data Entry Sheet'!C`+ rowCF +` + SUM('Data Entry Sheet'!D` + rowCF + `:` + columnCF + `` + rowCF + `))/'Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception") + `, 3)`
    ],
    [
      `Mean Quarterly Dose (2) - Over Review Period`,
      `=ROUND((SUM('Data Entry Sheet'!` + startingYearColCF + `` + rowCF + `:` + columnCF + `` + rowCF + `))/(1 + 'Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception") + `-` + `'Data Entry Sheet'!` + startingYearColCF + `` + dataEntryRowMap.get(region).get("Quarters Since Inception") + `), 3)`
    ],
    [ 
      `Median Quarterly Dose - Since Inception`, 
      `=MEDIAN('Data Entry Sheet'!D` + rowCF + `:` + columnCF + `` + rowCF + `)`
    ],
    [
      `Median Quarterly Dose - Over Review Period`, 
      `=MEDIAN('Data Entry Sheet'!` + startingYearColCF + `` + rowCF + `:` + columnCF + `` + rowCF + `)`
    ],
    [ 
      `Standard Deviation from (1)`,
      `=IF("` + columnCF + `" = "D", "NA", ROUND(STDEV('Data Entry Sheet'!D` + rowCF + `:` + columnCF + `` + rowCF + `), 3))`
    ],
    [
      `Standard Deviation from (2)`,
      `=IF("` + columnCF + `" = "` + startingYearColCF + `", "NA", ROUND(STDEV('Data Entry Sheet'!` + startingYearColCF + `` + rowCF + `:` + columnCF + `` + rowCF + `), 3))`
    ],
    [ //Done
      `Differential from Highest Historical Dose`, 
      `=ABS('Data Entry Sheet'!` + columnCF + rowCF + `-` + `'Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Highest Historical Dose (TDE)") + `)`
    ],
    [  //Done
      `Percent of total represented by this quarter`,  
      `=ROUND(('Data Entry Sheet'!C`+ rowCF +` + SUM('Data Entry Sheet'!D` + rowCF + `:` + columnCF + `` + rowCF + `))/('Data Entry Sheet'!`  + columnCF + `` + rowCF + `), 3)`
    ],
    [ //Done
      `Percent of review period represented by this quarter`,
      `=ROUND(1/` + reviewLength + `, 3)`
    ],
    [ //Done
      `Quarters Since Inception`,
      `='Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception")
    ],
    [ //Done
      `Percent of time represented by this quarter`, 
      `=ROUND(1/('Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception") + `), 3)`
    ]
  ]);

  return valuesRAW.get(calculation);
}

/*
/*    GET HISTORICAL FORMULA
/* Returns a formula for the historical total up to the start of analysis period.
Same Dependencies as getCellFormula.
*/
function getHistoricalFormula(region, type, calculation, dataEntryRowMap, startingYear)
{
  //      LOCAL VARRIABLES
  var diffFromDataEntry = Math.abs(2022 - startingYear);
  var columnIndexCF = 3 + 4*(diffFromDataEntry);
  var columnCF = indexToCol(columnIndexCF);
  var rowCF = dataEntryRowMap.get(region).get(type);
  //"RAW" Values are just the different types of analysis with open varriables for the specifics relevant to this particular formula.
  var valuesRAW = new Map([
    /*
    Total Dose Since Inception
    Mean Quarterly Dose (1) - Since Inception
    Mean Quarterly Dose (2) - Over Review Period
    Median Quarterly Dose - Since Inception
    Median Quarterly Dose - Over Review Period
    Standard Deviation from (1)
    Standard Deviation from (2)
    Differential from Highest Historical Dose
    Percent of total represented by this quarter
    Percent of review period represented by this quarter
    Quarters Since Inception
    Percent of time represented by this quarter
    */
    [
      `Total Dose Since Inception`,
      `=SUM('Data Entry Sheet'!C` + rowCF + `:` + indexToCol(columnIndexCF - 1) + `` + rowCF + `)`
    ],
    [ 
      `Mean Quarterly Dose (1) - Since Inception`, 
      `=ROUND(SUM('Data Entry Sheet'!C` + rowCF + `:` + indexToCol(columnIndexCF - 1) + `` + rowCF + `)/'Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception") + `, 3)`
    ],
    [
      `Mean Quarterly Dose (2) - Over Review Period`,
      `NA`
    ],
    [ 
      `Median Quarterly Dose - Since Inception`, 
      `NA`
    ],
    [
      `Median Quarterly Dose - Over Review Period`, 
      `NA`
    ],
    [ 
      `Standard Deviation from (1)`,
      `NA`
    ],
    [
      `Standard Deviation from (2)`,
      `NA`
    ],
    [ //Done
      `Differential from Highest Historical Dose`, 
      `NA`
      //`=ABS('Data Entry Sheet'!` + columnCF + rowCF + `-` + `'Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Highest Historical Dose (TDE)") + `)`
    ],
    [  //Done
      `Percent of total represented by this quarter`,  
      `NA`
    ],
    [ //Done
      `Percent of review period represented by this quarter`,
      `NA`
    ],
    [ //Done
      `Quarters Since Inception`,
      `='Data Entry Sheet'!` + columnCF + dataEntryRowMap.get(region).get("Quarters Since Inception")
    ],
    [ //Done
      `Percent of time represented by this quarter`, 
      `NA`
    ]
  ]);

  return valuesRAW.get(calculation);
}

/*
/*    GET TOTAL FORMULA
/* Returns an array of strings for the formula of each column of a row of totals analysis. 
Same Dependencies as getCellFormula.
*/
function getTotalFormula(column, row, calculation, startingYear, reviewLength, newSheet)
{
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var dataEntrySheet = sheetActive.getSheetByName("Data Entry Sheet");
  var loadingAction = newSheet.getRange(1, 1);

  //      CONFIG VARRIABLES
  // Dose threshhold will set what is considered a measurable dose at your facility.
  var doseThreshhold = 0.3
  //      LOCAL VARRIABLES
  var today = new Date();
  var currentYear = today.getFullYear();
  var diffFromDataEntry = Math.abs(2022 - startingYear);
  //Column CF will be parsed from Ints to Strings with Apropriate Column Letter (4 -> "D"), Rows will be correct for region/type pairing
  var columnIndexCF = (column - 1) + 4*(diffFromDataEntry);
  var columnCF = indexToCol(columnIndexCF);
  var rowCF = row;
  //"RAW" Values are just the different types of analysis with open varriables for the specifics relevant to this particular formula.

  //      LOCAL VARIABLE SET FOR TOTALS CALCULATIONS
  var totTED ="";
  var reviewPeriod = "";
  var highestHistoricalDose = "";
  var thisYear = "";
  var numPersonel = 0;
  var personel = "";
  var histData = "";

  for(i = 3; i < dataEntrySheet.getLastRow(); i += 7)
  {
    loadingAction.setValue("Compiling Cells for Totals: " + calculation)
    totTED += `'Data Entry Sheet'!D` + (i+4) + `:` + indexToCol(columnCF) + (i+4) + `, `;

    var startingYearCol;
    var tf = dataEntrySheet.createTextFinder(startingYear);
    var all = tf.findAll();
    for(let cell in all)
    {
      if(all[cell].getRow() == 1)
      {
        startingYearCol = all[cell].getColumn()
      }
    }
    reviewPeriod += `'Data Entry Sheet'!` + indexToCol(startingYearCol) + (i+4) + `:` + indexToCol(startingYearCol + 4*reviewLength - 1) + (i+4) + `, `;

    highestHistoricalDose += `'Data Entry Sheet'!` + indexToCol(startingYearCol) + (i+6) + `, `;

    var thisYearCol;
    var tf = dataEntrySheet.createTextFinder(currentYear);
    var all = tf.findAll();
    for(let cell in all)
    {
      if(all[cell].getRow() == 1)
      {
        thisYearCol = all[cell].getColumn()
      }
    }

    thisYear += `'Data Entry Sheet'!` + indexToCol(thisYearCol) + (i+4) + `:` + indexToCol(startingYearCol + 3) + (i+4) + `, `;

    if(sheetActive.getSheetByName("Staff").getRange(2, 1, sheetActive.getSheetByName("Staff").getMaxRows() - 1).getValues().indexOf(dataEntrySheet.getRange(i, 1).getCell(1, 1).getValue()) != -1)
    {
      numPersonel++;
      personel += `'Data Entry Sheet'!D` + (i+4) `:` + columnCF + `` + (i+4) + `, `;
    }

    histData += `'Data Entry Sheet'!C` + (i+4) + `, `;
  }

  var totTED = totTED.slice(0, -2);
  var reviewPeriod = reviewPeriod.slice(0, -2);
  var highestHistoricalDose = highestHistoricalDose.slice(0, -2);
  var thisYear = thisYear.slice(0, -2);
  if(numPersonel != 0)
  {
    var personel = personel.slice(0, -2);
  }
  var histData = histData.slice(0, -2);

  var valuesRAW = new Map([
    /*
    Mean Quarterly Dose (1) - Since Inception
    Mean Quarterly Dose (2) - Over Review Period
    Median Quarterly Dose - Since Inception
    Median Quarterly Dose - Over Review Period
    Standard Deviation from (1)
    Standard Deviation from (2)
    Highest Dose this Quarter
    Differential from Highest Historical Dose
    Number of Monitored Personnel
    Number of Persons Receiving Measurable Dose
    Percent of yearly dose represented by this quarter
    Percent of review period represented by this quarter
    */
    [
      `Mean Quarterly Dose (1) - Since Inception`,
      `=AVERAGE(` + totTED + `) + AVERAGE(` + histData + `)/('Data Entry Sheet'!` + columnCF + `8)`
    ],
    [ 
      `Mean Quarterly Dose (2) - Over Review Period`, 
      `=AVERAGE(` + reviewPeriod + `)`
    ],
    [ 
      `Median Quarterly Dose - Since Inception`,
      `=MEDIAN(` + totTED + `)`
    ],
    [ 
      `Median Quarterly Dose - Over Review Period`, 
      `=MEDIAN(` + reviewPeriod + `)`
    ],
    [
      `Standard Deviation from (1)`, 
      `=STDEV(` + totTED + `)`
    ],
    [ 
      `Standard Deviation from (2)`, 
      `=STDEV(` + reviewPeriod + `)`
    ],
    [ 
      `Highest Dose this Quarter`, 
      `=MAX(` + reviewPeriod + `)`
    ],
    [ 
      `Differential from Highest Historical Dose`, 
      `=MAX(` + highestHistoricalDose + `) - MAX(` + reviewPeriod + `)`
    ],
    [ 
      `Number of Monitored Personnel`, 
      `=` + numPersonel
    ],
    [ 
      `Number of Persons Receiving Measurable Dose`, 
      `=COUNTIF(` + personel + `, ` + doseThreshhold + `)`
    ],
    [ 
      `Percent of yearly dose represented by this quarter`, 
      `=SUM(` + thisYear + `)`
    ],
    [ 
      `Percent of review period represented by this quarter`, 
      `=1/reviewLength`
    ]
  ]);

  return valuesRAW.get(calculation);
}
