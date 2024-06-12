//Sends an xlxs with all data from the current open report in an email to the current user.
function sendData()
{
  //TODO: write this
}

//Runs auto-analysis and generates charts based on the data. Compiles these into a pdf and sends it in an email to the current user.
function sendAnalysis()
{
  //TODO: write this
}

//Generates Graphs and Charts for the open report and writes them to the analysis sheet.
function generateGraphs()
{
  //TODO: pretty sure this doesn't work
  var sheetActive = SpreadsheetApp.getActiveSpreadsheet();
  var chartSheet = sheetActive.getSheetByName("Chart Visualization");
  var dataSheet = sheetActive.getSheetByName("Data Entry Sheet");
  var charts = chartSheet.getCharts();

  var chart = chartSheet.newChart()
    .setChartType(Charts.ChartType.LINE)
    .setOption("title", "testSheet One Sheet")
    .addRange(dataSheet.getRange(6,5,2,4))
    .setPosition(2, 4, 0, 0)
    .build();

  chartSheet.insertChart(chart);
  SpreadsheetApp.getActive().moveChartToObjectSheet(chart);

  for(var i in charts)
  {
    var chart = charts[i];
    
  }
}

//Runs auto-analysis for the open report and writes it too the analysis sheet.
function analyzeTrends()
{
  //TODO: write this
}
