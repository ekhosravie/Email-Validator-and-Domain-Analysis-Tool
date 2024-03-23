function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Analysis')
      .addItem('Run Analysis', 'runAnalysis')
      .addItem('Format Sheet', 'formatSheet')
      .addItem('Add Pie Chart', 'addPieChart')
      .addToUi();
}

function runAnalysis() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var validated_emails = [];

  // Loop through each row of data (excluding headers)
  for (var i = 1; i < data.length; i++) {
    var email = data[i][0];
    var domain = email.split('@')[1];
    var matched = false;
    
    // Perform keyword matching
    var keywords = sheet.getSheetByName('Keywords for matching').getDataRange().getValues();
    for (var j = 0; j < keywords.length; j++) {
      var keyword = keywords[j][0].toLowerCase();
      if (domain.toLowerCase().includes(keyword)) {
        matched = true;
        break;
      }
    }
    
    // Push analyzed data to array
    validated_emails.push([email, domain, matched ? 'Matched' : 'Not Matched']);
  }
  
  // Update the sheet with the analyzed data
  var resultSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('email addresses');
  resultSheet.getRange(2, 1, validated_emails.length, 3).setValues(validated_emails);
}

function formatSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('email addresses');
  var range = sheet.getRange('B2:B');
  
  // Clear existing formatting
  range.setBackground(null);
  
  // Get values and backgrounds for formatting
  var values = range.getValues();
  var backgrounds = range.getBackgrounds();
  
  // Format cells based on status
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] == 'Not Matched') {
      backgrounds[i][0] = '#FF0000'; // Red for invalid
    } else if (values[i][0] == 'Matched') {
      backgrounds[i][0] = '#00FF00'; // Green for valid
    }
  }
  
  // Apply new backgrounds
  range.setBackgrounds(backgrounds);
}

function addPieChart() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('email addresses');
  var range = sheet.getRange('D2:E'); // Assuming D2:E for pie chart data
  
  var chartBuilder = sheet.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(range)
    .setPosition(5, 5, 0, 0)
    .setOption('title', 'Email Analysis')
    .setOption('width', 300)
    .setOption('height', 300);

  sheet.insertChart(chartBuilder.build());
}
