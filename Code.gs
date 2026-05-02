function doGet(e) {
  var action = e.parameter.action;
  
  if (action === 'getData') {
    return ContentService.createTextOutput(JSON.stringify(getData()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Hello from Vocabulary Web App GAS Backend!");
}

function doPost(e) {
  var action = e.parameter.action;
  
  if (action === 'markMastered') {
    var word = e.parameter.word;
    return ContentService.createTextOutput(JSON.stringify(markMastered(word)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === 'uploadCSV') {
    var csvData = e.parameter.csvData; // string
    return ContentService.createTextOutput(JSON.stringify(uploadCSV(csvData)))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({error: "Unknown action"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function getData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLearning = ss.getSheetByName("Learning");
  var sheetMastered = ss.getSheetByName("Mastered");
  
  // Create sheets if they don't exist
  if (!sheetLearning) {
    sheetLearning = ss.insertSheet("Learning");
    sheetLearning.appendRow(["Word", "Meaning", "Example", "Status"]);
  }
  if (!sheetMastered) {
    sheetMastered = ss.insertSheet("Mastered");
    sheetMastered.appendRow(["Word", "Meaning", "Example", "Status"]);
  }
  
  var learningData = sheetLearning.getDataRange().getValues();
  var masteredData = sheetMastered.getDataRange().getValues();
  
  var learningList = parseSheetData(learningData);
  var masteredList = parseSheetData(masteredData);
  
  return {
    learning: learningList,
    mastered: masteredList
  };
}

function parseSheetData(data) {
  var list = [];
  if (data.length <= 1) return list;
  var headers = data[0];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = row[j];
    }
    list.push(obj);
  }
  return list;
}

function markMastered(word) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetLearning = ss.getSheetByName("Learning");
  var sheetMastered = ss.getSheetByName("Mastered");
  
  var data = sheetLearning.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === word) {
      var rowToMove = data[i];
      // Update status to "mastered"
      rowToMove[3] = "mastered";
      
      // Append to Mastered tab
      sheetMastered.appendRow(rowToMove);
      
      // Delete from Learning tab (i+1 because spreadsheet rows are 1-indexed)
      sheetLearning.deleteRow(i + 1);
      return {success: true, word: word};
    }
  }
  return {success: false, message: "Word not found"};
}

function uploadCSV(csvText) {
  try {
    var data = Utilities.parseCsv(csvText);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetLearning = ss.getSheetByName("Learning");
    
    // Check if the first row is headers, if it is "Word", skip it
    var startIndex = 0;
    if (data.length > 0 && data[0][0].toLowerCase() === "word") {
      startIndex = 1;
    }
    
    for (var i = startIndex; i < data.length; i++) {
      if (data[i].length > 0 && data[i][0] !== "") {
        sheetLearning.appendRow([data[i][0], data[i][1] || "", data[i][2] || "", "new"]);
      }
    }
    return {success: true, count: data.length - startIndex};
  } catch (err) {
    return {success: false, error: err.toString()};
  }
}
