//config sources and destination -->
var packageSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Package')
var testCell = packageSheet.getRange(2,1).getValue()
var sourceWorkbook = packageSheet.getRange(2,2).getValue()
var destinationWorkbook = packageSheet.getRange(3,2).getValue()
var destinationSheet = packageSheet.getRange(4,2).getValue()
var arraySheetName = packageSheet.getRange(5,2).getValue()
var reusableVariablesSheetName = packageSheet.getRange(6,2).getValue()

//getting array from Array sheet
function makeObjectsArray() {

  var reusableVariablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(reusableVariablesSheetName)
  var reusableVariablesData = reusableVariablesSheet.getDataRange()
  var variablesLastColumn = reusableVariablesData.getLastColumn()
  var variablesLastRow = reusableVariablesData.getLastRow()
  
   //make rangeVariablesArray
  var rangeArrayRange = reusableVariablesSheet.getRange(3,1,variablesLastRow, 2)
  var rangeArrayValues = rangeArrayRange.getValues()
  var rangeVariablesArray = []
    for ( i = 0; i < variablesLastRow -2; i++){
      rangeVariablesArray.push({})
      rangeVariablesArray[i]['name'] = rangeArrayValues[i][0]
      rangeVariablesArray[i]['value'] = rangeArrayValues[i][1]
    }

   //make queryVariablesArray
  var queryArrayRange = reusableVariablesSheet.getRange(3,4,variablesLastRow, 2)
  var queryArrayValues = queryArrayRange.getValues()
  var queryVariablesArray = []
    for ( i = 0; i < variablesLastRow -2; i++){
      queryVariablesArray.push({})
      queryVariablesArray[i]['name'] = queryArrayValues[i][0]
      queryVariablesArray[i]['value'] = queryArrayValues[i][1]
    }
//arrayRangeValues is the referenced array
   //make sourceWBVariablesArray
  var sourceWBArrayRange = reusableVariablesSheet.getRange(3,7,variablesLastRow, 2)
  var sourceWBArrayValues = sourceWBArrayRange.getValues()
  var sourceWBVariablesArray = []
    for ( i = 0; i < variablesLastRow -2; i++){
      sourceWBVariablesArray.push({})
      sourceWBVariablesArray[i]['name'] = sourceWBArrayValues[i][0]
      sourceWBVariablesArray[i]['value'] = sourceWBArrayValues[i][1]
    }
  var arraySheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(arraySheetName)
  var rangeData = arraySheet.getDataRange()
  var lastColumn = rangeData.getLastColumn()
  var lastRow = rangeData.getLastRow()
  var arrayRange = arraySheet.getRange(2,1,lastRow, lastColumn)
  var arrayRangeValues = arrayRange.getValues()

//new dynamicObjectsArray
  var dynamicObjectsArray = []
  for ( i = 0; i < lastRow - 1; i++){
    dynamicObjectsArray.push({})
    dynamicObjectsArray[i]['name'] = arrayRangeValues[i][0]
    loop1:
      for ( j = 0; j < rangeVariablesArray.length; j++) {
      Logger.log(rangeVariablesArray[j])
      if (rangeVariablesArray[j]['name'] === arrayRangeValues[i][1]) {
        dynamicObjectsArray[i]['range'] = rangeVariablesArray[j]['value']
        break loop1
      } else {
      dynamicObjectsArray[i]['range'] = arrayRangeValues[i][1]
      }
      }
    loop2:
      for ( j = 0; j < queryVariablesArray.length; j++) {
      if (queryVariablesArray[j]['name'] === arrayRangeValues[i][2]) {
        dynamicObjectsArray[i]['query'] = queryVariablesArray[j]['value']
        break loop2
      } else {
      dynamicObjectsArray[i]['query'] = arrayRangeValues[i][2]
      }
      }
    loop3:
      for ( j = 0; j < sourceWBVariablesArray.length; j++) {
      if (sourceWBVariablesArray[j]['name'] === arrayRangeValues[i][3]) {
        dynamicObjectsArray[i]['sourceWB'] = sourceWBVariablesArray[j]['value']
        break loop3
      } else {
      dynamicObjectsArray[i]['sourceWB'] = arrayRangeValues[i][3]
      }
      }
    }
    return dynamicObjectsArray
}

//helper function to remove extra rows at bottom of active sheet
function removeEmptyRowsActive(){
  var sh = SpreadsheetApp.getActiveSheet()
  var maxRows = sh.getMaxRows()
  var lastRow = sh.getLastRow()
  if (maxRows !== lastRow) {
    sh.deleteRows(lastRow+1, maxRows-lastRow)
  }
}
//helper function to remove extra rows at bottom of destination sheet
function removeEmptyRowsDest(){
  var sh = SpreadsheetApp.openById(destinationWorkbook).getSheetByName(destinationSheet)
  var maxRows = sh.getMaxRows()
  var lastRow = sh.getLastRow()
  if (maxRows !== lastRow) {
    sh.deleteRows(lastRow+1, maxRows-lastRow)
  }
}

function queryFunction() {
  var giantQuery = ''
  
//logic to make my enormous query import function! -->

  var queryArray = []
//call makeObjectsArray
arrayFromSheet = makeObjectsArray()
// array of objects statement maker
    arrayFromSheet.forEach(function (item, itemIndex) {
//includedOjectsArray.forEach(function (item, itemIndex) {
    queryArray.push('QUERY(IMPORTRANGE("' + item.sourceWB + '", "' + item.name + '!' + item.range + '"), "'+ item.query + '", 0)')
  })
  giantQuery = queryArray.join("; ")
  giantQuery = "={" + giantQuery + "}"
  
  //setFormula in A1 of destination sheet
  SpreadsheetApp.openById(destinationWorkbook).getSheetByName(destinationSheet).getRange('A1').setFormula(giantQuery)
  //trim off empty rows
  removeEmptyRowsDest()
  }

//put ui dropdown to allow setFormula()
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu('Data Import')
      .addItem('Inclusive Data Import', 'queryFunction')
      .addItem('Trim Active Sheet Trailing Rows', 'removeEmptyRowsActive')
      .addToUi()
}