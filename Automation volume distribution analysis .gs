 /**
 * -------------------
 * 1. UI
 * -------------------
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Start the analysis') 
      .addItem('For PAR3', 'runFullAnalysis') 
     
      .addSeparator() 

      .addItem('For LYO1', 'runLyonAnalysis') 
      
      .addToUi();
}

 /**
 * -------------------
 * 2, run
 * -------------------
 */
function runFullAnalysis() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  processStepOne(sheet);
  SpreadsheetApp.flush(); 
  
  createLocalPivotTable(sheet);
  SpreadsheetApp.flush(); 

createFormattedCharts(sheet);
  
  SpreadsheetApp.getActive().toast("All done！");
}

 /**
 * -------------------
 * 3. data
 * -------------------
 */
function processStepOne(sheet) {
  var lastRow = sheet.getLastRow();
  sheet.insertColumns(4, 2); 
  sheet.getRange("D1:E1").setValues([["Finger", "Level"]]).setFontWeight("bold");

  var sourceValues = sheet.getRange(2, 3, lastRow - 1, 1).getValues();

  var outputValues = sourceValues.map(function(row) {
    var cellValue = String(row[0]); 
    if (!cellValue) return ["", ""];
    var fingerResult = (cellValue.length >= 6) ? cellValue.substring(0, 6).slice(-2) : "Err";
    var levelResult = (cellValue.substring(0, 2).toUpperCase() === "AD") ? "Lower" : "Upper";
    return [fingerResult, levelResult];
  });

  if (outputValues.length > 0) {
    sheet.getRange(2, 4, outputValues.length, 2).setValues(outputValues);
  }
  sheet.autoResizeColumn(4);
  sheet.autoResizeColumn(5);
}

/**
 * -------------------
 * 4. pivot table
 * -------------------
 */
function createLocalPivotTable(sheet) {
  var sourceDataRange = sheet.getRange(1, 1, sheet.getLastRow(), 7); 
  var maxCols = sheet.getMaxColumns();
  
  if (maxCols >= 9) {
    sheet.getRange(1, 9, sheet.getMaxRows(), maxCols - 9 + 1).clear(); 
    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      sheet.removeChart(charts[i]);
    }
  }

  var pivotTable = sheet.getRange("I1").createPivotTable(sourceDataRange);
  pivotTable.addRowGroup(4).showTotals(true);
  pivotTable.addColumnGroup(5).showTotals(true);
  pivotTable.addColumnGroup(6).showTotals(true);
  pivotTable.addPivotValue(7, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  sheet.setColumnWidth(8, 20); 

  sheet.setColumnWidth(8, 20); 
  
  SpreadsheetApp.flush(); 
  
  var lastColumn = sheet.getLastColumn();

if (lastColumn > 0) {
    sheet.autoResizeColumns(1, lastColumn);
  }

}

 /**
 * -------------------
 * 5. Stacked column chart
 * -------------------
 */
function createFormattedCharts(sheet) {
  
  var rangeX = sheet.getRange("I1:I15"); 

  var rangeSeries1 = sheet.getRange("M1:N15"); 
  
  var chartBuilder1 = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rangeX)
    .addRange(rangeSeries1)
    .setPosition(18, 9, 0, 0) 
    .setOption('isStacked', true)
    
    .setOption('title', 'PAR3 Lower Platform Volume Distribution')
    .setOption('titleTextStyle', { color: '#351c75', fontName: 'Verdana', fontSize: 20, bold: true })
    
    .setOption('hAxis', { title: 'Finger', titleTextStyle: { fontName: 'Verdana', fontSize: 18, color: '#000000' } })
    .setOption('vAxis', { title: 'Total parcels', titleTextStyle: { fontName: 'Verdana', fontSize: 18, color: '#000000' } })
.setOption('legend', { textStyle: { fontName: 'Verdana', fontSize: 18 } })
    .setOption('series', {
      0: { labelInLegend: 'bag', color: '#e69138', dataLabel: 'value', dataLabelPosition: 'center', textStyle: { fontName: 'Verdana', fontSize: 10 } },
      1: { labelInLegend: 'pallet', color: '#b4a7d6', dataLabel: 'value', dataLabelPosition: 'center', textStyle: { fontName: 'Verdana', fontSize: 10 } }
    });
    
  sheet.insertChart(chartBuilder1.build());

  var rangeSeries2 = sheet.getRange("Q1:R15"); 
  
  var chartBuilder2 = sheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(rangeX)
    .addRange(rangeSeries2)
    .setPosition(40, 9, 0, 0) 
    .setOption('isStacked', true)
    
    .setOption('title', 'PAR3 Upper Platform Volume Distribution')
    .setOption('titleTextStyle', { color: '#351c75', fontName: 'Verdana', fontSize: 20, bold: true })
    
    .setOption('hAxis', { title: 'Finger', titleTextStyle: { fontName: 'Verdana', fontSize: 18, color: '#000000' } })
    .setOption('vAxis', { title: 'Total parcels', titleTextStyle: { fontName: 'Verdana', fontSize: 18, color: '#000000' } })
    
    .setOption('legend', { textStyle: { fontName: 'Verdana', fontSize: 18 } })
    .setOption('series', {
0: { labelInLegend: 'bag', color: '#ea9999', dataLabel: 'value', dataLabelPosition: 'center', textStyle: { fontName: 'Verdana', fontSize: 10 } }, 
      1: { labelInLegend: 'pallet', color: '#6fa8dc', dataLabel: 'value', dataLabelPosition: 'center', textStyle: { fontName: 'Verdana', fontSize: 10 } }
    });

  sheet.insertChart(chartBuilder2.build());
}
