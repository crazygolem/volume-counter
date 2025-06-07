function countRangeList_(rangeList) {
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
  var values = []
  for (var i = 0; i < ranges.length; i++) {
    var range = ranges[i]
    range = cleanupRange_(range, SpreadsheetApp.getActiveSheet())
    values = values.concat(flatten_(range.getValues()))
  }
  return count(values)
}

function cleanupRange_(range, activeSheet) {
  // Unbounded ranges become bounded after an offset is applied, so we have to call offset only once at the end
  var startRow = 0
  var numRows = range.getNumRows()
  var startCol = 0
  var numCols = range.getNumColumns()
  
  Logger.log('Initial range: %s [%s +%s : %s +%s]', range.getA1Notation(), startRow, numRows, startCol, numCols)
  Logger.log('Initial bounds: %s %s %s %s', range.isStartRowBounded(), range.isEndRowBounded(), range.isStartColumnBounded(), range.isEndColumnBounded())
  
  if (!range.isStartRowBounded()) {
    startRow = 1
    numRows -= 1
  }
  if (!range.isEndRowBounded()) {
    if (range.getLastRow() > activeSheet.getLastRow()) // selection goes beyond last row with content
      numRows -= range.getLastRow() - activeSheet.getLastRow()
  }
  
  if (!range.isStartColumnBounded()) {
    startCol = 1
    numCols -= 1
  }
  if (!range.isEndColumnBounded()) {
    if (range.getLastColumn() > activeSheet.getLastColumn()) // selection goes beyond last column with content
    numCols -= range.getLastColumn() - activeSheet.getLastColumn()
  }
  
  range = range.offset(startRow, startCol, numRows, numCols)
  Logger.log('Cleaned range: %s [%s +%s : %s +%s]', range.getA1Notation(), startRow, numRows, startCol, numCols)
  Logger.log('Cleaned bounds: %s %s %s %s', range.isStartRowBounded(), range.isEndRowBounded(), range.isStartColumnBounded(), range.isEndColumnBounded())
  
  return range
}

function flatten_(input) {
  var flattened=[]
  for (var i=0; i<input.length; ++i) {
    var current = input[i]
    for (var j=0; j<current.length; ++j)
      flattened.push(current[j])
  }
  return flattened
}
