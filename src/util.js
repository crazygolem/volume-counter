/**
 * Clamp a range with unbounded end to the last row / column with data.
 */
function clampRange(range) {
  var activeSheet = range.getSheet()

  // Unbounded ranges become bounded after an offset is applied, so we have to call offset only once at the end
  var startRow = 0
  var numRows = range.getNumRows()
  var startCol = 0
  var numCols = range.getNumColumns()

  Logger.log('Initial range: %s [%s +%s : %s +%s]', range.getA1Notation(), startRow, numRows, startCol, numCols)
  Logger.log('Initial bounds: %s %s %s %s', range.isStartRowBounded(), range.isEndRowBounded(), range.isStartColumnBounded(), range.isEndColumnBounded())

  if (!range.isStartRowBounded()) {
    startRow = activeSheet.getFrozenRows()
    numRows -= startRow
  }
  if (!range.isEndRowBounded()) {
    if (range.getLastRow() > activeSheet.getLastRow()) // selection goes beyond last row with content
      numRows -= range.getLastRow() - activeSheet.getLastRow()
  }

  if (!range.isStartColumnBounded()) {
    startCol = activeSheet.getFrozenColumns()
    numCols -= startCol
  }
  if (!range.isEndColumnBounded()) {
    if (range.getLastColumn() > activeSheet.getLastColumn()) // selection goes beyond last column with content
    numCols -= range.getLastColumn() - activeSheet.getLastColumn()
  }

  range = range.offset(startRow, startCol, numRows, numCols)
  Logger.log('Clamped range: %s [%s +%s : %s +%s]', range.getA1Notation(), startRow, numRows, startCol, numCols)
  Logger.log('Clamped bounds: %s %s %s %s', range.isStartRowBounded(), range.isEndRowBounded(), range.isStartColumnBounded(), range.isEndColumnBounded())

  return range
}
