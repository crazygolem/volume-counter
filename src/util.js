/**
 * Clamp a range:
 * - header cells (in frozen rows/columns) are removed when the start of the range is unbounded
 * - trailing cells (past the last row/column with content) are always removed
 *
 * As ranges must have at least one row and column, returns nothing if the clamping results in an empty range.
 */
function clampRange(range) {
  // Unbounded ranges become bounded after an offset is applied, so we have to call offset only once at the end
  let startRow = 0
  let numRows = range.getNumRows()
  let startCol = 0
  let numCols = range.getNumColumns()

  if (!range.isStartRowBounded()) {
    startRow = range.getSheet().getFrozenRows()
    numRows -= startRow
  }

  if (!range.isStartColumnBounded()) {
    startCol = range.getSheet().getFrozenColumns()
    numCols -= startCol
  }

  if (range.getLastRow() > range.getSheet().getLastRow()) {
    numRows -= range.getLastRow() - range.getSheet().getLastRow()
  }

  if (range.getLastColumn() > range.getSheet().getLastColumn()) {
    numCols -= range.getLastColumn() - range.getSheet().getLastColumn()
  }

  if (numRows < 1 || numCols < 1) {
    return
  }

  return range.offset(startRow, startCol, numRows, numCols)
}
