/**
 * Use this function to display in a cell the volume count of the argument, a cell or a range.
 * E.g.:
 *   =COUNTVOLUMES(C23)
 *   =COUNTVOLUMES(C2:C55)
 *
 * Note: This function does not fully support nesting or being nested, and may return incorrect
 * results in some cases (headers and extra blank rows/columns might not get correctly removed).
 */
function countVolumes(input) {
  var formula = SpreadsheetApp.getActiveRange().getFormula()
  var ref = formula.match(/=\w+\((.*)\)/i)[1].split('!') // there is always a formula; might not be this one

  try {
    var range
    if (ref.length == 1) {
      range = SpreadsheetApp.getActiveSheet().getRange(ref[0])
    }
    else {
      range = SpreadsheetApp.getActive().getRange(ref.join('!'))
    }
    return count(flatten_(cleanupRange_(range, range.getSheet()).getValues()))
  }
  catch (e) {
    // Immediate value or range inside a nested formula
    if (input.map && input[0].map)
      return count(flatten_(input))
    else
      return count([input])
  }
}


/**
 * This function is to be called from a menu, after selecting ranges of volumes.
 */
function uiCount() {
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
  var values = []
  for (var i = 0; i < ranges.length; i++) {
    var range = ranges[i]
    range = cleanupRange_(range, SpreadsheetApp.getActiveSheet())
    values = values.concat(flatten_(range.getValues()))
  }
  
  var result = count(values)
  
  var html = HtmlService.createHtmlOutput(result)
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Number of volumes')
}

// Add an item to the Add-on menu, under a sub-menu whose name is set automatically.
function onOpen() {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Count volumes in selected cells', 'uiCount')
      .addToUi()
}