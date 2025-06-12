/**
 * Sum up the number of volumes of the argument, a cell, a range or a list of
 * ranges.
 *
 * Examples:
 *
 *   =COUNTVOLUMES("1-3, 6, 8-10")
 *   =COUNTVOLUMES(C23)
 *   =COUNTVOLUMES(C2:C55)
 *   =COUNTVOLUMES(Comics!C:C)
 *   =COUNTVOLUMES(French!C:C;English!C:C)
 *
 * Header rows and columns are automatically excluded if the start of the range
 * is unbounded (e.g. `C:C55`) and the rows/columns are frozen.
 *
 * Since empty values are counted as 1 volume, when given a range with an
 * unbounded end (e.g. `C2:C`), this function excludes cells past the last row
 * or column that contains data. This is a bit finicky, and might not always
 * work, in particular with nested functions.
 */
function countVolumes(input) {
  // When specified as a range, the `input` parameter contains the resolved
  // values. For range with an unbounded end, this goes to the end of the
  // document and not the end of the data.
  // We can't easily discriminate between a value, a range or a nested formula,
  // so we just try assuming it's a range, and if there is an exception we fall
  // back to using the resolved values.
  try {
    var formula = SpreadsheetApp.getActiveRange().getFormula()
    // There is always a formula (since COUNTVOLUMES was called), but it might
    // not be a simple COUNTVOLUMES call, or not one with a simple range.
    var refs = formula.match(/=\w+\((.*)\)/i)[1]

    var values = refs.split(';')
      .map(ref => (ref.includes('!')
        ? SpreadsheetApp.getActive()
        : SpreadsheetApp.getActiveSheet()
      ).getRange(ref))
      .map(range => clampRange(range))
      .map(range => range.getValues())
      .flat(Infinity)

    return count(values)
  }
  catch (e) {
    return (input.flat) ? count(input.flat(Infinity)) : count([input])
  }
}


/**
 * This function is to be called from a menu, after selecting ranges of volumes.
 */
function uiCount() {
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
  var values = ranges
    .map(range => clampRange(range))
    .map(range => range.getValues())
    .flat(Infinity)

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
