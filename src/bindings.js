/**
 * Sum up the number of volumes from a range.
 *
 * Examples:
 *
 *     =COUNTVOLUMES(C23)
 *     =COUNTVOLUMES(C2:C55)
 *     =COUNTVOLUMES(Comics!C:C)
 *     =COUNTVOLUMES(French!C:C;English!C:C)
 *
 * Header rows and columns are automatically excluded if the start of the range
 * is unbounded (e.g. `C:C55`) and the rows/columns are frozen.
 *
 * Since empty values are counted as 1 volume, when given a range with an
 * unbounded end (e.g. `C2:C`), this function excludes cells past the last row
 * or column that contains data. This is a bit finicky, and might not always
 * work, in particular with nested functions.
 */
function countVolumes() {
  // We can't use the argument passed to the function when called from the macro
  // because it contains the actual values resolved from the range. This does
  // allow to detect whether the range is unbounded on either side. So instead
  // we have to parse the formula to get the range specification and work it
  // from there.

  var formula = SpreadsheetApp.getActiveRange().getFormula()
  // There is always a formula (since COUNTVOLUMES was called), but it might
  // not be a simple COUNTVOLUMES call, or not one with a simple range.
  var refs = formula.match(/=.*COUNTVOLUMES\s*\(([^)]*)\)/i)[1]

  var values = refs.split(';')
    .map(ref => {
      try {
        return SpreadsheetApp.getActive().getRange(ref)
      } catch (e) {
        e.message = `${e.message} (${ref})`
        throw e
      }
    })
    .map(range => clampRange(range))
    .map(range => range.getValues())
    .flat(Infinity)

  return count(values)
}


/**
 * This function is to be called from a menu, after selecting ranges of volumes.
 */
function uiCount() {
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
    .map(range => clampRange(range))

  var values = ranges
    .map(range => range.getValues())
    .flat(Infinity)

  var html = `
    <b>${count(values)}</b> volumes
    <p>
    in cells ${ranges.map(range => range.getA1Notation()).join(';')}
  `

  SpreadsheetApp.getUi()
      .showModalDialog(HtmlService.createHtmlOutput(html), 'Number of volumes')
}

// Add an item to the Add-on menu, under a sub-menu whose name is set automatically.
function onOpen() {
  SpreadsheetApp.getUi()
      .createAddonMenu()
      .addItem('Count volumes in selected cells', 'uiCount')
      .addToUi()
}
