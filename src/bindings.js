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
 * Header cells are automatically excluded if the start of the range is
 * unbounded (e.g. `C:C55`). Only cells in frozen rows and columns are
 * considered as header cells.
 *
 * Trailing cells, i.e. cells in the selection past the last row or column with
 * content, are always excluded. Empty cells that are not trailing are counted
 * as 1 volume.
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
    .filter(x => x != null)
    .map(range => range.getDisplayValues())
    .flat(Infinity)

  return count(values)
}


/**
 * This function is to be called from a menu, after selecting ranges of volumes.
 */
function uiCount() {
  var ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
    .map(range => clampRange(range))
    .filter(x => x != null)

  var values = ranges
    .map(range => range.getDisplayValues())
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
