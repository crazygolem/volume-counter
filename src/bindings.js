/**
 * Sum up the number of volumes from a range.
 *
 * Examples:
 *
 *     =COUNTVOLUMES(C23)
 *     =COUNTVOLUMES(C2:C55)
 *     =COUNTVOLUMES(Comics!C:C)
 *     =COUNTVOLUMES(French!C:C, English!C:C)
 *     =COUNTVOLUMES('Français'!C:C; Anglais!C:C)
 *
 * Header cells are automatically excluded if the start of the range is
 * unbounded (e.g. `C:C55`). Only cells in frozen rows and columns are
 * considered as header cells.
 *
 * Trailing cells, i.e. cells in the selection past the last row or column with
 * content, are always excluded. Empty cells that are not trailing are counted
 * as 1 volume.
 *
 * Multiple arguments (separated by a comma or semicolon depending on the user's
 * locale) are supported, but only references are supported; nested functions,
 * array literals, string concatenation, etc. are not supported. Calling this
 * function multiple times in the same formula (e.g. when nested into other
 * functions) is also not supported.
 */
function countVolumes() {
  // We can't use the arguments passed to the function when called from the
  // macro because it contains the actual values resolved from the range. This
  // does allow to detect whether the range is unbounded on either side. So
  // instead we have to parse the formula to get the range specification and
  // work it from there.
  let refs = SpreadsheetApp
    // There is always a formula (since COUNTVOLUMES was called), but it might
    // not be a simple COUNTVOLUMES call, or not one with supported arguments
    // (i.e. that we can parse or resolve).
    .getActiveRange().getFormula()
    // Extract this function's argument(s). Does not support nested functions.
    // Closing parentheses can only appear in quoted sheet names, and quotes are
    // escaped by doubling them.
    .match(/=.*\bCOUNTVOLUMES\s*\(((?:[^()']*|'(?:[^']|'')*')*)\)/i)[1]
    // Split the arguments. Does not support array literals.
    // The argument separator is locale-dependent: some locales (e.g. US
    // English) use a comma, others (e.g. many European ones) a semicolon.
    .match(/(?:'[^']*(?:''[^']*)*'|[^,;])+/g)
    .map(ref => ref.trim())

  let values = refs
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
    .map(range => range.getDisplayValues()) // => 2D array
    .flat(2)

  return count(values)
}


/**
 * This function is to be called from a menu, after selecting ranges of volumes.
 */
function uiCount() {
  let ranges = SpreadsheetApp.getSelection().getActiveRangeList().getRanges()
    .map(range => clampRange(range))
    .filter(x => x != null)

  let values = ranges
    .map(range => range.getDisplayValues()) // => 2D array
    .flat(2)

  let html = `
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
