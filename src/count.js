/**********************************************************************************************************************

Script to parse human-readable book volumes lists and count them.

As the format of those lists is too complex for simple regexes, the peggy library (https://peggyjs.org) is used to
parse a grammar that generates an efficient parser for the lists.

**********************************************************************************************************************/


// Libraries loader ///////////////////////////////////////////////////////////////////////////////////////////////////
// Inspired by https://ctrlq.org/code/20380-load-external-javascript-with-eval
(function(libs) {
  Object.keys(libs).forEach(function(name) {
    src = (function(url) { eval(UrlFetchApp.fetch(url).getContentText()) })(libs[name]);
    eval('var ' + name + ' = ' + src);
  })
})({
  // The key will become the name of a variable under which the library will be available
  // The value is the URL to the external library's source
  peggy: "https://unpkg.com/peggy@5.0.3/browser/peggy.min.js"
})


// Rules loader ///////////////////////////////////////////////////////////////////////////////////////////////////////
// It doesn't seem possible to save arbitrary files along Google App Script scripts, so the rules must be stored in an
// XML file (with an HTML extension apparently).
// The file must have the following format:
// - A root element <parser>
// - A single <grammar> element containing the PEG grammar (possibly in a CDATA section)
// - A single <selftests> element containing any number of <test> elements, which in turn
//     - have an 'expect' attribute containing the result of the test
//     - contain the input to parse as only content (possibly in a CDATA section)
function loadRulesFromFile(filename) {
  function getElementsByTagName(root, tag) {
    var out = []
    var descendents = root.getDescendants()
    for (var i in descendents) {
      var elt = descendents[i].asElement()
      if (elt != null && elt.getName() == tag) out.push(elt)
    }
    return out
  }

  function asTests(elt) {
    var out = []
    var tests = getElementsByTagName(elt, 'test')
    for (var i in tests) {
      out.push({
        expected: parseInt(tests[i].getAttribute('expect').getValue(), 10),
        input: tests[i].getText()
      })
    }
    return out
  }

  var root = XmlService.parse(HtmlService.createHtmlOutputFromFile(filename).getContent()).getRootElement()
  return {
    grammar: getElementsByTagName(root, 'grammar')[0].getText(),
    selftests: asTests(getElementsByTagName(root, 'selftests')[0])
  }
}


// Logs test result in GAS' debugger logs (Ctrl + Enter or View → Logs)
// In addition, if some tests have failed, an exception is thrown.
function selftest() {
  var rules = loadRulesFromFile('rules') // Load parser rules from file 'rules.html'
  var parser = peggy.generate(rules.grammar)
  var tests = rules.selftests

  var errors = 0
  function test(input, expected) {
    try {
      var actual = parser.parse(input)
    } catch (e) {
      errors++
      if (e instanceof Error) {
        Logger.log("[FAIL] %s → %s: %s", input, e.name, e.message)
      }
      else {
        Logger.log("[FAIL] %s → %s", input, e)
      }
      return
    }
    if (expected !== actual) {
      errors++
      Logger.log("[FAIL] %s → %s; expected: %s", input, actual, expected)
    }
    else {
      Logger.log("[ OK ] %s → %s", input, actual)
    }
  }

  for (var i in tests) {
    test(tests[i].input, tests[i].expected)
  }

  if (errors > 0) {
    throw "Selftest failed with " + errors + " errors."
  }
}

function count(values) {
  var rules = loadRulesFromFile('rules') // Load parser rules from file 'rules.html'
  var parser = peggy.generate(rules.grammar)

  var count = 0

  for (var i = 0; i < values.length; i++) {
    var s = values[i].replace(/\s/g, '')
    count += parser.parse(s)
  }

  Logger.log("Number of volumes: %s", count)
  return count
}
