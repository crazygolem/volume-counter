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
    let out = []
    let descendents = root.getDescendants()
    for (let i in descendents) {
      let elt = descendents[i].asElement()
      if (elt != null && elt.getName() == tag) out.push(elt)
    }
    return out
  }

  function asTests(elt) {
    let out = []
    let tests = getElementsByTagName(elt, 'test')
    for (let i in tests) {
      out.push({
        expected: parseInt(tests[i].getAttribute('expect').getValue(), 10),
        input: tests[i].getText()
      })
    }
    return out
  }

  let root = XmlService.parse(HtmlService.createHtmlOutputFromFile(filename).getContent()).getRootElement()
  return {
    grammar: getElementsByTagName(root, 'grammar')[0].getText(),
    selftests: asTests(getElementsByTagName(root, 'selftests')[0])
  }
}


// Logs test result in GAS' debugger logs (Ctrl + Enter or View → Logs)
// In addition, if some tests have failed, an exception is thrown.
function selftest() {
  let rules = loadRulesFromFile('rules') // Load parser rules from file 'rules.html'
  let parser = peggy.generate(rules.grammar)
  let tests = rules.selftests

  let errors = 0
  function test(input, expected) {
    let actual
    try {
      actual = parser.parse(input)
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
    } else {
      Logger.log("[ OK ] %s → %s", input, actual)
    }
  }

  for (let i in tests) {
    test(tests[i].input, tests[i].expected)
  }

  if (errors > 0) {
    throw "Selftest failed with " + errors + " errors."
  }
}

function count(values) {
  let rules = loadRulesFromFile('rules') // Load parser rules from file 'rules.html'
  let parser = peggy.generate(rules.grammar)

  let count = 0

  for (let i = 0; i < values.length; i++) {
    count += parser.parse(values[i])
  }

  Logger.log("Number of volumes: %s", count)
  return count
}
