/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * Creates a menu entry in the Google Docs UI when the document is opened.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  DocumentApp.getUi().createAddonMenu()
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Add Headings Numbers', 'numberHeadingsAdd')
    .addItem('Remove Heading Numbers', 'numberHeadingsRemove')
    .addSeparator()
    .addItem('Promote Headings (H1➙Title ... H6➙H5)', 'increaseHeadingLevels')
    .addItem('Demote Headings (Title➙Title, H1➙H2 ... H6➙Normal)', 'decreaseHeadingLevels')
    .addToUi();

  DocumentApp.getUi().createMenu('Heading Tools')
    .addItem('Open sidebar', 'showSidebar')
    .addSeparator()
    .addItem('Add Headings Numbers', 'numberHeadingsAdd')
    .addItem('Remove Heading Numbers', 'numberHeadingsRemove')
    .addSeparator()
    .addItem('Promote Headings (H1➙Title ... H6➙H5)', 'increaseHeadingLevels')
    .addItem('Demote Headings (Title➙Title, H1➙H2 ... H6➙Normal)', 'decreaseHeadingLevels')

    .addToUi();

}

/**
 * Runs when the add-on is installed.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE.)
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar in the document containing the add-on's user interface.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 */
function showSidebar() {
  const ui = HtmlService.createHtmlOutputFromFile('sidebar')
    .setTitle('Auto number headings');
  DocumentApp.getUi().showSidebar(ui);
}

/**
 * 
 * ////////////////////////////////////////////////////////////////////////////
 *                            MY FUNCTIONS
 * ////////////////////////////////////////////////////////////////////////////
 */

function numberHeadingsAdd() {
  let up = getPreferences();
  if (up.anyHeadings === null) {
    // Assume that if anyHeadings is missing, then there are no preferences
    numberHeadings(true, false, '', false, false, null, true);
  } else {
    numberHeadings(true, (up.skipHeadings.toLowerCase() === "true"), up.skippedLevels, (up.titlesRestartNumbering.toLowerCase() === "true"), (up.selectionOnly.toLowerCase() === "true"), up.styleData, (up.anyHeadings.toLowerCase() === "true"));
  }
}
function numberHeadingsRemove() {
  let up = getPreferences();
  if (up.anyHeadings === null) {
    // Assume that if anyHeadings is missing, then there are no preferences
    numberHeadings(false, false, '', false, false, null, true);
  } else {
    numberHeadings(false, (up.skipHeadings.toLowerCase() === "true"), up.skippedLevels, (up.titlesRestartNumbering.toLowerCase() === "true"), (up.selectionOnly.toLowerCase() === "true"), up.styleData, (up.anyHeadings.toLowerCase() === "true"));
  }
}
function increaseHeadingLevels() {
  changeHeadingLevels("up")
}
function decreaseHeadingLevels() {
  changeHeadingLevels("down")
}


/**
 * Gets the user options and calls the required function to process headings
 *
 * @param {string} action The single word description of the action to perform.
 * @param {boolean} skipHeadings Whether to process all or only on some levels.
 * @param {string} skippedLevels The levels to skip as a comma separated list.
 * @param {boolean} titlesRestartNumbering Whether a Title will reset numbering.
 * @param {object} styleData JSON object with the styling information
 * @param {boolean} anyHeadings Whether to only process Headings starting with # or the defined pattern
 * @return {Object} Not implemented: Object containing the resulting text and the result of the
 *     operation (success or error).
 */
function processHeadings(action, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, styleData, anyHeadings) {
  let result
  switch (action) {
    case 'promote':
      // 
      result = changeHeadingLevels("up", skipHeadings, skippedLevels, selectionOnly);
      break

    case 'demote':
      // 
      result = changeHeadingLevels("down", skipHeadings, skippedLevels, selectionOnly);
      break

    case 'remove':
      // 
      result = numberHeadings(false, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, styleData, anyHeadings);
      break

    case 'save':
      if (skipHeadings) {
        if (skippedLevels.match(/^[1-6,; e and y-]+$/) == null) {
          Logger.log(`skippedLevels is in the wrong format.  Received ${skippedLevels}`);
          throw new Error("skippedLevels is in the wrong format.");
        }
        skippedLevels = skippedLevels.replace(/\D/g, '')
      }

      PropertiesService.getDocumentProperties()
        .setProperty('action', action)
        .setProperty('skipHeadings', skipHeadings)
        .setProperty('skippedLevels', skippedLevels)
        .setProperty('titlesRestartNumbering', titlesRestartNumbering)
        .setProperty('selectionOnly', selectionOnly)
        .setProperty('styleData', JSON.stringify(styleData))
        .setProperty('anyHeadings', anyHeadings);

      result = {
        before: "",
        after: ""
      }

    break

    default:
      //
      result = numberHeadings(true, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, styleData, anyHeadings);
      break
  }
  // const text = getSelectedText().join('\n');
  return result;
}


function numberHeadings(add = false, skipHeadings = false, skippedLevels, titlesRestartNumbering, selectionOnly, styleData, anyHeadings) {
  let document = DocumentApp.getActiveDocument();
  let paragraphs = selectionOnly ? document.getSelection().getRangeElements().map(re => re.getElement().asParagraph()) : document.getParagraphs();
  let numbers = [0, 0, 0, 0, 0, 0, 0];
  let appendix = false;
  let appendixHeaders = 'ABCDEFGHIJKLMNOPQRSTUVWXZY';
  let headingsToProcessRegex = /HEADING\d/
  let before = []
  let after = []

  if (skipHeadings) {
    headingsToProcessRegex = eval('/HEADING[' + skippedLevels + ']/')
  }

  if (styleData === null) {
    styleData = {
        h1style: "number",
        h1breaker: "running-dot",
        h2style: "number",
        h2breaker: "running-dot",
        h3style: "number",
        h3breaker: "running-dot",
        h4style: "number",
        h4breaker: "running-dot",
        h5style: "number",
        h5breaker: "running-dot",
        h6style: "number",
        h6breaker: "running-dot",
        hseparator: "space",
        appendix: false,
        appendixPrefix: "Appendix "
    }
  }

  if (styleData.h1breaker == "running-dot") {
    styleData.h1style = "number";
    styleData.h2style = "number";
    styleData.h2breaker = "running-dot";
    styleData.h3style = "number";
    styleData.h3breaker = "running-dot";
    styleData.h4style = "number";
    styleData.h4breaker = "running-dot";
    styleData.h5style = "number";
    styleData.h5breaker = "running-dot";
    styleData.h6style = "number";
    styleData.h6breaker = "running-dot";
  }
  let ultimateRegex = getRegexStringFromStyle(styleData);

  if (styleData.appendix && styleData.appendixPrefix.length == 0) {
    styleData.appendixPrefix = "Appendix ";
  }
  const appendixPrefix = styleData.appendixPrefix;
  const allPostfix = getSeparator(styleData);

  const appendixFind = new RegExp(`^${appendixPrefix}# `);
  const appendixHeadingFind = new RegExp(`^(${appendixPrefix}# |# )`);
  const appendixFindText = `^${appendixPrefix}[A-Z]${allPostfix}`;
  const appendixFindHash = `^${appendixPrefix}# `;
  const appendixReplaceHash = `${appendixPrefix}# `;

  var headingMatch = "^";


  for (let i in paragraphs) {
    let element = paragraphs[i];
    let text = element.getText() + '';
    let type = element.getHeading() + '';

    if (type === 'TITLE' && titlesRestartNumbering) {
      numbers = [0, 0, 0, 0, 0, 0, 0];
    }

    // exclude everything but headings
    if (!type.match(headingsToProcessRegex)) {
      continue;
    }

    // exclude empty headings (e.g. page breaks generate these)
    if (text.match(/^\s*$/)) {
      continue;
    }

    before.push(element.getText())
    // If I am a Heading, replace the number/letter with the placemarker #
    element.replaceText(ultimateRegex, "# ")
    if (styleData.appendix) {
      element.replaceText(appendixFindText, appendixReplaceHash)
      text = element.getText() + '';
      if (anyHeadings && !(text.startsWith(appendixReplaceHash) || text.startsWith("# "))) {
        element.editAsText().insertText(0, '# ');
      }
    } else {
      text = element.getText() + '';
      if (anyHeadings && !text.startsWith("# ")) {
        element.editAsText().insertText(0, '# ');
      }
    }

    // Remove the # from headings (except Appendicies) if remove and anyHeadings
    if (!add && anyHeadings) {
      element.replaceText("^# ", "");
    }

    text = element.getText() + '';
    
    if (add == true && text.match(appendixHeadingFind)) {
      let level = new RegExp(/HEADING(\d)/).exec(type)[1];
      let numbering = '';
      // Reset numbering if we are the 1st Appendix (only level 1), or the 1st level 1 that isn't an appendix.
      if (level == 1 && text.match(appendixFind) && appendix == false) {
        appendix = true;
        numbers = [0, 0, 0, 0, 0, 0, 0];
      } else if (level == 1 && text.match(appendixFind) == false && appendix == true) {
        appendix = false;
        numbers = [0, 0, 0, 0, 0, 0, 0];
      }

      numbers[level]++;
      for (let currentLevel = 1; currentLevel <= 6; currentLevel++) {
        if (appendix && currentLevel == 1 && level == currentLevel) {
          numbering += convertToAlpha(numbers[currentLevel], true);
        } else {
          if (currentLevel <= level) {
            if ((appendix && currentLevel > 1) || !appendix) {
              numbering += generateHeadingNumber(numbers[currentLevel], styleData, currentLevel);
            }
          } else {
            numbers[currentLevel] = 0;
          }
        }
      }
      if (appendix && level == 1) {
        element.replaceText(appendixFindHash, appendixPrefix + numbering + allPostfix)
      } else {
        element.replaceText("^# ", numbering + allPostfix)
      }
    }
    after.push(element.getText())
  }

  return {
    before: before.join("\n"),
    after: after.join("\n")
  }
}

function changeHeadingLevels(direction = '', skipHeadings = false, skippedLevels, selectionOnly) {
  let document = DocumentApp.getActiveDocument()
  let body = document.getBody()
  let paragraphs = selectionOnly ? document.getSelection().getRangeElements().map(re => re.getElement().asParagraph()) : document.getParagraphs();
  let headingsToProcessRegex = /HEADING\d/
  let before = []
  let after = []

  if (skipHeadings) {
    headingsToProcessRegex = eval('/HEADING[' + skippedLevels + ']/')
  }

  let inserted_paragraph
  for (let i in paragraphs) {
    let current_paragraph = paragraphs[i];
    let text = current_paragraph.getText() + '';
    let type = current_paragraph.getHeading() + '';

    // exclude everything but headings
    if (!type.match(headingsToProcessRegex)) {
      continue;
    }

    // exclude empty headings (e.g. page breaks generate these)
    if (text.match(/^\s*$/)) {
      continue;
    }
    console.log(type)

    before.push(current_paragraph.getText())

    // as integer
    let currentLevel = new RegExp(/HEADING(\d)/).exec(type)[1] * 1;

    let problemCurrentLevel = 6
    let newLevel = currentLevel + 1
    let problemLevelFix = "NORMAL"

    if (direction == "up") {
      problemCurrentLevel = 1
      newLevel = currentLevel - 1
      problemLevelFix = "TITLE"
    }

    let newHeadingLevel = eval("DocumentApp.ParagraphHeading.HEADING" + newLevel)
    if (currentLevel == problemCurrentLevel) {
      newHeadingLevel = eval("DocumentApp.ParagraphHeading." + problemLevelFix)
    }
    let style = {};
    style[DocumentApp.Attribute.HEADING] = newHeadingLevel

    let curr_para_id = body.getChildIndex(current_paragraph)
    let new_paragraph = current_paragraph.copy().setText(" ")
    inserted_paragraph = body.insertParagraph(curr_para_id + 1, new_paragraph).setAttributes(style).merge()

    // current_paragraph.setAttributes(style)
    after.push(inserted_paragraph.getText())
  }

  return {
    before: before.join("\n"),
    after: after.join("\n")
  }
}

/**
 * Gets the stored user preferences for the origin and destination languages,
 * if they exist.
 * This method is only used by the regular add-on, and is never called by
 * the mobile add-on version.
 *
 * @return {Object} The user's origin and destination language preferences, if
 *     they exist.
 */
function getPreferences() {
  const userProperties = PropertiesService.getDocumentProperties();
  return {
    action: userProperties.getProperty('action'),
    styleData: JSON.parse(userProperties.getProperty('styleData')),
    anyHeadings: userProperties.getProperty('anyHeadings'),
    selectionOnly: userProperties.getProperty('selectionOnly'),
    titlesRestartNumbering: userProperties.getProperty('titlesRestartNumbering'),
    skipHeadings: userProperties.getProperty('skipHeadings'),
    skippedLevels: userProperties.getProperty('skippedLevels')
  };
}

/**
 * Converts the supplied number into Roman Numerals.
 * There is no Zero in Roman Numerals!
 *
 * @param  {number} num The number to convert into Roman Numerals
 * @param  {boolean} uppercase Whether the result is to be in upper case
 * @return {string} The resultant Roman Numberal
 **/
function convertToRoman(num, uppercase = false) {
  const lookup = {
    M: 1000, CM: 900, D: 500, CD: 400,
    C: 100, XC: 90, L: 50, XL: 40,
    X: 10, IX: 9, V: 5, IV: 4, I: 1
  };
  let roman = '';
  for (let i in lookup) {
    while (num >= lookup[i]) {
      roman += i;
      num -= lookup[i];
    }
  }
  return uppercase ? roman : roman.toLowerCase();
}

/**
 * Converts the supplied number into a-z (or A-Z if uppercase is true).
 * On reaching 27 it goes to aa, 53 it goes to aaa, etc.
 *
 * @param  {number} num The number to convert into Alphas
 * @param  {boolean} uppercase Whether the result is to be in upper case
 * @return {string} The resultant Alphas
 **/
function convertToAlpha(num, uppercase = false) {
  const validchars = 'abcdefghijklmnopqrstuvwxyz';
  num = Math.abs(num);
  var remainder = num % 26;
  num = Math.floor(num / 26);
  if (remainder == 0 && num > 0) {
    remainder = 26;
    num--;
  }
  let char = validchars.substring(remainder-1, remainder);
  return uppercase ? char.repeat(num + 1).toUpperCase() : char.repeat(num + 1);
}

function generateHeadingNumber(num, style, level) {
  switch (level) {
    case 1: 
      return generateNumber(num, style.h1style, style.h1breaker);
    case 2:
      return generateNumber(num, style.h2style, style.h2breaker);
    case 3: 
      return generateNumber(num, style.h3style, style.h3breaker);
    case 4: 
      return generateNumber(num, style.h4style, style.h4breaker);
    case 5: 
      return generateNumber(num, style.h5style, style.h5breaker);
    case 6: 
      return generateNumber(num, style.h6style, style.h6breaker);
  }
}

function generateNumber(num, levelstyle, levelbreaker) {
  var result = "";

  if (levelbreaker == "open-close-bracket") {
    result += "(";
  }

  switch (levelstyle) {
    case "number":
    case "d-number":
      result += num;
      break;
    case "l-alpha":
      result += convertToAlpha(num, false);
      break
    case "u-alpha":
      result += convertToAlpha(num, true);
      break
    case "l-roman":
      result += convertToRoman(num, false);
      break
    case "u-roman":
      result += convertToRoman(num, true);
      break
  }
  switch (levelbreaker) {
    case "dot":
    case "running-dot":
      result += ".";
      break;
    case "close-bracket":
      result += ")";
      break;
    case "open-close-bracket":
      result += ")";
      break;
    case "dash":
      result += "-";
      break;
    case "colon":
      result += ":";
      break;
    case "semicolon":
      result += ";";
      break;
  }
  return result;
}

/**
 * Returns the actual separator character from the style object
 *
 * @param {string} style  styleData from the preferences
 * @return {string} The character that is the heading separator
 **/

function getSeparator(style) {
  switch (style.hseparator) {
    case "space":
      return " ";
    case "tab":
      return "\t";
    case "dash":
      return "-";
    case "colon":
      return ":";
    case "semicolon":
      return ";";
    default:
      return " ";
  }
}


function getRegexStringFromStyle(style) {
  return "^\\(?(([0-9]+)|([a-z]+)|([A-Z]+))(([\\)\\.\\-:;])\\(?(([0-9]+)|([a-z]+)|([A-Z]+)))*([\\)\\.\\-:;])" + getSeparator(style);
}

function getRegexStringFromStyleV2(style) {
  var styleSettings = {
    leadingOpen: false,
    trailingClose: false,
    trailingDot: false,
    trailingDash: false,
    trailingColon: false,
    trailingSemicolon: false,
    numbers: false,
    lowerAlpha: false,
    upperAlpha: false,
    lowerRoman: false,
    upperRoman: false
  };

  if (style.h1breaker == "running-dot") {
    styleSettings = getStyleSettings(style.h1style, style.h1breaker, styleSettings);
  } else {
    styleSettings = getStyleSettings(style.h1style, style.h1breaker, styleSettings);
    styleSettings = getStyleSettings(style.h2style, style.h2breaker, styleSettings);
    styleSettings = getStyleSettings(style.h3style, style.h3breaker, styleSettings);
    styleSettings = getStyleSettings(style.h4style, style.h4breaker, styleSettings);
    styleSettings = getStyleSettings(style.h5style, style.h5breaker, styleSettings);
    styleSettings = getStyleSettings(style.h6style, style.h6breaker, styleSettings);
  }

  var orNeeded = false;
  var regexString = "^";
  if (styleSettings.leadingOpen) {
    regexString += "\\(?";
  }
  regexString += "(?:";
  if (styleSettings.numbers) {
    regexString += "(?:[0-9]+(?:\\.[0-9]+)*)";  // this handles 1. and 1.1.1.1.1.1., but doesn't handle 1)1)1)1)1)1)  ToDo: Fix
    orNeeded = true;
  }
  if (styleSettings.lowerAlpha) {
    if (orNeeded) regexString += "|";
    // regexString += "(?:[a-z]|(?<az>[a-z])\\k<az>+)";  // re2 doesn't support \k<name> :-(
    regexString += "(?:[a-z]+)"; // this handles a. and aaaaa. but not a.a.  (and . must be of the valid set)  ToDo: Fix
    orNeeded = true;
  }
  if (styleSettings.upperAlpha) {
    if (orNeeded) regexString += "|";
    // regexString += "(?:[A-Z]|(?<AZ>[A-Z])\\k<AZ>+)";
    regexString += "(?:[A-Z]+)";
    orNeeded = true;
  }
  if (styleSettings.lowerRoman) {
    if (orNeeded) regexString += "|";
    // regexString += "(?:\\b(?=[mdclxvi]+\\b)m{0,3}(?:cm|cd|d?c{0,3})(?:xc|xl|l?x{0,3})(?:ix|iv|v?i{0,3})\\b)";  // re2 doesn't support this
    regexString += "(?:[mdclxvi]+)";
    orNeeded = true;
  }
  if (styleSettings.upperRoman) {
    if (orNeeded) regexString += "|";
    // regexString += "(?:\\b(?=[MDCLXVI]+\\b)M{0,3}(?:CM|CD|D?C{0,3})(?:XC|XL|L?X{0,3})(?:IX|IV|V?I{0,3})\\b)";
    regexString += "(?[MDCLXVI]+)";
    orNeeded = true;
  }
  regexString += ")(?:";
  orNeeded = false;
  if (styleSettings.trailingClose) {
    regexString += "\\)";
    orNeeded = true;
  }
  if (styleSettings.trailingDot) {
    if (orNeeded) regexString += "|";
    regexString += "\\.";
    orNeeded = true;
  }
  if (styleSettings.trailingDash) {
    if (orNeeded) regexString += "|";
    regexString += "-";
    orNeeded = true;
  }
  if (styleSettings.trailingColon) {
    if (orNeeded) regexString += "|";
    regexString += ":";
    orNeeded = true;
  }
  if (styleSettings.trailingSemicolon) {
    if (orNeeded) regexString += "|";
    regexString += ";";
  }
  regexString += ")";
  regexString += getSeparator(style);
  return regexString;
}

function getStyleSettings(levelstyle, levelbreaker, styleSettings) {
  switch (levelstyle) {
    case "number":
    case "d-number":
      styleSettings.numbers = true;
      break;
    case "l-alpha":
      styleSettings.lowerAlpha = true;
      break
    case "u-alpha":
      styleSettings.upperAlpha = true;
      break
    case "l-roman":
      styleSettings.lowerRoman = true;
      break
    case "u-roman":
      styleSettings.upperRoman = true;
      break
  }
  switch (levelbreaker) {
    case "dot":
    case "running-dot":
      styleSettings.trailingDot = true;
      break;
    case "close-bracket":
      styleSettings.trailingClose = true;
      break;
    case "open-close-bracket":
      styleSettings.leadingOpen = true;
      styleSettings.trailingClose = true;
      break;
    case "dash":
      styleSettings.trailingDash = true;
      break;
    case "colon":
      styleSettings.trailingColon = true;
      break;
    case "semicolon":
      styleSettings.trailingSemicolon = true;
      break;
  }

  return styleSettings;
}

/*

// Deprecated functions

function getRegexStringFromStyleV1(style) {
  switch (style) {
    case "number":
    case "d-number":
      return "[0-9]+";
    case "l-alpha":
      return "[a-z]{1}|([a-z])\1+";
    case "u-alpha":
      return "[A-Z]+";
    case "l-roman":
      return "m{0,3}(cm|cd|d?c{0,3})(xc|xl|l?x{0,3})(ix|iv|v?i{0,3})";
    case "u-roman":
      return "M{0,3}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})";
  }
}

function getRegexStringFromBreaker(breaker) {
  switch (breaker) {
    case "dot":
    case "running-dot":
      return "\\.";
    case "close-bracket":
      return "\\)";
    case "open-close-bracket":
      return "\\(%s\\)";
    case "dash":
      return "-";
    case "colon":
      return ":";
    case "semicolon":
      return ";";
  }
}

function getRegexStringFromStyleAndBreaker(style, breaker) {
  switch (style) {
    case "l-alpha":
    case "u-alpha":
    default:
      switch (breaker) {
        case "dot":
          return getRegexStringFromStyle(style) + "\\.";
        case "running-dot":
          return getRegexStringFromStyle(style) + "\\.";
        case "close-bracket":
          return getRegexStringFromStyle(style) + "\\)";
        case "open-close-bracket":
          return "\\(" + getRegexStringFromStyle(style) + "\\)";
        case "dash":
          return getRegexStringFromStyle(style) + "-";
        case "colon":
          return getRegexStringFromStyle(style) + ":";
        case "semicolon":
          return getRegexStringFromStyle(style) + ";";
      }
  }
}

*/