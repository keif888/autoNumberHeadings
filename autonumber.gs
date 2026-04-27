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
  numberHeadings(true, (up.skipHeadings.toLowerCase() === "true"), up.skippedLevels, (up.titlesRestartNumbering.toLowerCase() === "true"), false, (up.appendixUsesLettering.toLowerCase() === "true"), up.appendixText);
}
function numberHeadingsRemove() {
  let up = getPreferences();
  numberHeadings(false, (up.skipHeadings.toLowerCase() === "true"), up.skippedLevels, (up.titlesRestartNumbering.toLowerCase() === "true"), false, (up.appendixUsesLettering.toLowerCase() === "true"), up.appendixText);
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
 * @param {boolean} savePrefs Whether to save the current options.
 * @param {boolean} appendixUsesLettering Whether the text from appendixText will trigger letter and number reset
 * @param {string} appendixText The text string that an Appendix is indicated by.  Default = Appendix #:
 * @return {Object} Not implemented: Object containing the resulting text and the result of the
 *     operation (success or error).
 */
function processHeadings(action, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, savePrefs, appendixUsesLettering, appendixText) {
  if (savePrefs) {

    if (skipHeadings) {
      if (skippedLevels.match(/^[1-6,; e and y-]+$/) == null) {
        Logger.log(`skippedLevels is in the wrong format.  Received ${skippedLevels}`);
        return false
      }
      skippedLevels = skippedLevels.replace(/\D/g, '')
    }

    PropertiesService.getDocumentProperties()
      .setProperty('action', action)
      .setProperty('skipHeadings', skipHeadings)
      .setProperty('skippedLevels', skippedLevels)
      .setProperty('titlesRestartNumbering', titlesRestartNumbering)
      .setProperty('selectionOnly', selectionOnly)
      .setProperty('appendixUsesLettering', appendixUsesLettering)
      .setProperty('appendixText', appendixText);
  }

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
      result = numberHeadings(false, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, appendixUsesLettering, appendixText);
      break

    default:
      //
      result = numberHeadings(true, skipHeadings, skippedLevels, titlesRestartNumbering, selectionOnly, appendixUsesLettering, appendixText);
      break
  }
  // const text = getSelectedText().join('\n');
  return result;
}


function numberHeadings(add = false, skipHeadings = false, skippedLevels, titlesRestartNumbering, selectionOnly, appendixUsesLettering = true, appendixText = "Appendix #:") {
  let document = DocumentApp.getActiveDocument();
  let paragraphs = selectionOnly ? document.getSelection().getRangeElements().map(re => re.getElement().asParagraph()) : document.getParagraphs();
  let numbers = [0, 0, 0, 0, 0, 0, 0];
  let appendix = false;
  let appendixHeaders = 'ABCDEFGHIJKLMNOPQRSTUVWXZY';
  let appendixPrefix = 'Appendix ';
  let appendixPostfix = ':';
  // let appendixFind = /^placeholder #:/
  // let appendixHeadingFind = /^(placeholder #:|# )/
  let headingsToProcessRegex = /HEADING\d/
  let before = []
  let after = []

  if (skipHeadings) {
    headingsToProcessRegex = eval('/HEADING[' + skippedLevels + ']/')
  }

  if (appendixUsesLettering && appendixText.length > 0) {
    if (appendixText.includes("#")) {
      appendixPrefix = appendixText.substring(0, appendixText.indexOf("#"));
      if (appendixText.length == appendixText.indexOf("#")) {
        appendixPostfix = " "
      } else {
        appendixPostfix = appendixText.substring(appendixText.indexOf("#")+1);
      }
    } else {
      Logger.log(`Appendix Text string of ${appendixText} is missing the required #`);
      return {
        before: before.join("\n"),
        after: after.join("\n")
      }
    }
  }

  let appendixFind = new RegExp(`^${appendixPrefix}#${appendixPostfix}`);
  let appendixHeadingFind = new RegExp(`^(${appendixPrefix}#${appendixPostfix}|# )`);
  let appendixFindText = `^${appendixPrefix}[A-Z]${appendixPostfix}`;
  let appendixFindHash = `^${appendixPrefix}#${appendixPostfix}`;
  let appendixReplaceHash = `${appendixPrefix}#${appendixPostfix}`;

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
    element.replaceText("^[0-9]+(\\.[0-9]+)*\\. ", "# ")
    if (appendixUsesLettering) {
      element.replaceText(appendixFindText, appendixReplaceHash)
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
          numbering += appendixHeaders.substring(numbers[currentLevel]-1, numbers[currentLevel]);
        } else {
          if (currentLevel <= level) {
            if ((appendix && currentLevel > 1) || !appendix) {
              numbering += numbers[currentLevel] + '.';
            }
          } else {
            numbers[currentLevel] = 0;
          }
        }
      }
      if (appendix && level == 1) {
        element.replaceText(appendixFindHash, appendixPrefix + numbering + appendixPostfix)
      } else {
        element.replaceText("^# ", numbering + ' ')
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
    skipHeadings: userProperties.getProperty('skipHeadings'),
    skippedLevels: userProperties.getProperty('skippedLevels'),
    titlesRestartNumbering: userProperties.getProperty('titlesRestartNumbering'),
    appendixUsesLettering: userProperties.getProperty('appendixUsesLettering'),
    appendixText: userProperties.getProperty('appendixText')
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
