var defaultTheme = "solarized-light"
var activeTheme = defaultTheme;

function changeTheme(themeName) {
  saveUserPrefs({
    theme: themeName || defaultTheme
  });
}

/**
 * @returns {Object} the users's preferences
 */
function getUserPrefs() {
  try {
    var userProps = PropertiesService.getUserProperties().getProperties();
  } catch (err) {
    Logger.log(constants.errors.getUserPreferences + ": " + err);
    throw constants.errors.getUserPreferences;
  }

  return {
    theme: userProps.theme || defaultTheme,
    noBackground: userProps.noBackground == "true"
  }
}

/**
 * @param {Object} prefs
 * @param {string} prefs.language
 * @param {string} prefs.theme
 * @param {boolean} prefs.noBackground
 */
function saveUserPrefs(prefs) {
  var curPrefs = getUserPrefs();
  if (!("theme" in prefs)) {
    prefs.theme = curPrefs.theme;
  }
  if (!("noBackground" in prefs)) {
    prefs.noBackground = curPrefs.noBackground;
  }
  PropertiesService.getUserProperties().setProperties(prefs);
}

function formatSelectionAsCode() {
  var selection = getSelection();
  var selectedText = trim(replaceSpecialChars(getTextFromSelection(selection)));
  var highlightedText;
  if (/^```/.test(selectedText) && /```$/.test(selectedText)) {
    Logger.log("Code block starts with def.");
    var m = selectedText.match(/^```[\s]*(:?lang=)?([a-z0-9-]+)[\s\n\r]+/);
    selectedText = selectedText.replace(/^[`]+.*[\n\r]+/, "").replace(/[`]+[\n\r\s]*$/, "");
    if (m && m[2] && hljs.getLanguage(trim(m[2]))) {
      // Specific language defined!
      Logger.log("HIGHLIGHTING with lang:: " + m[2]);
      highlightedText = hljs.highlight(trim(m[2]), selectedText);
    }
  }
  if (!highlightedText) {
    highlightedText = hljs.highlightAuto(selectedText);
  }

  var html = applyTheme(highlightedText.value);
  Logger.log("Selection: " + html);

  try {
    replaceSelection(selection, html)
  } catch (err) {
    Logger.log(constants.errors.insert + ": " + err);
    throw constants.errors.insert;
  }
}

function trim(s) {
  return s.replace(/[\s\n\r]*$/, "").replace(/^[\s\n\r]*/, "");
};

function applyTheme(highlightedText) {
  // Turn <span class="hljs-keyword">var</span> to something that uses inline
  // styles.
  var activeTheme = getUserPrefs().theme;
  Logger.log("Applying theme: " + activeTheme);
  var theme = hljsStyles[activeTheme];
  return ("<div class=\"hljs\">" + highlightedText + "</div>")
    .replace(/class="([a-z0-9\s-_]+)"/gim, function(m, classNames) {
      classNames = trim(classNames).split(/\s+/g);
      var styles, styleNames;
      var styleSet = [];
      for (var i = 0, l = classNames.length; i < l; ++i) {
        styles = theme["." + classNames[i]];
        if (!styles) {
          continue;
        }
        styleNames = Object.keys(styles);
        for (var j = 0, l2 = styleNames.length; j < l2; ++j) {
          styleSet.push(styleNames[j] + ":" + styles[styleNames[j]]);
        }
      }
      return styleSet.length ? "style=\"" + styleSet.join(";") + "\"" : "";
    });
}

/**
 *
 * @param {string} text
 * @returns {string} the text with special characters replaced
 */
function replaceSpecialChars(text) {
  var re = new RegExp(Object.keys(replacements).join("|"), "g");
  return text.replace(re, function getReplacement(match) {
    return replacements[match];
  });
}

const replacements = {
  "\u2018": "'",
  "\u2019": "'",
  "\u201A": "'",
  "\uFFFD": "'",
  "\u201c": "\"",
  "\u201d": "\"",
  "\u201e": "\"",
  "\u02C6": "^",
  "\u2039": "<",
  "\u203A": ">",
  "\u2013": "-",
  "\u2014": "--",
  "\u2026": "...",
  "\u00A9": "(c)",
  "\u00AE": "(r)",
  "\u2122": "TM",
  "\u00BC": "1/4",
  "\u00BD": "1/2",
  "\u00BE": "3/4",
  "\u02DC": " ",
  "\u00A0": " "
};
