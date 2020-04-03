/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the document is open, used to add a custom
 * menu to the document.
 */
function onOpen() {
  var ui = DocumentApp.getUi();
  var subMenu = ui.createMenu("Change theme");
  for (var themeName, i = 0, l = hljsStyles.themeNames.length; i < l; ++i) {
    themeName = hljsStyles.themeNames[i];
    hljsStyles[themeName] = changeTheme.bind(themeName);
    subMenu.addItem(themeName, "hljsStyles." + camelize("change-" + themeName));
  }

  DocumentApp.getUi()
    .createMenu("Code formatter")
    .addItem("Selection as code", "formatSelectionAsCode")
    .addSeparator()
    .addSubMenu(subMenu)
    .addToUi();
}

function camelize(s) {
  s = s.split("-").map(p => p.charAt(0).toUpperCase() + p.substr(1)).join("");
  return s.charAt(0).toLowerCase() + s.substr(1);
}
