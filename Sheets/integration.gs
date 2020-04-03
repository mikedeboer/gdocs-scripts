/**
 * @OnlyCurrentDoc Limits the script to only accessing the current sheet.
 */

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: "Generate Velocity Sheet", functionName: "generateVelocity" }
  ];
  spreadsheet.addMenu("Manager Tools", menuItems);
}
