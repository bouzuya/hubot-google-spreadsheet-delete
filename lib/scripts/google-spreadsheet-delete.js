// Description
//   A Hubot script to delete the specified cell
//
// Configuration:
//   HUBOT_GOOGLE_SPREADSHEET_DELETE_EMAIL
//   HUBOT_GOOGLE_SPREADSHEET_DELETE_KEY
//   HUBOT_GOOGLE_SPREADSHEET_DELETE_SHEET_KEY
//
// Commands:
//   hubot google spreadsheet delete <R1C1> - delete the specified cell
//
// Author:
//   bouzuya <m@bouzuya.net>
//
var config, newClient, parseConfig;

newClient = require('../google-sheet');

parseConfig = require('hubot-config');

config = parseConfig('google-spreadsheet-delete', {
  email: null,
  key: null,
  sheetKey: null
});

module.exports = function(robot) {
  return robot.respond(/google spreadsheet delete R(\d+)C(\d+)$/i, function(res) {
    var client, col, colString, row, rowString, spreadsheet, _, _ref;
    _ref = res.match, _ = _ref[0], rowString = _ref[1], colString = _ref[2];
    row = parseInt(rowString, 10);
    col = parseInt(colString, 10);
    client = newClient({
      email: config.email,
      key: config.key
    });
    spreadsheet = client.getSpreadsheet(config.sheetKey);
    return spreadsheet.getWorksheetIds().then(function(worksheetIds) {
      return spreadsheet.getWorksheet(worksheetIds[0]);
    }).then(function(worksheet) {
      return worksheet.deleteValue({
        row: row,
        col: col
      });
    }).then(function() {
      return res.send('OK!');
    })["catch"](function(e) {
      res.send('hubot-google-spreadsheet-delete: error');
      return robot.logger.error(e);
    });
  });
};
