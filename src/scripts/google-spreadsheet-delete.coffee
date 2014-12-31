# Description
#   A Hubot script to delete the specified cell
#
# Configuration:
#   HUBOT_GOOGLE_SPREADSHEET_DELETE_EMAIL
#   HUBOT_GOOGLE_SPREADSHEET_DELETE_KEY
#   HUBOT_GOOGLE_SPREADSHEET_DELETE_SHEET_KEY
#
# Commands:
#   hubot google spreadsheet delete <R1C1> - delete the specified cell
#
# Author:
#   bouzuya <m@bouzuya.net>
#
newClient = require '../google-sheet'
parseConfig = require 'hubot-config'

config = parseConfig 'google-spreadsheet-delete',
  email: null
  key: null
  sheetKey: null

module.exports = (robot) ->
  robot.respond /google spreadsheet delete R(\d+)C(\d+)$/i, (res) ->
    [_, rowString, colString] = res.match
    row = parseInt(rowString, 10)
    col = parseInt(colString, 10)
    client = newClient({ email: config.email, key: config.key })
    spreadsheet = client.getSpreadsheet(config.sheetKey)
    spreadsheet.getWorksheetIds()
    .then (worksheetIds) ->
      spreadsheet.getWorksheet(worksheetIds[0])
    .then (worksheet) ->
      worksheet.deleteValue({ row, col })
    .then ->
      res.send 'OK!'
    .catch (e) ->
      res.send 'hubot-google-spreadsheet-delete: error'
      robot.logger.error e
