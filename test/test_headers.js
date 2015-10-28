var id = '1v1UnZOCXUS7f01Z24YhGy6y0V2vevKrIsnYH3FIt1cU'

var GoogleSpreadsheet = require("../index.js");
var doc = new GoogleSpreadsheet(id)

var creds = require('./test_creds');

doc.useServiceAccountAuth(creds, function (err) {
  doc.getInfo(function (err, info) {
    console.log('got info')
    console.log(err, info)

    var sheet = info.worksheets[0]
    sheet.getRows(function (err, rows) {
      console.log('got rows', rows)
      rows.forEach(function (row) {
        row['is_a'] = row['is_a'] + 'x'
        row.save()
      })
      sheet.addRow({ 'is_a': 10, 'THIS': 'hi', 'test.of': 'x', '1row headers': 'doy' }, function () {
        console.log('done!')
      })
    })
  })
})
