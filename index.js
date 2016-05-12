var async = require("async")
var request = require("request")
var xml2js = require("xml2js")
var http = require("http")
var querystring = require("querystring")
var _ = require('lodash')
var GoogleAuth = require("google-auth-library")

var GOOGLE_FEED_URL = "https://spreadsheets.google.com/feeds/"
var GOOGLE_AUTH_SCOPE = ["https://spreadsheets.google.com/feeds"]

// The main class that represents a single sheet
// this is the main module.exports
var GooogleSpreadsheet = function (ss_key, auth_id, options) {
  var google_auth = null
  var visibility = 'public'
  var projection = 'values'

  var auth_mode = 'anonymous'

  var auth_client = new GoogleAuth()
  var jwt_client

  options = options || {}

  var xml_parser = new xml2js.Parser({
    // options carried over from older version of xml2js
    // might want to update how the code works, but for now this is fine
    explicitArray: false,
    explicitRoot: false
  })

  if (!ss_key) throw new Error("Spreadsheet key not provided.")

  this.setAuthAndDependencies = (auth) => {
    google_auth = auth
    if (!options.visibility) {
      visibility = google_auth ? 'private' : 'public'
    }
    if (!options.projection) {
      projection = google_auth ? 'full' : 'values'
    }
  }

  // auth_id may be null
  this.setAuthAndDependencies(auth_id)

  this.headers = {}
  this.headerMap = {}

  // Authentication Methods

  this.setAuthToken = (auth_id) => {
    if (auth_mode === 'anonymous') auth_mode = 'token'
    this.setAuthAndDependencies(auth_id)
  }

  this.useServiceAccountAuth = (creds, cb) => {
    if (typeof creds === 'string') creds = require(creds)
    jwt_client = new auth_client.JWT(creds.client_email, null, creds.private_key, GOOGLE_AUTH_SCOPE, null)
    this.renewJwtAuth(cb)
  }

  this.renewJwtAuth = (cb) => {
    auth_mode = 'jwt'
    jwt_client.authorize((err, token) => {
      if (err) return cb(err)
      this.setAuthToken({
        type: token.token_type,
        value: token.access_token,
        expires: token.expiry_date
      })
      cb()
    })
  }

  // This method is used internally to make all requests
  this.makeFeedRequest = (url_params, method, query_or_data, cb) => {
    var url
    var headers = {}
    if (!cb) cb = () => {}
    if (typeof(url_params) === 'string') {
      // used for edit / delete requests
      url = url_params
    } else if (Array.isArray(url_params)) {
      //used for get and post requets
      url_params.push(visibility, projection)
      url = GOOGLE_FEED_URL + url_params.join("/")
    }

    async.series({
      auth: (step) => {
        if (auth_mode !== 'jwt') return step()
        // check if jwt token is expired
        if (google_auth.expires > +new Date()) return step()
        this.renewJwtAuth(step)
      },
      request: (result, step) => {
        if ( google_auth ) {
          if (google_auth.type === 'Bearer') {
            headers['Authorization'] = 'Bearer ' + google_auth.value
          } else {
            headers['Authorization'] = "GoogleLogin auth=" + google_auth
          }
        }

        if ( method === 'POST' || method === 'PUT' ){
          headers['content-type'] = 'application/atom+xml'
        }

        if ( method === 'GET' && query_or_data ) {
          url += "?" + querystring.stringify( query_or_data )
        }

        var resource = {
          url: url,
          method: method,
          headers: headers,
          body: method === 'POST' || method === 'PUT' ? query_or_data : null
        }

        var handler = (err, response, body) => {
          if (err) {
            return cb(err)
          } else if (response.statusCode === 401) {
            return cb(new Error("Invalid authorization key."))
          } else if (response.statusCode >= 400) {
            return cb(new Error("HTTP error " + response.statusCode + ": " + http.STATUS_CODES[response.statusCode]) + " " + JSON.stringify(body))
          } else if (response.statusCode === 200 && response.headers['content-type'].indexOf('text/html') >= 0) {
            return cb(new Error("Sheet is private. Use authentication or make public. (see https://github.com/lapwinglabs/node-google-spreadsheet#a-note-on-authentication for details)"))
          }

          if (body) {
            xml_parser.parseString(body, function(err, result) {
              if (err) return cb(err)
              cb(null, result, body)
            })
          } else {
            if (err) cb(err)
            else cb(null, true)
          }
        }

        request(resource, handler)
      }
    })
  }

  // public API methods
  this.getInfo = (cb) => {
    this.makeFeedRequest(["worksheets", ss_key], 'GET', null, (err, data, xml) => {
      if (err) return cb(err)
      if (data === true) return cb(new Error('No response to getInfo call'))
      var ss_data = {
        title: data.title["_"],
        updated: data.updated,
        author: data.author,
        worksheets: []
      }
      var worksheets = forceArray(data.entry)
      worksheets.forEach((ws_data) => ss_data.worksheets.push(new SpreadsheetWorksheet(this, ws_data)))
      cb(null, ss_data)
    })
  }

  // NOTE: worksheet IDs start at 1

  this.getRows = (worksheet_id, opts, cb) => {
    // the first row is used as titles/keys and is not included
    if (!this.headers[worksheet_id]) {
      return this.getCells(worksheet_id, { 'max-row': 1, 'return-empty': true }, (err, cells) => {
        if (err) throw err
        this.headers[worksheet_id] = cells.map((cell) => cell.value)
        this.getRows(worksheet_id, opts, cb)
      })
    }

    // opts is optional
    if (typeof opts === 'function') { cb = opts; opts = {} }

    var query = {}
    if (opts.start) query["start-index"] = opts.start
    if (opts.num) query["max-results"] = opts.num
    if (opts.orderby) query["orderby"] = opts.orderby
    if (opts.reverse) query["reverse"] = opts.reverse
    if (opts.query) query['sq'] = opts.query

    this.makeFeedRequest(["list", ss_key, worksheet_id], 'GET', query, (err, data, xml) => {
      if (err) return cb(err)
      if (data === true) return cb(new Error('No response to getRows call'))

      // gets the raw xml for each entry -- this is passed to the row object so we can do updates on it later
      var entries_xml = xml.match(/<entry[^>]*>([\s\S]*?)<\/entry>/g)
      var rows = []
      var entries = forceArray(data.entry)
      var i=0
      entries.forEach((row_data) => {
        rows.push(new SpreadsheetRow(this, this.headers[worksheet_id], row_data, entries_xml[i++]))
      })
      this.headerMap[worksheet_id] = rows[0] && rows[0].headerMap
      cb(null, rows)
    })
  }

  this.addRow = (worksheet_id, data, cb) => {
    if (!this.headerMap[worksheet_id]) {
      return this.getRows(worksheet_id, { 'max-row': 1 }, () => this.addRow(worksheet_id, data, cb))
    }
    var data_xml = '<entry xmlns="http://www.w3.org/2005/Atom" xmlns:gsx="http://schemas.google.com/spreadsheets/2006/extended">' + "\n"
    Object.keys(data).forEach((key) => {
      if (key !== 'id' && key !== 'title' && key !== 'content' && key !== '_links') {
        var prop = this.headerMap[worksheet_id][key]
        data_xml += '<gsx:'+ xmlSafeColumnName(prop) + '>' + xmlSafeValue(data[key]) + '</gsx:'+ xmlSafeColumnName(prop) + '>' + "\n"
      }
    })
    data_xml += '</entry>'
    this.makeFeedRequest(["list", ss_key, worksheet_id], 'POST', data_xml, (err, data, xml) => {
      if (err) return cb(err)
      cb(null, new SpreadsheetRow(this, this.headers[worksheet_id], data, xml))
    })
  }

  this.getCells = (worksheet_id, opts, cb) => {
    if (typeof opts === 'function') { cb = opts; opts = {} }

    // Supported options are:
    // min-row, max-row, min-col, max-col, return-empty
    var query = _.assign({}, opts)

    this.makeFeedRequest(["cells", ss_key, worksheet_id], 'GET', query, (err, data, xml) => {
      if (err) return cb(err)
      if (data === true) return cb(new Error('No response to getCells call'))

      var cells = []
      var entries = forceArray(data['entry'])
      var i = 0
      entries.forEach((cell_data) => cells.push(new SpreadsheetCell(this, worksheet_id, cell_data)))

      cb(null, cells)
    })
  }
}

// Classes
var SpreadsheetWorksheet = function (spreadsheet, data) {
  this.id = data.id.substring( data.id.lastIndexOf("/") + 1 )
  this.title = data.title["_"]
  this.rowCount = data['gs:rowCount']
  this.colCount = data['gs:colCount']
  this.getRows = (opts, cb) => spreadsheet.getRows(this.id, opts, cb)
  this.getCells = (opts, cb) => spreadsheet.getCells(this.id, opts, cb)
  this.addRow =(data, cb) => spreadsheet.addRow(this.id, data, cb)
}

var SpreadsheetRow = function (spreadsheet, headers, data, xml) {
  this['_xml'] = xml
  this.headers = headers
  this.headerMap = {}
  var idx = 0

  Object.keys(data).forEach((key) => {
    var val = data[key]
    if (key.substring(0, 4) === "gsx:") {
      var prop = headers[idx++]
      if(typeof val === 'object' && Object.keys(val).length === 0) val = null

      if (key === "gsx:") this.headerMap[prop] = key.substring(0, 3)
      else this.headerMap[prop] = key.substring(4)

      if (val === 'TRUE') val = true
      else if (val === 'FALSE') val = false
      else if (val !== '' && val !== null && !Number.isNaN(+val)) val = +val
      this[prop] = val
    } else {
      if (key === "id") this[key] = val
      else if (val['_']) this[key] = val['_']
      else if (key === 'link') {
        this['_links'] = []
        val = forceArray(val)
        val.forEach((link) => this['_links'][ link['$']['rel'] ] = link['$']['href'])
      }
    }
  })

  this.toJSON = () => {
    return this.headers.reduce((json, key) => {
      if (this[key] === undefined) return json
      json[key] = this[key]
      return json
    }, {})
  }

  this.save = (cb) => {
    /*
    API for edits is very strict with the XML it accepts
    So we just do a find replace on the original XML.
    It's dumb, but I couldnt get any JSON->XML conversion to work reliably
    */
    var data_xml = this['_xml']
    // probably should make this part more robust?
    data_xml = data_xml.replace('<entry>', "<entry xmlns='http://www.w3.org/2005/Atom' xmlns:gsx='http://schemas.google.com/spreadsheets/2006/extended'>")
      Object.keys(this).forEach((key) => {
        if (~this.headers.indexOf(key) || (key.substr(0,1) !== '_' && typeof this[key] === 'string')) {
          var prop = this.headerMap[key]
          if (!prop) return
          data_xml = data_xml.replace(
            new RegExp('<gsx:' + xmlSafeColumnName(prop) + ">([\\s\\S]*?)</gsx:" + xmlSafeColumnName(prop) + '>'),
            (str) => {
              var replacement = '<gsx:' + xmlSafeColumnName(prop) + '>' + xmlSafeValue(this[key]) + '</gsx:' + xmlSafeColumnName(prop) + '>'
              return replacement
            }
          )
        }
    })
    spreadsheet.makeFeedRequest(this['_links']['edit'], 'PUT', data_xml, cb)
  }

  this.del = (cb) => spreadsheet.makeFeedRequest(this['_links']['edit'], 'DELETE', null, cb)

}

var SpreadsheetCell = function (spreadsheet, worksheet_id, data) {
  this.id = data['id']
  this.row = parseInt(data['gs:cell']['$']['row'])
  this.col = parseInt(data['gs:cell']['$']['col'])
  this.value = data['gs:cell']['_']
  this.numericValue = data['gs:cell']['$']['numericValue']

  this['_links'] = []
  links = forceArray( data.link )
  links.forEach((link) => this['_links'][ link['$']['rel'] ] = link['$']['href'])

  this.setValue = (new_value, cb) => {
    this.value = new_value
    this.save(cb)
  }

  this.save = (cb) => {
    new_value = xmlSafeValue(this.value)
    var edit_id = `https://spreadsheets.google.com/feeds/cells/key/worksheetId/private/full/R${this.row}C${this.col}`
    var data_xml =`
      <entry xmlns='http://www.w3.org/2005/Atom' xmlns:gs='http://schemas.google.com/spreadsheets/2006'>
        <id>${edit_id}</id>
        <link rel="edit" type="application/atom+xml" href="${edit_id}"/>
        <gs:cell row="${this.row}" col="${this.col}" inputValue="${new_value}"/>
      </entry>
    `
    spreadsheet.makeFeedRequest(this['_links']['edit'], 'PUT', data_xml, cb)
  }

  this.del = (cb) => this.setValue('', cb)
}

module.exports = GooogleSpreadsheet

var forceArray = function(val) {
  if ( Array.isArray( val ) ) return val
  if ( !val ) return []
  return [ val ]
}

var xmlSafeValue = function(val){
  if ( val === null || val === undefined ) return ''
  return String(val).replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
}

var xmlSafeColumnName = function(val){
  if (!val) return ''
  return String(val).replace(/[\s_]+/g, '')
      .toLowerCase()
}



