var request = require("request");
var http = require("http");
var querystring = require("querystring");
var FeedMe = require('feedme');

var FEED_URL = "https://spreadsheets.google.com/feeds/";

var forceArray = function(val) {
  if(Array.isArray(val)) {
    return val;
  }

  return [val];
};

var getFeed = function(params, auth, query, cb) {
  var headers = {};
  var visibility = "public";
  var projection = "values";
  var parser = new FeedMe(true);

  if(auth) {
    headers.Authorization = "GoogleLogin auth=" + auth;
    visibility = "private";
    projection = "full";
  }
  params.push(visibility, projection);

  query = query || {};

  var url = FEED_URL + params.join("/");
  if(query) {
    url += "?" + querystring.stringify(query);
  }

  parser.on('end', function() {
    cb(null, parser.done());
  });

  var req = request.get({
    url: url,
    headers: headers
  })

  req.on('response', function(response) {
    if (response.statusCode == 401) {
      return cb(new Error("Invalid authorization key."));
    }

    if (response.statusCode >= 400) {
      return cb(new Error("HTTP error " + response.statusCode + ": " + http.STATUS_CODES[response.statusCode]));
    }

    response.pipe(parser);
  });
};

var Spreadsheets = module.exports = function(opts, cb) {
  if(!opts) {
    throw new Error("Invalid arguments.");
  }
  if(!opts.key) {
    throw new Error("Spreadsheet key not provided.");
  }

  getFeed(["worksheets", opts.key], opts.auth, null, function(err, data) {
    if(err) {
      return cb(err);
    }

    cb(null, new Spreadsheet(opts.key, opts.auth, data));
  });
};

Spreadsheets.rows = function(opts, cb) {
  if(!opts) {
    throw new Error("Invalid arguments.");
  }
  if(!opts.key) {
    throw new Error("Spreadsheet key not provided.");
  }
  if(!opts.worksheet) {
    throw new Error("Worksheet not specified.");
  }

  var query = {};
  if(opts.start) {
    query["start-index"] = opts.start;
  }
  if(opts.num) {
    query["max-results"] = opts.num;
  }
  if(opts.orderby) {
    query["orderby"] = opts.orderby;
  }
  if(opts.reverse) {
    query["reverse"] = opts.reverse;
  }
  if(opts.sq) {
    query["sq"] = opts.sq;
  }

  getFeed(["list", opts.key, opts.worksheet], opts.auth, query, function(err, data) {
    if(err) {
      return cb(err);
    }

    var rows = [];

    if(typeof data.items != "undefined" && data.items !== null) {
      var entries = forceArray(data.items);

      entries.forEach(function(entry) {
        rows.push(new Row(entry));
      });
    }

    cb(null, rows);
  });
};

Spreadsheets.cells = function(opts, cb) {
  if(!opts) {
    throw new Error("Invalid arguments.");
  }
  if(!opts.key) {
    throw new Error("Spreadsheet key not provided.");
  }
  if(!opts.worksheet) {
    throw new Error("Worksheet not specified.");
  }

  var query = {
  };
  if(opts.range) {
    query["range"] = opts.range;
  }
  if (opts.maxRow) {
    query["max-row"] = opts.maxRow;
  }
  if (opts.minRow) {
    query["min-row"] = opts.minRow;
  }
  if (opts.maxCol) {
    query["max-col"] = opts.maxCol;
  }
  if (opts.minCol) {
    query["min-col"] = opts.minCol;
  }

  getFeed(["cells", opts.key, opts.worksheet], opts.auth, query, function(err, data) {
    if(err) {
      return cb(err);
    }

    if(typeof data.items != "undefined" && data.items !== null) {
      return cb(null, new Cells(data));
    } else {
      return cb(null, { cells: {} }); // Not entirely happy about defining the data format in 2 places (here and in Cells()), but the alternative is moving this check for undefined into that constructor, which means it's in a different place than the one for .rows() above -- and that mismatch is what led to it being missed
    }
  });
};

var Spreadsheet = function(key, auth, data) {
  this.key = key;
  this.auth = auth;
  this.title = data.title;
  this.updated = data.updated;
  this.author = {
    name: data.author.name,
    email: data.author.email
  };

  this.worksheets = [];
  var worksheets = forceArray(data.items);

  worksheets.forEach(function(worksheetData) {
    this.worksheets.push(new Worksheet(this, worksheetData));
  }, this);
};

var Worksheet = function(spreadsheet, data) {
  // This should be okay, unless Google decided to change their URL scheme...
  var id = data.id;
  this.id = id.substring(id.lastIndexOf("/") + 1);
  this.spreadsheet = spreadsheet;
  this.rowCount = data['gs:rowcount'];
  this.colCount = data['gs:colcount'];
  this.title = data.title;
};

Worksheet.prototype.rows = function(opts, cb) {
  opts = opts || {};
  Spreadsheets.rows({
    key: this.spreadsheet.key,
    auth: this.spreadsheet.auth,
    worksheet: this.id,
    start: opts.start,
    num: opts.num,
    sq: opts.sq,
    orderby: opts.orderby,
    reverse: opts.reverse
  }, cb);
};

Worksheet.prototype.cells = function(opts, cb) {
  opts = opts || {};
  Spreadsheets.cells({
    key: this.spreadsheet.key,
    auth: this.spreadsheet.auth,
    worksheet: this.id,
    range: opts.range,
    maxRow: opts.maxRow,
    minRow: opts.minRow,
    maxCol: opts.maxCol,
    minCol: opts.minCol
  }, cb);
};

var Row = function(data) {
  Object.keys(data).forEach(function(key) {
    var val;
    val = data[key];
    if(key.substring(0, 4) == "gsx:")  {
      if(typeof val == 'object' && Object.keys(val).length === 0) {
        val = null;
      }
      if (key == "gsx:") {
        this[key.substring(0, 3)] = val;
      } else {
        this[key.substring(4)] = val;
      }
    } else if(key.substring(0, 4) == "gsx$") {
      if (key == "gsx$") {
        this[key.substring(0, 3)] = val;
      } else {
        this[key.substring(4)] = val.text || val;
      }
    } else {
      if (key == "id") {
        this[key] = val;
      } else if (val.text) {
        this[key] = val.text;
      }
    }
  }, this);
};

var Cells = function(data) {
  // Populate the cell data into an array grid.
  this.cells = {};

  var entries = forceArray(data.items);
  var cell, row, col;
  entries.forEach(function(entry) {
    cell = entry['gs:cell'];
    row = cell.row;
    col = cell.col;

    if(!this.cells[row]) {
      this.cells[row] = {};
    }

    this.cells[row][col] = {
      row: row,
      col: col,
      value: cell.text || ""
    };
  }, this);
};
