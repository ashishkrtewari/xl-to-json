var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

exports = module.exports = XLSX_json;

function XLSX_json(config, callback) {
  if (!config.input) {
    console.error("You miss a input file");
    process.exit(1);
  }

  var cv = new CV(config, callback);

}

function CV(config, callback) {
  var wb = this.load_xlsx(config.input)
  var ws = this.ws(config, wb);
  var csv = this.csv(ws)
  this.cvjson(csv, config.output, config.lowerCaseHeaders, callback)
}

CV.prototype.load_xlsx = function (input) {
  return xlsx.readFile(input);
}

CV.prototype.ws = function (config, wb) {
  var target_sheet = config.sheet;

  if (target_sheet == null)
    target_sheet = wb.SheetNames[0];

  ws = wb.Sheets[target_sheet];
  return ws;
}

CV.prototype.csv = function (ws) {
  return csv_file = xlsx.utils.make_csv(ws)
}

CV.prototype.cvjson = function (csv, output, lowerCaseHeaders, callback) {
  var record = {}
  var header = []
  var tiles = [];
  var breaks = [];
  var rows = [];

  cvcsv()
    .from.string(csv)
    .on('record', function (row, index) {
      rows.push(row);
      findInArray();
      function findInArray(array, term, pushArray) {
        for (var i = 0; i < row.length; i++) {
          if (row[i].indexOf("_tile") > -1) {
            tiles.push(index);
          }
          else if (row[i].indexOf("_break") > -1) {
            breaks.push(index);
          }
        }
      }
    })
    .on('end', function (count) {
      for (var tileIndex = 0; tileIndex < tiles.length; tileIndex++) {
        var currentTile = tiles[tileIndex], 
              nextTile = tiles[tileIndex + 1] || rows.length,
              currentTileValue = rows[currentTile][0].toLowerCase().substr(6);
              record[currentTileValue] = {};
        var breakObj = (function () {
          var obj = {};
          for (var i = 0; i < breaks.length; i++) {
            if (breaks[i] > tiles[tileIndex]) {
              obj.min = i;
              break;
            }
          }
          for (var j = breaks.length; j > 0; j--) {
            if (breaks[j] < tiles[tileIndex + 1]) {
              obj.max = j;
              break;
            }
          }
          obj.max = obj.max || breaks.length - 1;
          return obj;
        })();
        for (var breakIndex = breakObj.min; breakIndex <= breakObj.max; breakIndex++) {
          var currentBreak = breaks[breakIndex],
              nextBreak = breaks[breakIndex + 1] < nextTile ? breaks[breakIndex + 1] : nextTile,
              breakValue = rows[currentBreak][0].substr(7),
              keys = [];

          if (breakValue !== 'transdata') {
            (function(){
              var header = rows[currentBreak + 1];
              var row = rows[currentBreak + 2];
              record[currentTileValue][breakValue] = {};
              header.forEach(function (column, index) {
                var foundCounter = 0;
                var key = lowerCaseHeaders ? column.trim().toLowerCase() : column.trim();
                if (key) {
                  for (var j = 0; j < keys.length; j++) {
                    if (keys[j].indexOf(key) > -1) {
                      foundCounter++;
                    }
                  }
                  key += foundCounter + 1;
                  keys.push(key);
                }
                key && (record[currentTileValue][breakValue][key] = row[index].trim());
              })
            })();
          }
          else {
            (function() {
              record[currentTileValue][breakValue] = {};
              for (var i = currentBreak + 2; i < nextBreak; i++) {
                var header = rows[currentBreak + 1].slice(1);
                var row = rows[i];
                var copyId = row[0].toLowerCase();
                var foundCounter = 0;
                

                for (var j = 0; j < keys.length; j++) {
                  if (keys[j].indexOf(copyId) > -1) {
                    foundCounter++;
                  }
                }
                copyId += foundCounter + 1;
                keys.push(copyId);                
                record[currentTileValue][breakValue][copyId] = {};
                header.forEach(function (column, index) {
                  if (index) {
                    var key = lowerCaseHeaders ? column.trim().toLowerCase() : column.trim();
                    key && (record[currentTileValue][breakValue][copyId][key] = row[index].trim());
                  }
                })                
              }
            })();
          }
        }
      }
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      if (output !== null) {
        var stream = fs.createWriteStream(output, { flags: 'w' });
        stream.write(JSON.stringify(record));
        callback(null, record);
      } else {
        callback(null, record);
      }

    })
    .on('error', function (error) {
      console.error(error.message);
    });
}
