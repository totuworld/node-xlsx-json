var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

exports = module.exports = XLSX_json;

function XLSX_json (config, callback) {
  if(!config.input) {
    console.error("You miss a input file");
    process.exit(1);
  }

  var cv = new CV(config, callback);

}

function CV(config, callback) {
  var wb = this.load_xlsx(config.input)
  var ws = this.ws(config, wb);
  var csv = this.csv(ws)
  this.cvjson(csv, config.output, config.lowerCaseHeaders, config.exceptionColumn, callback)
}

CV.prototype.load_xlsx = function(input) {
  return xlsx.readFile(input);
}

CV.prototype.ws = function(config, wb) {
  var target_sheet = config.sheet;
  var exceptionPtn = config.exceptionSheet;

  if (target_sheet == null)
    target_sheet = wb.SheetNames[0];
  
  //check exeption sheet
  if (exceptionPtn != null) {
    exceptionPtn = new RegExp(exceptionPtn);
    if(exceptionPtn.test(target_sheet)==true) {
      console.error("The target sheet and destination sheet are the same.\nRemove 'exceptionSheet' from config or select another sheet.");
      process.exit(1);
    }
  }

  ws = wb.Sheets[target_sheet];
  return ws;
}

CV.prototype.csv = function(ws) {
  return csv_file = xlsx.utils.make_csv(ws)
}

CV.prototype.cvjson = function(csv, output, lowerCaseHeaders, exceptionColumn, callback) {
  var record = []
  var header = []

  cvcsv()
    .from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){

      if(exceptionColumn != null)
        exceptionColumn = new RegExp(exceptionColumn);

      if(index === 0) {
        header = row;
      }else{
        var obj = {};
        var haveKey = false;
        header.forEach(function(column, index) {
          var key = lowerCaseHeaders ? column.trim().toLowerCase() : column.trim();
          if(exceptionColumn != null) {
            if(exceptionColumn.test(key)==false) {
              if(row[index].trim() !== "") {
                haveKey = true;
                obj[key] = row[index].trim();
              }
            }
          } else {
            if(row[index].trim() !== "") {
                haveKey = true;
                obj[key] = row[index].trim();
              }
          }
        })
        if(haveKey)
          record.push(obj);
      }
    })
    .on('end', function(count){
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      if(output !== null) {
      	var stream = fs.createWriteStream(output, { flags : 'w' });
      	stream.write(JSON.stringify(record));
	      callback(null, record);
      }else {
      	callback(null, record);
      }

    })
    .on('error', function(error){
      console.error(error.message);
    });
}
