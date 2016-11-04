var fs = require('fs');
var path = require('canonical-path');
var _ = require('lodash');
var xl=require('excel4node');



// Custom reporter
var Reporter = function(options) {
  var _defaultOutputFile = path.resolve(process.cwd(), './_test-output', 'protractor-results.xlsx');
  options.outputFile = options.outputFile || _defaultOutputFile;

  var wb=new xl.Workbook();
  var ws = wb.addWorksheet('Sheet 1');
  var style = wb.createStyle({
    font: {
        color: '#6B0909',
        size: 14
    }
 });
  initOutputFile(options.outputFile);
  options.appDir = options.appDir ||  './';
  var _root = { appDir: options.appDir, suites: [] };
  log('AppDir: ' + options.appDir, +1);
  var _currentSuite;

  this.suiteStarted = function(suite) {
    _currentSuite = { description: suite.description, status: null, specs: [] };
    _root.suites.push(_currentSuite);
    log('Suite: ' + suite.description, +1);
  };

  this.suiteDone = function(suite) {
    var statuses = _currentSuite.specs.map(function(spec) {
      return spec.status;
    });
    statuses = _.uniq(statuses);
    var status = statuses.indexOf('failed') >= 0 ? 'failed' : statuses.join(', ');
    _currentSuite.status = status;
    log('Suite ' + _currentSuite.status + ': ' + suite.description, -1);
  };

  this.specStarted = function(spec) {

  };

  this.specDone = function(spec) {
    var currentSpec = {
      description: spec.description,
      status: spec.status
    };
    if (spec.failedExpectations.length > 0) {
      currentSpec.failedExpectations = spec.failedExpectations;
    }

    _currentSuite.specs.push(currentSpec);
    log(spec.status + ' - ' + spec.description);
  };

  this.jasmineDone = function() {
    outputFile = options.outputFile;
    var output = formatOutput(_root);
    wb.write(outputFile);
  };

  function ensureDirectoryExistence(filePath) {
    var dirname = path.dirname(filePath);
    if (directoryExists(dirname)) {
      return true;
    }
    ensureDirectoryExistence(dirname);
    fs.mkdirSync(dirname);
  }

  function directoryExists(path) {
    try {
      return fs.statSync(path).isDirectory();
    }
    catch (err) {
      return false;
    }
  }

  function initOutputFile(outputFile) {
    ensureDirectoryExistence(outputFile);
    var header = "Protractor results for: " + (new Date()).toLocaleString() + "\n\n";

    ws.cell(1,2,1,7,true).string("Protractor results for: " + (new Date()).toLocaleString()).style(style);
    wb.write(outputFile);
  }

  // for output file output
  function formatOutput(output) {
    var indent = '  ';
    var pad = '  ';
    var results = [];
    var i=3,
        j=1;
    results.push('AppDir:' + output.appDir);
    ws.cell(2,1).string('AppDir:'+ output.appDir).style(style);;
    output.suites.forEach(function(suite) {
      ws.cell(i,1).string(pad + 'Suite: ' + suite.description + ' -- ' + suite.status).style(style);
      results.push(pad + 'Suite: ' + suite.description + ' -- ' + suite.status);
      pad+=indent;
      suite.specs.forEach(function(spec) {
        results.push(pad + spec.status + ' - ' + spec.description);
        ws.cell(i+j,1).string(pad + spec.status + ' - ' + spec.description).style(style);
        if (spec.failedExpectations) {
          pad+=indent;
          spec.failedExpectations.forEach(function (fe) {
            ws.cell(i+j,j+1).string(pad + 'message: ' + fe.message).style(style);
            results.push(pad + 'message: ' + fe.message);
          });
          pad=pad.substr(2);
        }
        j++;
      });
      pad = pad.substr(2);
      results.push('');
      i=i+j;
    });
    results.push('');
    return results.join('\n');
  }

  // for console output
  var _pad;
  function log(str, indent) {
    _pad = _pad || '';
    if (indent == -1) {
      _pad = _pad.substr(2);
    }
    console.log(_pad + str);
    if (indent == 1) {
      _pad = _pad + '  ';
    }
  }
};

module.exports = Reporter;
