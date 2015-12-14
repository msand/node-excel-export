
'use strict';

var excel = require('./lib/excel');
module.exports = {
  buildExport: function(params) {
    if( ! (params instanceof Array)) throw 'buildExport expects an array';

    var sheets = [];
    params.forEach(function(sheet, index) {
      var specification = sheet.specification;
      var dataset = sheet.data;
      var sheet_name = sheet.name || 'Sheet' + index+1;
      var data = [];
      var config = {
        cols: []
      };

      if( ! specification || ! dataset) throw 'missing specification or dataset.';

      if(sheet.heading) {
        sheet.heading.forEach(function(row) {
          data.push(row);
        });
      }

      //build the header row
      var header = [];
      for (var col in specification) {
        header.push({
          value: specification[col].displayName,
          style: (specification[col].headerStyle) ? specification[col].headerStyle : undefined
        });

        if(specification[col].width) {
          if(Number.isInteger(specification[col].width)) config.cols.push({wpx: specification[col].width});
          else if(Number.isInteger(parseInt(specification[col].width))) config.cols.push({wch: specification[col].width});
          else throw 'Provide column width as a number';
        } else {
          config.cols.push({});
        }

      }
      data.push(header); //Inject the header at 0

      dataset.forEach(function(record) {
        var row = [];
        for (var col in specification) {
          var cell_value = record[col];

          if(specification[col].cellFormat && typeof specification[col].cellFormat == 'function') {
            cell_value = specification[col].cellFormat(cell_value);
          }

          if(specification[col].cellStyle) {
            if (typeof specification[col].cellStyle == 'function') {
              cell_value = {
                value: cell_value,
                style: specification[col].cellStyle(cell_value, record)
              };
            } else {
              cell_value = {
                value: cell_value,
                style: specification[col].cellStyle
              };
            }
          }
          row.push(cell_value); // Push new cell to the row
        }
        data.push(row); // Push new row to the sheet
      });

      sheets.push({
        name: sheet_name,
        data: data,
        config: config
      });

    });

    return excel.build(sheets);

  }
};
