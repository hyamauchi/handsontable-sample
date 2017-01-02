(function(){
  "use strict";
  angular.module('myApp', ["ui.bootstrap"])
    .controller('MainController', ['$http', function ($http) {

      var that = this;

      that.data = [];

      that.init = function () {

        var dataObject = [
          {col1: 1, col2: 'あああ', col3: 12345},
          {col1: 2, col2: 'いいい', col3: 6789012345}
        ];

        that.data = dataObject;

        var hotElement = document.querySelector('#hot');
        var hotSettings = {
          data: dataObject,
          columns: [
            {
              data: 'col1',
              type: 'numeric',
              width: 40
            },
            {
              data: 'col2',
              type: 'text'
            },
            {
              data: 'col3',
              type: 'numeric',
              format: '0,000'
            }
          ],
          // stretchH: 'all',
          width: 806,
          autoWrapRow: true,
          height: 441,
          maxRows: 22,
          rowHeaders: true,
          colHeaders: [
            '列1', '列2', '列3'
          ]
        };

        var hot = new Handsontable(hotElement, hotSettings);

      };

      function downloadCSV() {
        function datenum(v, date1904) {
          if (date1904) v += 1462;
          var epoch = Date.parse(v);
          return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
        }

        function sheet_from_array_of_arrays(data, opts) {
          var ws = {};
          var range = {s: {c: 100, r: 100}, e: {c: 0, r: 0}};
          for (var R = 0; R != data.length; ++R) {
            for (var C = 0; C < Object.keys(data[R]).length; ++C) {
              if (range.s.r > R) range.s.r = R;
              if (range.s.c > C) range.s.c = C;
              if (range.e.r < R) range.e.r = R;
              if (range.e.c < C) range.e.c = C;
              var cell = {v: Object.values(data[R])[C]};
              if (cell.v == null) continue;
              var cell_ref = XLSX.utils.encode_cell({c: C, r: R});

              if (typeof cell.v === 'number') cell.t = 'n';
              else if (typeof cell.v === 'boolean') cell.t = 'b';
              else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = datenum(cell.v);
              }
              else cell.t = 's';

              ws[cell_ref] = cell;
            }
          }
          if (range.s.c < 100) ws['!ref'] = XLSX.utils.encode_range(range);
          return ws;
        }

        function Workbook() {
          if (!(this instanceof Workbook)) return new Workbook();
          this.SheetNames = [];
          this.Sheets = {};
        }

        var key = XLSX.utils.encode_cell({c: 0, r: 0});
        var ws = sheet_from_array_of_arrays(that.data);

        var workbook = new Workbook();
        workbook.SheetNames.push("シート1");
        workbook.Sheets["シート1"] = ws;

        var wbout = XLSX.write(workbook, {
          bookType: 'xlsx',
          bookSST: true,
          type: 'binary'
        });

        function s2ab(s) {
          var buf = new ArrayBuffer(s.length);
          var view = new Uint8Array(buf);
          for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
          return buf;
        }

        saveAs(new Blob([s2ab(wbout)], {type: ""}), "report.xlsx");
      }

      that.downloadCSV = function () {
        downloadCSV();
      };

    }]);
})();
