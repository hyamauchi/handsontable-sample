(function(){
  "use strict";
  document.addEventListener("DOMContentLoaded", function(event) {
    var dataObject = [
      {col1: 1, col2: 'あああ', col3: 12345},
      {col1: 2, col2: 'いいい', col3: 6789012345}
    ];

    var hotElement = document.getElementById('hot');
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

    var downloadCSV = function() {
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
      var ws = sheet_from_array_of_arrays(dataObject);

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
    };

    document.getElementById('downloadCSV').addEventListener('click', downloadCSV, false);


    var X = XLSX;

    function to_csv(workbook) {
      var result = [];
      workbook.SheetNames.forEach(function(sheetName) {
        var csv = X.utils.sheet_to_csv(workbook.Sheets[sheetName]);
        if(csv.length > 0){
          result.push("SHEET: " + sheetName);
          result.push("");
          result.push(csv);
        }
      });
      return result.join("\n");
    }

    function to_json(workbook) {
      var result = {};
      workbook.SheetNames.forEach(function(sheetName) {
        var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        if(roa.length > 0){
          result[sheetName] = roa;
        }
      });
      return result;
    }

    function importToHot(workbook) {
      var rows = [];
      var columns = [];
      var worksheet = workbook.Sheets[workbook.SheetNames[0]];
      for (var z in worksheet) {
        /* all keys that do not begin with "!" correspond to cell addresses */
        if(z[0] === '!') continue;
        console.log(z + "=" + JSON.stringify(worksheet[z].v));

        var cellAddr = z.match( /^([A-Z]+)(\d+)$/ );
        var idx = cellAddr[2] - 1;
        var row = rows[idx];
        if (!row) row = {};
        row[cellAddr[1]] = worksheet[z].w;
        rows[idx] = row;

        var newCol = true;
        for (var i = 0; i < columns.length; i++) {
          if (columns[i].data == cellAddr[1]) {
            newCol = false;
            break;
          }
        }
        if (newCol) columns.push({
          data: cellAddr[1],
          type: worksheet[z].t == 'n' ? 'numeric' : 'text'
        });
      }

      columns.sort(function(a, b){
        if (a.data.length < b.data.length) return -1;
        if (a.data.length > b.data.length) return 1;
        if (a.data < b.data) return -1;
        if (a.data > b.data) return 1;
        return 0;
      });

      var colHeaders = [];
      columns.forEach(function(col){
        colHeaders.push(col.data);
      });

      var elm = document.getElementById('drop');

      dataObject = rows;
      var hotSettings = {
        data: dataObject,
        columns: columns,
        // stretchH: 'all',
        // preventOverflow: 'horizontal',
        width: elm.clientWidth,
        autoWrapRow: true,
        height: 441,
        maxRows: rows.length,
        maxColumns: columns.length,
        rowHeaders: true,
        colHeaders: colHeaders
      };
      hot.destroy();
      hot = new Handsontable(hotElement, hotSettings);
    }

    function process_wb(wb) {
      // var output = to_csv(wb);
      // var output = to_json(wb);
      // console.dir(output);
      importToHot(wb);
    }

    var drop = document.getElementById('drop');
    function handleDrop(e) {
      e.stopPropagation();
      e.preventDefault();
      var files = e.dataTransfer.files;
      var f = files[0];
      importXls(f);
    }

    function importXls(f) {
      var reader = new FileReader();
      var name = f.name;
      reader.onload = function (e) {
        var data = e.target.result;
        var wb = X.read(data, {type: 'binary'});
        process_wb(wb);
      };
      reader.readAsBinaryString(f);
    }

    function handleDragover(e) {
      e.stopPropagation();
      e.preventDefault();
      e.dataTransfer.dropEffect = 'copy';
    }

    if (drop.addEventListener) {
      drop.addEventListener('dragenter', handleDragover, false);
      drop.addEventListener('dragover', handleDragover, false);
      drop.addEventListener('drop', handleDrop, false);
    }

    var xlf = document.getElementById('xlf');
    function handleFile(e) {
      var files = e.target.files;
      var f = files[0];
      importXls(f);
    }

    if (xlf.addEventListener) xlf.addEventListener('change', handleFile, false);

  });

})();
