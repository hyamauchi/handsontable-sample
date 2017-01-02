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

    }]);
})();
