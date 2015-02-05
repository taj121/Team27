/// <reference path="../App.js" />
//nigel fdgsfg
//danny v

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () 
           { app.initialize();

            $('#get--selected-data').click(getDataFromSelection);
        });
    };

    function getText() {
        return $('#get').val();
    }

    /*
    Function that will return a specific row of an array.
    rowNum = the number of the row you wish to reutrn 
    result = the array you wish to return the row from
    ~Thea
    */
    function returnRow(rowNum, result) {
        return result.value[rowNum];
    }

    /*
    Function that will find a string in an array (selection of cells passed in) and will print the index's 
    that it appears at. 
    Could be altered easily to return array of index's (at the moment in string form) that could be used to reference 
    another array
    toFind is the item to be searched for in the array
    result is the array (2D object returned from getting selection)
    ~Thea
    */
    function findIndex(toFind, result) {
        var indexToFind = [];
        for (var i = 0; i < result.value.length; i++) {
            for (var j = 0; j < result.value[i].length; j++) {
                if (toFind == result.value[i][j]) {
                    indexToFind.push('[' + i + ']' + '[' + j + ']');
                }
            }
        }
        if (indexToFind.length > 0) {
            app.showNotification('The data item ' + toFind + '  is in the Data at ', indexToFind.toString() + '');
        }
        else {
            app.showNotification('The data item ' + toFind + '  is not in the Data', '');
        }
    }


    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    //get string value from user
                    var toFind = getText();
                    //check if it is in the selection
                    findIndex(toFind, result);                    
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
})();