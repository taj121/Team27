/// <reference path="../App.js" />
//nigel fdgsfg
//danny v

(function () {
    "use strict";
    
    // The initialize function must be run each time a new page is loaded
    //Dan note: This is causing an error for me all of a sudden..
    //only happens on some program executions.. not sure if this is happening for others...
    Office.initialize = function (reason) {
        $(document).ready(function () 
           { app.initialize();

           $('#get--selected-data').click(getDataFromSelection);
           $('#get-range-selection').click(selectRange);
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
    Function that will return a set rows of an array.
    Return type is an array of rows
    rowNums = the numbers of the rows you wish to reutrn 
    result = the array you wish to return the row from
    ~Thea
    */
    function returnRows(rowNums, result) {
        var rows = [];
        for (var i = 0; i < rowNums.length; i++) {
            rows[i] = returnRow[i];
        }
        return rows;
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

    // Function to specify a range, needs to be modified
    function selectRange() {
        Office.context.document.bindings.addFromPromptAsync(Office.BindingType.Matrix,
            { id: "MatrixBinding" },
            function (asyncResult) {
                if (Office.AsyncResultStatus.Succeeded) {
                    app.showNotification("Added range: " + asyncResult.value.type
                        + " and id: " + asyncResult.value.id);
                } else {
                    app.showNotification('Error:', asyncResult.error.message);
                }
            }
        );
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