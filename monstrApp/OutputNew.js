(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#results-to-new-sheet').click(resultsToNewSheet);
        });
    };

    var searchResults = getResults();
    searchResults = convertTwoDimToOneDim(searchResults);
    var arraySize = searchResults.length;
    searchResults = convertOneDimToTwoDim(searchResults);
    var bindingArea = "Sheet2!A1:A" + arraySize; // storing in just one column 'A', on 'Sheet2'


    /*
    *   Function to display results on a new sheet / specified range within current sheet.
    *   Creates a binding with the search results and binding range passed. 
    *   'bindingArea' specifies the sheet and range, 'searchResults' the filtered results. - nigel
    */
    function resultsToNewSheet(elementId) {
        Office.context.document.bindings.addFromNamedItemAsync(bindingArea, Office.BindingType.Matrix, { id: "NewBinding" }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification("Error : create a new sheet.");
            }
            else {
                app.showNotification('Displaying your search results.');
                Office.select("bindings#NewBinding").setDataAsync(searchResults,
                    {
                        coercionType: "matrix"
                    }, function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            app.showNotification('Error: ' + asyncResult.error.message);
                        }
                    });
            }
        })
    };


    /*
    function gets rows to return. 
    ~Thea
    */
    function getResults() {
        var data = JSON.parse(localStorage["userDataSelection"]);
        console.log('data getResults() ' + data);
        var temp = localStorage["currentSearch"];
        var searchFor = JSON.parse(temp);
        var newData = [];

        for (var i = 0; i < searchFor.length; i++) {
            var column = searchFor[i][0];
            var toFind = searchFor[i][1];
            console.log('toFind ' + toFind);
            var rowsToGet = findIndex(toFind, column, data);
            for (var k = 0; k < rowsToGet.length; k++) {
                console.log('data.value before ' + data.value);
                if (data.value != undefined) {
                    for (var j = 0; j < data.value.length; j++) {
                        if (j == rowsToGet[k]) {
                            //newData.push(data.value[j]);
                            // remove commas when adding terms nigel
                            newData.push(deleteArrayCommas(data.value[j]));
                        }
                    }
                }
                else {
                    for (var j = 0; j < data.length; j++) {
                        if (j == rowsToGet[k]) {
                            newData.push(data[j]);
                        }
                    }
                }
            }

            data = newData;
        }
        return data;
    }


    /*
    Function to find the index of an array of serch parameters passed in. In a specific column
    ~Thea
    */
    function findIndex(toFind, col, result) {
        var foundAt = [];
        if (result.value != undefined) {
            for (var i = 0; i < result.value.length; i++) {
                for (var j = 0; j < toFind.length; j++) {
                    var currentCheck = result.value[i][col];
                    if (typeof currentCheck != "string") {
                        currentCheck = String(currentCheck);
                    }
                    // added this for when item contained in comma 
                    // separated cell with other words/numbers. nigel
                    if (containsCommas(currentCheck)) {
                        var array = currentCheck.split(',');
                        if (containsElementSeparatedByCommas(array, toFind[j])) {
                            foundAt.push(i);
                        }
                    }
                    else if (toFind[j].toUpperCase().replace(/\s+/g, '') == currentCheck.toUpperCase().replace(/\s+/g, '')) {
                        foundAt.push(i);
                    }
                }
            }
        }
        else {
            for (var i = 0; i < result.length; i++) {
                for (var j = 0; j < toFind.length; j++) {
                    if (toFind[j] == result[i][col]) {
                        foundAt.push(i);
                    }
                }
            }
        }
        return foundAt;
    }

    /*
     *  Convert user input target sheet to correct binding range.
     *  For now this is still going down column A, needs to be linked to bindingArea variable. nigel
     */
    function convertInputToValidRange(inputSheet) {
        var bindingTarget = inputSheet + "!A1:A";
        return bindingTarget;
    }

    /*
    *   Convert numbers to column letters - nigel
    */
    function convertNumberToLetter(number) {
        var asciiCodeA = 65; // capital A
        var asciiCodeZ = 90;
        var alphabetLength = 26;
        var columnName = "";
        for (number; number >= 0; number = Math.floor(number / alphabetLength) - 1) {
            columnName += String.fromCharCode(number % alphabetLength + asciiCodeA);
        }
        return columnName;
    }

    /*   
     *  Check for element in array. nigel
     */
    function containsElementSeparatedByCommas(array, element) {
        for (var i = 0; i < array.length; i++) {
            if (array[i].toUpperCase().replace(/\s+/g, '') == element.toUpperCase().replace(/\s+/g, '')) {
                console.log('true match');
                return true;
            }
        }
        return false;
    }

    /*
     *   Check if a string contains commas. nigel
     */
    function containsCommas(string) {
        for (var i = 0; i < string.length; i++) {
            if (string[i] == ',') {
                return true;
            }
        }
        return false;
    }


    /*
    *   Converts 1d to 2d
    */
    function convertOneDimToTwoDim(array) {
        var twoDimArray = [];
        while (array.length) {
            twoDimArray.push(array.splice(0, 1));
        }
        return twoDimArray;
    }

    /*
    *   convert 2d array to 1d. nigel
    */
    function convertTwoDimToOneDim(array) {
        var oneDimArray = [];
        oneDimArray = oneDimArray.concat.apply(oneDimArray, array);
        return oneDimArray;
    }


    /*
    *   Delete extra commas from 
    *   previous empty cells in row. nigel
    */
    function deleteArrayCommas(array) {
        var newArray = [];
        for (var i = 0; i < array.length; i++) {
            console.log('array[i] ' + array[i]);
            if (array[i] != "") {
                //console.log('array[i][j] ' + array[i]);
                newArray.push(array[i]);
            }
        }
        console.log('newArray ' + newArray);
        return newArray;
    }

})();