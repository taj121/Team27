(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#output_sheet_submit').click(function () {
                if (submitSheet($('#output_sheet_value').val())) {
                    resultsToNewSheetIterate();
                }
            });
        });
    };

    var searchResults = getResults();
    var tempInnerResult = [];
    var tempResult = [];
    var innerResultArraySize = 0;
    var startColumnAuto = 0;
    var startRowAuto = 0;
    var bindingArea = "Sheet2!A1:";
    var arraySize = searchResults.length;
    var jIncrement = 0;
    var iIncrement = 0;
    var functionCheck = 0;
    var iterateCheck = 0;


    function resultsToNewSheetIterate() {
        for (var i = 0; i < arraySize; i++) {
            tempInnerResult = searchResults[i];
            innerResultArraySize = tempInnerResult.length;
            functionCheck = 0;
            //console.log('tempInnerResult just before loop ' + tempInnerResult);
            //console.log('functionCheck outside loop ' + functionCheck);
            for (var j = 0; j < innerResultArraySize; j++) {
                console.log('functionCheck loop after function ' + functionCheck);
                functionCheck = 0;
                while (functionCheck <= 1) {
                    //console.log('tempResult in iterate ' + tempResult);
                    //console.log('innerArraySize ' + innerResultArraySize);
                    //console.log('innerArray[i] ' + tempInnerResult[j]);
                    iIncrement = i + 1;
                    jIncrement = j + 1;
                    bindingArea = "Sheet2!A" + iIncrement + ":";
                    startRowAuto = i;
                    startColumnAuto = j;
                    bindingArea += convertNumberToLetter(innerResultArraySize) + 1;
                    //console.log("computedBindingArea " + bindingArea);
                    tempResult = tempInnerResult[j];
                    //console.log('functionCheck inside loop ' + functionCheck);
                    resultsToNewSheet();
                }
            }
        }
    }


    /*
    *   Function to display results on a new sheet / specified range within current sheet.
    *   Creates a binding with the search results and binding range passed. 
    *   'bindingArea' specifies the sheet and range, 'searchResults' the filtered results. - nigel
    */
    function resultsToNewSheet() {
        Office.context.document.bindings.addFromNamedItemAsync(bindingArea, Office.BindingType.Matrix, { id: "NewBinding" }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification("Error", "Please make sure you have created a new sheet and entered the right sheet name (caps-sensitive)");
            }
            else {
                //app.showNotification('Displaying your search results.');
                console.log('tempResult ' + tempResult);
                console.log('startColInside ' + startColumnAuto);
                console.log('startRowInside ' + startRowAuto);

                functionCheck++;
                console.log('functionCheck inside function ' + functionCheck);
                Office.select("bindings#NewBinding").setDataAsync(tempResult,
                {
                    coercionType: "matrix",
                    startColumn: startColumnAuto,
                    startRow: startRowAuto,
                }, function (asyncResult) {
                    if (asyncResult.status == "failed") {
                        //app.showNotification('Error: ' + asyncResult.error.message);
                    }
                });
            }
        })
    };


    function submitSheet(sheet) {
        if (arraySize == 0) {
            app.showNotification("Error:", "No entries satisfy the search parameters");
            return false;
        }
        else if (sheet == "" || sheet == "New spreadsheet name") {
            app.showNotification("Error:", "Please enter a sheet name");
            return false;
        }
        return true;
    }

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
                            //newData.push(deleteArrayCommas(data.value[j]));
                            newData.push(data.value[j]);
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
                    // added for items separated by commas in cell, nigel
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
     *  Convert user input target sheet to correct binding range. nigel
     */
    function convertInputToValidRange(inputSheet) {
        var bindingTarget = inputSheet + "!A1:";
        return bindingTarget;
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
    *   Function that converts 2d array to 1d. nigel
    */
    function convertTwoDimToOneDim(array) {
        var oneDimArray = [];
        oneDimArray = oneDimArray.concat.apply(oneDimArray, array);
        return oneDimArray;
    }


    /*
    *  Check if a string contains commas. nigel
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
    *   Function to delete extra commas from 
    *   previous empty cells in row. nigel
    */
    function deleteArrayCommas(array) {
        var newArray = [];
        for (var i = 0; i < array.length; i++) {
            console.log('array[i] ' + array[i]);
            if (array[i] != "") {
                //console.log('array[i] ' + array[i]);
                newArray.push(array[i]);
            }
        }
        console.log('newArray ' + newArray);
        return newArray;
    }


    /*   
    *  Check for element in array. nigel
    */
    function containsElementSeparatedByCommas(array, element) {
        for (var i = 0; i < array.length; i++) {
            if (array[i].toUpperCase().replace(/\s+/g, '') == element.toUpperCase().replace(/\s+/g, '')) {
                //console.log('match');
                return true;
            }
        }
        return false;
    }


})();

