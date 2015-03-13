(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            
            $('#select_column').click(function () {
                selectColumn();
            });
            $('#results-to-new-sheet').click(function () {
                resultsToNewSheet();
            });
            $('#output_sheet_submit').click(function () {
                if (submitSheet($('#output_sheet_value').val())) {
                    resultsToNewSheet();
                }
            });
        });
    };
    var searchResults = JSON.parse(localStorage["data"]);
    var arraySize = searchResults.length;
    var bindingArea; // storing in just one column 'A', on 'Sheet2'
    var letter = convertSizetoLetter(searchResults[0].length);


    function submitSheet(sheet) {
        if (arraySize == 0) {
            app.showNotification("Error:", "No entries satisfy the search parameters");
            return false;
        }
        else if (sheet == "" || sheet == "New spreadsheet name") {
            app.showNotification("Error:", "Please enter a sheet name");
            return false;
        }
        bindingArea = sheet + "!A1:" + letter + arraySize
        return true;
    }

    function convertSizetoLetter(idx) {
        var sb = "";
        var num = idx - 1;
        while (num >=  0) {
            var numChar = (num % 26)  + 65;
            sb += String.fromCharCode(numChar);
            num = (num  / 26) - 1;
        }
        return sb.split("").reverse().join("");
    }

    /*
    *   Function to display results on a new sheet / specified range within current sheet.
    *   Creates a binding with the search results and binding range passed. 
    *   'bindingArea' specifies the sheet and range, 'searchResults' the filtered results. - nigel
    */
    function resultsToNewSheet(elementId) {
        Office.context.document.bindings.addFromNamedItemAsync(bindingArea, Office.BindingType.Matrix, { id: "NewBinding" }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                app.showNotification("Error:", "Please make sure you have: 1) Created a new sheet and 2) Entered the right sheet name (caps-sensitive)");
            }
            else {
                app.hideAllNotification();
                app.showNotification("Success!", "Displaying your search results");
                Office.select("bindings#NewBinding").setDataAsync(searchResults,
                    {
                        coercionType: "matrix"
                    }, function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            app.hideAllNotification();
                            app.showNotification('Error: ', asyncResult.error.message);
                        }
                    });
                window.location = "/App/New_Search/Save_Search.html"
               
            }
        })
    };



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

})();
