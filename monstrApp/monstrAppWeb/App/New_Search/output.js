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
    
    var sheet;
    var searchResults = JSON.parse(localStorage["data"]);
    for (var i = 0; i < searchResults.length; i++) {
        Debug.writeln('[' + searchResults[i] + ']');
    }

    function submitSheet(inputSheet) {
        if (searchResults.length == 0) {
            app.showNotification("Error:", "No entries satisfy the search parameters");
            return false;
        }
        else if (sheet == "" || sheet == "New spreadsheet name") {
            app.showNotification("Error:", "Please enter a sheet name");
            return false;
        }
        sheet = inputSheet;
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
        var searchResults = JSON.parse(localStorage["data"]);
        var Idx = 0;
        var letter;
        var bindingArea;
        var data;
        Debug.writeln("LENGTH: " + searchResults.length);

        for (var i = 0; i < searchResults.length; i++) {
            Idx++;
            letter = convertSizetoLetter(searchResults[i].length);
            bindingArea = sheet + "!A" + Idx + ":" + letter + Idx;
            data = [searchResults[i]];
            Debug.writeln(searchResults[i])
            Debug.writeln("[" + searchResults[i] + "] " + i);
            outputResult(bindingArea, data);

        }


        
    }


    function outputResult(bindingArea, data) {

       var myTable = new Office.TableData();

       myTable.rows = searchResults;
       
        var type = myTable.rows.length;
        var letter = convertSizetoLetter(searchResults[0].length);
        var namedRange = sheet + "!A1:" + letter + searchResults.length;
        var id = "Rec" + type;
        var bnding = "bindings#" + id;
        Office.context.document.bindings.addFromNamedItemAsync(namedRange, "table", { id: id },
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    app.showNotification('Error: ' , asyncResult.error.message);
                }
            });
        Office.select(bnding).setDataAsync(myTable, { coercionType: "table" },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Error: ' , asyncResult.error.message);
                } else {
                    app.showNotification("Success!", "Displaying your data");
                    window.location = "/App/New_Search/Save_Search.html";
                }
            });

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
