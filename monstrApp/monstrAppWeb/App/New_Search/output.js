(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            
            $('#select_column').click(function () {
                selectColumn();
            });
            
            $('#output_sheet_submit').click(function () {
                if (submitSheet($('#output_sheet_value').val())) {
                    outputResult();
                }
            });
        });
    };
    
    var sheet;
    var searchResults = JSON.parse(localStorage["data"]);
    

    function submitSheet(inputSheet) {
        if (searchResults.length == 0) {
            app.showNotification("Error:", "No entries satisfy the search parameters");
            return false;
        }
        else if (sheet == "" || sheet == "New spreadsheet name" || sheet==null) {
            app.showNotification("Error:", "Please enter the exact sheet name, eg 'Sheet2'");
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
    


    function outputResult() {

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



  


})();
