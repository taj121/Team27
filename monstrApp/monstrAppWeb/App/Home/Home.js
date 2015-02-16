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
           $('#add_data').click(test);
        });
    };

    function test() {
        app.showNotification("Added '" + $('#data_input').val() + "' to search");
        var newLabel = document.createElement("label");
        newLabel.innerHTML = $('#data_input').val();
        document.body.appendChild(newLabel);

        document.getElementById('data_input').value = "";
        
    }

    function getText() {
        return $('#get').val();
    }

    /*
    Updates global variable to user selection when called
    ~Thea
    */
    function getDataFromSelection() {
        localStorage.clear();
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    //userDataSelection = result;
                    localStorage["userDataSelection"] = JSON.stringify(result);
                    //json.parse(localstorage["userDataSelection"]); how to get values back DO NOT UNCOMMENT OR DELETE!!!!!
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    /*
    gets text from input box with this id
    ~Thea
    */
    function getText() {
        return $('#get').val();
    }


    /*
    *     Function to write data from one excel sheet to another
    *     The idea is to make the returned result from the search 
    *     the "originalData" parameter, and then write this to the new excel sheet.. 
    *     Needs frontend and testing.. nigel
    */
    function writeDataToNewSheet(targetRange, originalData) {
        // Add binding
        Office.context.document.bindings.addFromNamedItemAsync($(targetRange).val(), Office.BindingType.Matrix, { id: "exportedBinding" }, function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                writeMessage("Error: " + asyncResult.error.message);
            }
            else {
                // Write data in original to the binding in new 
                Office.select("bindings#exportedBinding").setDataAsync(originalData, { coercionType: "matrix" }, function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        app.showNotification("Error: " + asyncResult.error.message);
                    }
                    else {
                        app.showNotification("Export data from original sheet to target sheet successfully");
                    }
                });
            }
        });
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
})();