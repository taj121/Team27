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

           //$('#get--selected-data').click(getDataFromSelection);
           $('#new-search').click(newSearch);
           $('#saved-search').click(savedSearch);
           //$('#get-range-selection').click(selectRange); commented since button currently does not exist ~Thea
           //$('#add_data').click(test);
        });
    };
    function getText() {
        return $('#get').val();
    }

    /*
    Updates global variable to user selection when called and loads new search page if data selected correctly
    ~Thea
    */
    function newSearch() {
        localStorage.clear();
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value.length == 1) {
                        app.showNotification('Error: Only one row has being selected!');
                    } else {
                        //userDataSelection = result;
                        localStorage["userDataSelection"] = JSON.stringify(result);
                        window.location = "/App/New_Search/Range_Type_Specification/SelectSearchRange.html";
                        //json.parse(localstorage["userDataSelection"]); how to get values back DO NOT UNCOMMENT OR DELETE!!!!!
                    }
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    /*
    Updates global variable to user selection when called and changes to saved search page if data selected correctly
    ~Thea
    */
    function savedSearch() {
        localStorage.clear();
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value.length == 1) {
                        app.showNotification('Error: Only one cell has being selected!');
                    } else {
                        //userDataSelection = result;
                        localStorage["userDataSelection"] = JSON.stringify(result);
                        window.location = "/App/Saved_Search/Saved_Search_Menu.html";
                        //json.parse(localstorage["userDataSelection"]); how to get values back DO NOT UNCOMMENT OR DELETE!!!!!
                    }
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


    function test() {
        app.showNotification("Added '" + $('#data_input').val() + "' to search");
        var newLabel = document.createElement("label");
        newLabel.innerHTML = $('#data_input').val();
        document.body.appendChild(newLabel);

        document.getElementById('data_input').value = "";

    }
})();