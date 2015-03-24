/// <reference path="../App.js" />

(function () {
    "use strict";
    
    // The initialize function must be run each time a new page is loaded
    //Dan note: This is causing an error for me all of a sudden..
    //only happens on some program executions.. not sure if this is happening for others...
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            //this is the cookie stuff but it gets really annoying so I'll put it back later. Dan
            //if (!(document.cookie)) {
            //    document.cookie = "FirstTime=true";
            //    window.location = "/App/Tutorial/Tutorial_Page.html"
            //    var myWindow = window.open("", "MsgWindow", "width=500, height=500");
            //    myWindow.document.write("<pre>" + "Welcome first time user!" + "</pre>");
            //}
            bindButton();
            checkIfSavedSearches();
        });        
    };

    function bindButton() {
        $('#new-search').click(function () { newSearch(); });
        $('#saved-search').click(function () { savedSearch(); });
    }


    /*
    Updates global variable to user selection when called and loads new search page if data selected correctly
    ~Thea
    */
    function newSearch() {
        getSelectedData("/App/New_Search/Range_Type_Specification/SelectSearchRange.html");
    }

    /*
    Updates global variable to user selection when called and changes to saved search page if data selected correctly
    ~Thea
    */
    function savedSearch() {
        getSelectedData("/App/Saved_Search/Saved_Search_Menu.html");
    }

    /*
    updates the global variable and sets the new page to be the corerct page. Will change if more than one row is inputed
    ~Thea
    */
    function getSelectedData(url) {
        localStorage.clear();
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Matrix,
            {   
                valueFormat: Office.ValueFormat.Formatted,
            },
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    if (result.value.length == 1) {
                        app.hideAllNotification();
                        app.showNotification('Error:', 'Only one row has being selected!');
                    } else {

                        app.hideAllNotification();
                        for (var i = 0; i < result.value.length; i++) {
                            for (var j = 0; j < result.value[i].length; j++) {
                                if (result.value[i][j] == "") {
                                    result.value[i][j] = " "
                                }
                            }
                        }
                        localStorage["userDataSelection"] = JSON.stringify(result);
                        
                        window.location = url;
                        //json.parse(localstorage["userDataSelection"]); how to get values back DO NOT UNCOMMENT OR DELETE!!!!!
                    }
                } else {
                    app.hideAllNotification();
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function getNamesSaved() {
        var names = Office.context.document.settings.get('search_names');
        if (names || names === "") {
            names = JSON.parse(names);
            var select = document.getElementById('saved_searches');
            for (var i = 0; i < names.length; i++) {
                var opt = document.createElement('option');
                opt.innerHTML = names[i];
                opt.value = names[i];
                select.appendChild(opt);
            }

        }
        else {
            app.showNotification("Error:", "You must save a search before you can use this menu")
        }
    }




    
    //check if saved searches
    //~Thea
    function checkIfSavedSearches() {
        var names = Office.context.document.settings.get('search_names');
        if (names || names === "") {
        }
        else {
            //$("#saved-search").attr("disabled", true);
            $("#saved-search").attr("hidden",true)
        }
    }
})();