(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded

    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            getNamesSaved();

            $('#view_selected').click(function () {
                loadSelectedSearch();
                var toPrint = searchToString();
                $("#return_search_view").text(toPrint);

            })

            $('#run_selected').click(function () {
                loadSelectedSearch();
                getResults();
            });
        });
    };

    /*
    load names of saved searches
    ~Thea
    */
    function getNamesSaved() {
        var names = Office.context.document.settings.get('search_names');
        if (names || names === "") {
            names=JSON.parse(names);
            var select = document.getElementById('saved_searches');
            for (var i = 0; i < names.length; i++) {
                var opt = document.createElement('option');
                opt.innerHTML = names[i];
                opt.value = names[i];
                select.appendChild(opt);
            }

        }
        else {
            app.showNotification("Error:","You must save a search before you can use this menu")
        }
    }

    /*
    function that will set current search to be the selected saved search
    ~Thea
    */
    function loadSelectedSearch() {
        var searchName = $('#saved_searches option:selected').val();
        var search = Office.context.document.settings.get(searchName);
        localStorage["currentSearch"] = search;
        
    }

    function getResults() {
        var data = JSON.parse(localStorage["userDataSelection"]);
        //console.log('data getResults() ' + data);
        var temp = localStorage["currentSearch"];
        var searchFor = JSON.parse(temp);
        var newData = [];

        for (var i = 0; i < searchFor.length; i++) {
            var newData = [];
            var column = searchFor[i][0];
            var toFind = searchFor[i][1];
            //console.log('toFind ' + toFind);
            var rowsToGet = findIndex(toFind, column, data);
            for (var k = 0; k < rowsToGet.length; k++) {
                //console.log('data.value before ' + data.value);
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
        if (data.length == 0) {
            window.location = "/App/New_Search/Oops.html"
        }
        else {
            localStorage["data"] = JSON.stringify(data);
            window.location = "/App/New_Search/Bindings.html"

        }
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
        return newArray;
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
    function searchToString() {
        var details = "";
        var curr = localStorage["currentSearch"];
        var currentSearch = JSON.parse(curr);
        for (var i = 0; i < currentSearch.length; i++) {
            details += "\n-Search in column " + currentSearch[i][0] + " for:\n"
            for (var j = 0; j < currentSearch[i][1].length; j++) {
                details += "  -" + currentSearch[i][1][j] + "\n";

            }

        }
        //app.showNotification(details);
        $('#view').show();
        return details;
    }

})();