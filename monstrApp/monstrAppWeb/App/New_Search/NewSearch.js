(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            //will add a new column to search in, initially will have nothing to search for ~Thea
            //$('#get-new-column').click(addNewColumn);
            //adds new values to the most resent column added to current search ~Thea
            //$('#add-search-param').click(addNewParam);
            checkLocal();
            $('#add-search').click(function () {
                addNewSearch();
            });
        });
    };

    //stores the current search array of form column-dataToSearchFor column being index 0 of internal array ~Thea
    var currentSearch = [];

    function addNewSearch() {
        var newColIndex = convertLettersToNumbers($('#get-col').val());
        var param = $('#get-param').val().replace(/\s+/g, ''); //remove all whitespaces
        var params = param.split(",");
        var searchCol = [newColIndex, params]; // [X,[1,2,3]]
        currentSearch.push(searchCol);
        localStorage["currentSearch"] = JSON.stringify(currentSearch);
    }

    function checkLocal() {
        currentSearch = localStorage.getItem("currentSearch");
        if (currentSearch || currentSearch === "") {
            currentSearch = JSON.parse(localStorage["currentSearch"]);
        }
        else {
            currentSearch = [];
        }
    }

    /*
    Function to convert letters to numbers, nigel
    */
    function convertLettersToNumbers(inputLetters) {
        var valueToReturn = 0;
        var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

        for (var i = 0, j = inputLetters.length - 1; i < inputLetters.length; j--, i++) {
            valueToReturn += (alphabet.indexOf(inputLetters[i].toUpperCase()) + 1) * Math.pow(alphabet.length, j);
        }
        return valueToReturn - 1;
    }

    /*
    Adds a new column to search in to the array of current search. 
    Defaults with nothing to search for. 
    ~Thea
    */
    function addNewColumn() {
        var newColIndex = convertLettersToNumbers($('#get-col').val());
        currentSearch.push([newColIndex]);
    }

    /*
    Fucntion to add new search parameters to the most recent column added to the current search
    ~Thea
    */
    function addNewParam() {
        var column = currentSearch.pop();
        var param = $('#get-param').val();
        column.push(param);
        currentSearch.push(column);
    }
})();