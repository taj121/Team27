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
            $('#add_term').click(function () {
                addTerm($('#get-param').val());
            })
        });
    };

    //stores the current search array of form column-dataToSearchFor column being index 0 of internal array ~Thea
    var currentSearch = [];
    var curTerms = []

    //adds terms one at a time to the search paramaters for a given column.
    //dynamically generates a button displaying each term.
    //removes the button from the page and the term from the search parameters
    //when the button is clicked. Dan
    function addTerm(term) {

        if (term != -1) {
            //terms pushed to current search as they are entered
            curTerms.push(term);
            app.showNotification(curTerms);
        }

        //create button dynamically based on term entry
        var $button = $('<button/>', {
            type: 'button',
            'class': 'dynBtn',
            id: term,
            text: term,
            value: term,
            click: function () {
                app.showNotification("Removed term '" + $(this).attr("value") + "' from search");

                var index = curTerms.indexOf($(this).attr("value"));
                if (index > -1) {
                    //terms spliced from current search as they are deleted
                    curTerms.splice(index, 1);
                }
                $(this).remove();

            }
        });
        $button.appendTo('#button_div');
        $('#get-param').val("");

    }

    function addNewSearch() {
        var newColIndex = convertLettersToNumbers($('#get-col').val());
        var searchCol = [newColIndex, curTerms]; // [X,[1,2,3]]
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