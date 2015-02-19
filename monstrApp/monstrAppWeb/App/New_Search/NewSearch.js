(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            //will add a new column to search in, initially will have nothing to search for ~Thea
            //$('#get-new-column').click(addNewColumn);
            //adds new values to the most resent column added to current search ~Thea
            //$('#ch-param').click(addNewParam);
            checkLocal();
            $('#select_column').click(function () {
                selectColumn();
            });

            $('#add_term').click(function () {
                addTerm($('#get-param').val());
            });
            $('#add_search').click(function () {
                addNewSearch();
            });
            $('#col_submit').click(function () {
                    colSubmit($('#get_col').val());
            })
        });
    };

    var newColIndex;

    function colSubmit(col) {
        newColIndex = convertLettersToNumbers($('#get_col').val());
        $("#get_col").remove();
        $("#col_submit").remove();
        $("#select_col").remove();

        var $label = $('<label/>', {
            type: 'label',
            id: col,
            text: "You selected column: " + col.toUpperCase(),
        });
        $label.appendTo("#col_area");
    }

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
        var searchCol = [newColIndex, curTerms]; // [X,[1,2,3]]
        currentSearch.push(searchCol);
        localStorage["currentSearch"] = JSON.stringify(currentSearch);
    }

    //Function that checks to make sure the user entered in an acceptable column. Dan
    function checkColumn(column) {
        if (column = "") {
            return false;
        }
        var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        for (var i = 0; i < column.length; i++) {
            if (column[i] in alphabet) {
                continue;
            }
            else
                return false;
        }
        return true;
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