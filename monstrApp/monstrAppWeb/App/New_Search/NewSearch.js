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
            $('#check_column').click(function () {
                checkColumn();
            })
            $('#add_to_search').click(function () {
                addToSearch($('#get_param').val());
            })
            $('#complete_entry').click(function () {
                completeEntry();
            })
        });
    };

    //stores the current search array of form column-dataToSearchFor column being index 0 of internal array ~Thea
    var MasterSearch = [];
    var CurrentSearch = [];
    var column = "";
    var col_saved = false;
    var term_count = 0;

    //Function that selects the column and fills the canvas with the column letter. Dan
    function selectColumn() {
        col_saved = true;
        column = $('#get_col').val();
        var c=document.getElementById("myCanvas");
        var ctx = c.getContext("2d");
        ctx.clearRect(0,0,200,100);
        ctx.font = "20px Georgia";
        ctx.fillText("Column " + column.toUpperCase() + " selected", 10, 50);
        CurrentSearch = [];
        //Then want to highlight column is possible -- will most likely require input
        //from initial data selection.
        //Office.context.document.setSelectedDataAsync(Office.CoercionType.Matrix)  
    }

    //Function that checks to make sure the user entered in an acceptable column. Dan
    function checkColumn() {
        if (column = "" || col_saved == false) {

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
        col_saved = false;
        return true;

    }
    
    //Adds terms one at a time to search, dynamically generating buttons to show the user what is in
    //the current search. Clicking on the buttons makes them disappear and then they are removed from the
    //current search as well. Dan
    function addToSearch(term) {
        term_count++;
        if (term != -1) {
            //terms pushed to current search as they are entered 
            CurrentSearch.push(term);
            app.showNotification(CurrentSearch);
        }

        var $button = $('<button/>', {
            type: 'button',
            'class': 'dynBtn',
            id: term_count,
            text: term,
            value: term,
            click: function () {
                app.showNotification("Removed term '" + $(this).attr("value") + "' from search");
                
                var index = CurrentSearch.indexOf($(this).attr("value"));
                if (index > -1) {
                    //terms spliced from current search as they are deleted
                    CurrentSearch.splice(index, 1);
                }
                $(this).remove();
                
            }
        });
        $button.appendTo('#button_div');
        $('#get_param').val("");
    }

    function completeEntry() {
        app.showNotification(column);
        var newColIndex = convertLettersToNumbers(column);
        MasterSearch = [newColIndex, CurrentSearch];
        app.showNotification(MasterSearch);
        localStorage["CurrentSearch"] = JSON.stringify(MasterSearch);

    }

       

    function checkLocal() {
        CurrentSearch = localStorage.getItem("CurrentSearch");
        if (CurrentSearch || CurrentSearch === "") {
            CurrentSearch = JSON.parse(localStorage["CurrentSearch"]);
        }
        else {
            CurrentSearch = [];
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