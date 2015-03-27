(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#get_col').focus()
            app.initialize();
            //will add a new column to search in, initially will have nothing to search for ~Thea
            //$('#get-new-column').click(addNewColumn);
            //adds new values to the most resent column added to current search ~Thea
            //$('#ch-param').click(addNewParam);
            checkLocal(); //load the current search
            $('#select_column').click(function () {
                selectColumn();
            });

            $('#OK_single_input').click(function () {
                addTerm($('#single_input').val()); 
            });
            

            $('#submit_input').click(function () {
                if (checkEntry()) {
                    getResults();
                }
            });

            $('#col_submit').click(function () {
                colSubmit($('#get_col').val());
            })

            $('#name_submit').click(function () {
                saveSearch($('#save_name').val());
            })
            $('#run_again').click(function () {
                addNewField();
            })
            $('#add_another').click(function () {
                addAnother();
            })
        });
    };

    var newColIndex;
    var newColLetter;
    var num_cols = 0;
    //stores the current search array of form column-dataToSearchFor column being index 0 of internal array ~Thea
    var currentSearch = [];
    var curTerms = [];
    var curBts = [];
    var searchData = [];
    var colAdded = 0; //boolean value to check if the user has entered a column ~Thea

    function getResults() {
        var data = JSON.parse(localStorage["userDataSelection"]);
        console.log('data getResults() ' + data);
        var temp = localStorage["currentSearch"];
        var searchFor = JSON.parse(temp);
        var newData = [];
        Debug.writeln(searchFor);
        Debug.writeln(data.value);


        for (var i = 0; i < searchFor.length; i++) {
            var column = searchFor[i][0];
            var toFind = searchFor[i][1];
            console.log('toFind ' + toFind);
            var rowsToGet = findIndex(toFind, column, data);
            for (var k = 0; k < rowsToGet.length; k++) {
                console.log('data.value before ' + data.value);
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

    function addNewField() {

        if (curTerms.length != 0) {
            var searchCol = [newColIndex, curTerms, num_cols]; // [X,[1,2,3],Z]
            currentSearch.push(searchCol);
            localStorage["currentSearch"] = JSON.stringify(currentSearch);

            var str = curTerms[0];
            $('#' + String(curTerms[0])).remove();
            for (var i = 1; i < curTerms.length; i++) {
                $('#' + String(curTerms[i])).remove();
                str += ", " + curTerms[i];
            }
            var $button = $('<button/>', {
                type: 'button',
                'class': 'dynBtn',
                id: num_cols,
                text: "Searching Col: " + newColLetter + " for " + str,
                value: num_cols,
                click: function () {
                    var idx;
                    for (var i = 0; i < currentSearch.length; i++) {
                        Debug.writeln(currentSearch[i][2] + " " + $(this).attr("value"));

                        if (currentSearch[i][2] == $(this).attr("value")) {
                            idx = i;
                        }
                    }
                    if (idx > -1) {
                        //terms spliced from current search as they are deleted
                        currentSearch.splice(idx, 1);
                    }
                    app.showNotification("Item removed from search");
                    $(this).remove();
                }
            });
            $button.appendTo('#all_terms');
            $('#all_terms').show();
            curTerms = [];
            $('#get_col').val('');
            $('#word_num_menu').hide();
            $('#end_btns').hide();
            $('#final_btns').show();
            $('#add_another').show();
            $('#submit_input').show();
            $('#submit_input').removeAttr('disabled');
            $("#show_col").hide();
            $('#add_another').removeAttr('disabled');
            $('#get_col').focus();
            
        }
        else {
            app.hideAllNotification();
            app.showNotification("Error:", "You must add at least one search value!");
        } 
    }

    function addAnother() {
        app.hideAllNotification
        $('#col_area').show();
        $('#show_col').contents().remove();
        $("#get_col").show();
        $("#col_submit").show();
        $('#add_another').hide();
        $('#submit_input').hide();
        $('#submit_input').attr("disabled", true);

    }

    /*
    Gets the column to search
    */
    function colSubmit(col) {
        newColIndex = $('#get_col').val();
        newColLetter = $('#get_col').val().toUpperCase();
        if (newColIndex == "") {
            app.hideAllNotification();
            app.showNotification("Error:", "No Column Selected");
        }
        else if (!(verifyColInput($('#get_col').val()))) {
            app.showNotification("Error:", "Invalid column");
        } else {
            newColIndex = convertLettersToNumbers(newColIndex);
            if (newColIndex < searchData.value[0].length) {
                app.hideAllNotification();
                $("#get_col").hide();
                $("#col_submit").hide();
                num_cols++;
                var $label = $('<label/>', {
                    type: 'label',
                    id: String(num_cols),
                    text: "Search column " + col.toUpperCase() + " for: ",
                });
                $label.appendTo("#show_col");
                $("#show_col").show();
                colAdded = 1;
                $('#col_area').hide();
                $('#word_num_menu').show();
                $('#single_input').focus();
                $('#end_btns').show();
              
            }
            //Thea -- added a check to make sure that column input is of valid characters
            //i.e. user can't add non-letters. Dan
            else {
                app.hideAllNotification();
                app.showNotification("Error:", "Column index is not in selected data range.");
            }
        }
        
    }

    //Check if column input is of valid letters
    function verifyColInput(column) {
        var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        for (var i = 0; i < column.length; i++) {
            if (alphabet.indexOf(column.substring(i, (i+1)).toUpperCase()) == -1) {
                return false;
            }
        }
        return true;
    }

    //adds terms one at a time to the search paramaters for a given column.
    //dynamically generates a button displaying each term.
    //removes the button from the page and the term from the search parameters
    //when the button is clicked. Dan
    function addTerm(term) {
        
        if (colAdded == 1) {
            if (term != "") { 
                if (term != -1) {
                    //terms pushed to current search as they are entered
                    curTerms.push(term);
                    app.hideAllNotification();
                    //app.showNotification("You are searching for:",curTerms);
                }
                //create button dynamically based on term entry
                var $button = $('<button/>', {
                    type: 'button',
                    'class': 'dynBtn',
                    id: String(term),
                    text: term,
                    value: term,
                    click: function () {
                        app.hideAllNotification();
                        app.showNotification("Removed the following term from search:" , $(this).attr("value") );

                        var index = curTerms.indexOf($(this).attr("value"));
                        if (index > -1) {
                            //terms spliced from current search as they are deleted
                            curTerms.splice(index, 1);
                        }
                        $(this).remove();

                    }
                });
                $button.appendTo('#word_num_menu');
                $('#single_input').val("");
                $('#single_input').focus();
            }
            else {
                app.hideAllNotification();
                app.showNotification("Error:", "You must enter a value to search for!");
            }
        } else {
            app.hideAllNotification();
            app.showNotification("Error:","You must submit a column before entering search values");
        }
    }


    function checkEntry() {
        if (curTerms.length > 0) {
            app.hideAllNotification();
            app.showNotification("Error:", "Please add this entry or delete it before continuing");
            return false;
        }
        else {
            return true;
        }
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
    
    /*
    loads current search from local storage to allow new parameters to be added
    loads data from local storage to allow check on column index to ensure it is valid
    */
    function checkLocal() {
        currentSearch = localStorage.getItem("currentSearch");
        if (currentSearch || currentSearch === "") {
            currentSearch = JSON.parse(localStorage["currentSearch"]);
        }
        else {
            currentSearch = [];
        }
        searchData = JSON.parse(localStorage.getItem("userDataSelection"));
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
    function to save a search, must be passed a name given by the user.
    untested since not sure if fron end is correct atm and also dont have time to add in dialog box for it right now
    ~Thea
    */
    function saveSearch(searchName) {
        //get names of searches already saved
        var names = Office.context.document.settings.get('search_names');
        if (names || names === "") {
            names=JSON.parse(names);
        }
        else {
            names = [];
        }
        //check if name is in use
        var nameFound = 0;
        for (var i = 0; i < names.length; i++) {
            if (searchName === names[i]) {
                app.hideAllNotification();
                app.showNotification("Error:", "This name is already in use.")
                nameFound = 1;
            }
        }
        //if the name isn't already in use
        if (nameFound == 0) {
            //save actuall search
            Office.context.document.settings.set(searchName, JSON.stringify(currentSearch));
            Office.context.document.settings.saveAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Settings save failed. Error: ' + asyncResult.error.message);
                } else {
                    app.showNotification('Settings saved.');
                }
            });
            names.push(searchName);
            //save name of search in array of names of search
            var temp = JSON.stringify(names);
            Office.context.document.settings.set('search_names', JSON.stringify(names));
            Office.context.document.settings.saveAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    app.showNotification('Settings save failed. Error: ' + asyncResult.error.message);
                } else {
                    app.showNotification('Settings saved.');

                    window.location = "/App/Home/Home.html";
                }
            });
        }
    }

})();