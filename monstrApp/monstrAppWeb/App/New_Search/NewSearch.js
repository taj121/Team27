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
                checkEntry();
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


    function addNewField() {

        if (curTerms.length != 0) {
            var searchCol = [newColIndex, curTerms]; // [X,[1,2,3]]
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
                id: str,
                text: "Searching Col: " + newColLetter + " for " + str,
                value: "Searching Col: " + newColLetter + " for " + str,
                click: function () {
                    app.hideAllNotification();
                    app.showNotification("Working on deleting backend");

                    //var index = curTerms.indexOf($(this).attr("value"));
                    //if (index > -1) {
                    //terms spliced from current search as they are deleted
                    //    curTerms.splice(index, 1);
                    //}
                    //$(this).remove();

                }
            });
            $button.appendTo('#all_terms');
            $('#all_terms').show();
            curTerms = [];
            $('#get_col').val('');
            $('#word_num_menu').hide();
            $('#end_btns').hide();
            $("#final_btns").show();
            $("#show_col").hide();
            $('#add_another').removeAttr('disabled');
            
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
        $('#add_another').attr('disabled', 'disabled');

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
            app.showNotification("Error:", "Not a valid column entry");
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
                        app.showNotification("Removed the following term from search:'" , $(this).attr("value") );

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

    function addRange(rbegin, rend) {
        var rangeContainer = [rbegin, rend] //to hold lo/hi of range
        if (colAdded == 1) {
            if (rbegin != "" && rend != "") {
                if (rbegin != -1 && rend != -1) {
                    if (rbegin <= rend) {
                        //terms pushed to current search as they are entered
                        var str = ""
                        for (var i = rbegin; i <= rend; i++) {
                            str += i + ", "
                        }
                        curTerms.push(rangeContainer);
                        //i envision this as 1)check if entry is a list 2)if so do an or with the elements in that list
                        var $button = $('<button/>', {
                            type: 'button',
                            'class': 'dynBtn',
                            id: rbegin + "-" + rend,
                            text: "range: " + rbegin + "-" + rend,
                            value: "range: " + rbegin + "-" + rend,
                            click: function () {
                                app.hideAllNotification();
                                app.showNotification("Removed the following term from search:'", $(this).attr("value"));

                                for (var i = rbegin; i <= rend; i++) {
                                    var index = curTerms.indexOf(i);
                                    if (index > -1) {
                                        //terms spliced from current search as they are deleted
                                        curTerms.splice(index, 1);
                                    } 
                                }
                                $(this).remove();
                                

                            }
                        });
                        $button.appendTo('#all_terms');
                        $('#range_begin').val("");
                        $('#range_end').val("");
                    } else {
                        app.hideAllNotification();
                        app.showNotification("Error:", "The beginning of the range must be a lower value than the end of the range.")
                    }
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
        app.showNotification("Searching for: " + str);
    }

    function checkEntry() {
        if (curTerms.length > 0) {
            app.hideAllNotification();
            app.showNotification("Error:", "Please add this entry or delete it before continuing");
            
        } else {
            app.hideAllNotification();
            window.location = "/App/New_Search/Bindings.html";
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

    /*Function to convert currentSearch[] to a String, for printing search details
    -Mark
    */
    function searchToString() {
        var details;
        for (var i = 0; i < currentSearch.length; i++) {
            details += "-Search in column " + currentSearch[i][0] + "for:\n  -"
            for (var j = 1; j < currentSearch[i].length; i++) {

                if (currentSearch[i].length == 1) {
                    details += currentSearch[i][j] + ". \n \n"
                }
                else{
                    if (j == currentSearch[i].length) {
                        details += "and" + currentSearch[i][j] + ". \n \n"
                    }
                    else {
                        details += currentSearch[i][j] + ", ";
                    }
                }

            }
            return details;
        }
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