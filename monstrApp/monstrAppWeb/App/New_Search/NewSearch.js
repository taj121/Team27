(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            //will add a new column to search in, initially will have nothing to search for ~Thea
            $('#get-new-column').click(addNewColumn);
            //adds new values to the most resent column added to current search ~Thea
            $('#add-search-param').click(addNewParam);
        });
    };

    //stores the current search array of form column-dataToSearchFor column being index 0 of internal array ~Thea
    var currentSearch = [];

    /*
    Function to convert letters to numbers, nigel
    */
    function convertLettersToNumbers(inputLetters) {
        var valueToReturn = 0;
        var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        for (var i = 0, j = inputLetters.length - 1; i < inputLetters.length; j--, i++) {
            valueToReturn += (alphabet.indexOf(inputLetters[i]) + 1) * Math.pow(alphabet.length, j);
            // hi, 'BA' should return 52 or? 'AAA' 702, as far as I can tell from the way 
            // excel sheet works anyways. As it increases 'A, B... Z, AA, AB,...AZ, BA, BB...
            // YA, ... ZA, AAA...
            // Let me know if it makes sense. Nigel

            //Nigel you are 100% right. Sorry I thought it went A..Z, AA..ZZ, AAA..ZZZ. Must have
            //been using a different version of excel at some point in my life. Dan

            //The last thing is does javascript differentiate lower and upper case? In java I'd
            //use: valueToReturn += (alphabet.indexOf(inputLetters[i].toUpperCase()) + 1)... etc.
            //just in case the user enters lower but idk about javascript.
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