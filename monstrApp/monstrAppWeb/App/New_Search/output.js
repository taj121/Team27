(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            var toPrint = toString(getResults());
            $("#return-data").text(toPrint);
        });
    };

    /*
    function gets rows to return. 
    ~Thea
    */
    function getResults() {
        var data = JSON.parse(localStorage["userDataSelection"]);
        var temp = localStorage["currentSearch"];
        var searchFor = JSON.parse(temp);
        for (var i = 0; i < searchFor.length; i++) {
            var column = searchFor[i][0];
            var toFind = searchFor[i][1];
            var rowsToGet = findIndex(toFind, column, data);
            var newData = [];
            for (var k = 0; k < rowsToGet.length; k++) {
                if (data.value != undefined) {
                    for (var j = 0; j < data.value.length; j++) {
                        if (j == rowsToGet[k]) {
                            newData.push(data.value[j]);
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
        printToWindow(toString(data));
        return data;
    }
    /*
    Function to write out the rows to return in a nice way.
    ~Thea
    */
    function toString(returnData) {
        var returnString = "";
        for (var i = 0; i < returnData.length; i++) {
            for (var j = 0; j < returnData[i].length; j++) {
                returnString += (returnData[i][j] + "\t");
            }
            returnString += "\n";
        }
        return returnString;
    }
    
    /*
    function to write out return data to a new browser window, takes string version of data.
    ~Thea
    */
    function printToWindow(data) {
        var myWindow = window.open("", "MsgWindow", "width=500, height=500");
        myWindow.document.write("<pre>"+data+"</pre>");
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
                    var currentCheck = result.value[i][col].replace(/\s+/g, ''); //remove all whitespaces
                    if (toFind[j].toUpperCase() == currentCheck.toUpperCase()) {
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
    Function that will return a specific row of an array.
    rowNum = the number of the row you wish to reutrn 
    result = the array you wish to return the row from
    ~Thea
    */
    function returnRow(rowNum, result) {
        return result.value[rowNum];
    }

    /*
    Function that will return a set rows of an array.
    Return type is an array of rows
    rowNums = the numbers of the rows you wish to reutrn 
    result = the array you wish to return the row from
    ~Thea
    */
    function returnRows(rowNums, result) {
        var rows = [];
        for (var i = 0; i < rowNums.length; i++) {
            rows[i] = returnRow[i];
        }
        return rows;
    }
})();