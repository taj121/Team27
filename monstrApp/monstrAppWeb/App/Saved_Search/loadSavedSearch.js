(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            getNamesSaved();
            $('#run_selected').click(loadSelectedSearch());
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
        localStorage["currentSearch"] =search;
    }

})();