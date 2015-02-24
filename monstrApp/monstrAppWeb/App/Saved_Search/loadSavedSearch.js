(function () {
    "use strict";
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            getNamesSaved();
        });
    };

    //saved_searches
    function getNamesSaved() {
        var names = Office.context.document.settings.get('search_names');
        if (names || names === "") {
            JSON.parse(names);
            select = document.getElementById('names');
            for (name in names) {
                select.add(new Option(names[name]));
            }
        }
        else {
            app.showNotification("Error:","You must save a search before you can use this menu")
        }
    }

})();