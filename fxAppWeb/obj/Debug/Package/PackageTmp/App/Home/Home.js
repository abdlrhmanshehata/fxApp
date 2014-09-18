/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();
            $('#get-data-from-selection').click(getDataFromSelection);
            $("#writeDataBtn").click(function (event) {
                writeData();
            });
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }
    function writeData() {
        Office.context.document.setSelectedDataAsync([["red"], ["black"], ["blue"]], function (asyncResult) {
            if (asyncResult.status === "failed") {
                writeToPage('Error: ' + asyncResult.error.message);
            }
        });
    }
    function writeToPage(text) {
        document.getElementById('results').innerText = text;
    }
})();