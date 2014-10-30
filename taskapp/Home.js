/// <reference path="../App.js" />

(function () {
   "use strict";
    // The initialize function must be run each time a new page is loaded
    // The initialize function is required for all apps.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
            // Add any initialization logic to this function.
            $('#write-rally-user-stories').click(getRallyStore);
        });
    }
    var rallyStoreData = null;
    Rally.onReady(function () {
        console.log('Rally.ready');
    });
    function getRallyStore(){
        Ext.create('Rally.data.wsapi.Store', {
            model: 'User Story',
            autoLoad: true,
            listeners: {
                scope: this,

                load: function(store, data, success) {
                    //Add Rally data to Office table format
                    writeRallyDataToOffice(data);
                }
            },
            fetch: ['Name', 'ScheduleState']
        });
    }

    function writeRallyDataToOffice(data) {
        var sampleTableData = new Office.TableData();
        sampleTableData.headers = ["Name", "ScheduleState"];
        Ext.each(data, function (d) {
            var row = [];
            row.push(d.get('Name'));
            row.push(d.get('ScheduleState'));
            sampleTableData.rows.push(row);
        });

        //Write the table to the office document 
        Office.context.document.setSelectedDataAsync(sampleTableData, { coercionType: "table" }, function (result) {
            if (result.status === Office.AsynResultStatus.Failed) {
                app.showNotification('Dang!  an Error!');
            }
        });
    }

})();