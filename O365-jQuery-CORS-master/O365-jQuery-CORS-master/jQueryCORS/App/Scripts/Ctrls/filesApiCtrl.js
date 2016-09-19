// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

(function () {


    var resource = 'https://outlook.office.com';
    var calenderEndpoint = 'https://outlook.office.com/api/v2.0/me/calendarview?startDateTime=2016-09-10&endDateTime=2016-09-19';
    var postCalenderEndpoint = 'https://outlook.office.com/api/v2.0/me/events';



    function refreshViewData() {

        // Empty old view contents.
        var $dataContainer = $(".data-container");
        $dataContainer.empty();
        var $loading = $(".view-loading");
        console.log("Fetching files from OneDrive...");
        var authContext = new AuthenticationContext(config);

        //call rest endpoint
        authContext.acquireToken(resource, function (error, token) {
            if (error || !token) {
                jQuery("#loginMessage").text('ADAL Error Occurred: ' + error);
                return;
            }

            // Get calendar events
            jQuery.ajax({
                type: 'GET',
                url: calenderEndpoint,
                headers: {
                    'Accept': 'application/json',
                    'Authorization': 'Bearer ' + token,
                },
            }).done(function (data) {
                jQuery("#restDataEventName").text(data.value[0].Subject);
                jQuery("#restDataEventStart").text(data.value[0].Start.DateTime + ' ' + data.value[0].Start.TimeZone);
                console.log('Successfully got calendar data');
                console.log(data);

                var $html = $(viewHTML);
                var $template = $html.find(".data-container");
                var output = '';

                for (var i = 0, len = data.value.length; i < len; i++) {
                    var $entry = $template;

                    $entry.find(".view-data-type").html('Subject');
                    $entry.find(".view-data-name").html(data.value[i].Subject);
                    output += $entry.html();
                }
                
                // Update the UI.
                $loading.hide();
                $dataContainer.html(output);
            });
        });
    }


    function clearErrorMessage() {
        var $errorMessage = $(".app-error");
        $errorMessage.empty();
    };

    function printErrorMessage(mes) {
        var $errorMessage = $(".app-error");
        $errorMessage.html(mes);
    }

    // Module definition. 
    window.filesApiCtrl = {
        requireADLogin: true ,
        preProcess: function (html) {
        },
        postProcess: function (html) {
            viewHTML = html;
            refreshViewData();
        },
    };
}());

//*********************************************************  
//  
//O365 jQuery and CORS Sample, https://github.com/OfficeDev/O365-jQuery-CORS
// 
//Copyright (c) Microsoft Corporation 
//All rights reserved.  
// 
//MIT License: 
// 
//Permission is hereby granted, free of charge, to any person obtaining 
//a copy of this software and associated documentation files (the 
//""Software""), to deal in the Software without restriction, including 
//without limitation the rights to use, copy, modify, merge, publish, 
//distribute, sublicense, and/or sell copies of the Software, and to 
//permit persons to whom the Software is furnished to do so, subject to 
//the following conditions: 
// 
//The above copyright notice and this permission notice shall be 
//included in all copies or substantial portions of the Software. 
// 
//THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND, 
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF 
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND 
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE 
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION 
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION 
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE. 
//  
//********************************************************* 