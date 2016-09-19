// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

(function () {
     

    var resource = 'https://outlook.office.com';
    var endpoint = 'https://outlook.office.com/api/v2.0/me/mailfolders/inbox/messages?$top=10';
    var clientID = 'xxxxx';
    var tenantName = 'xxxxx';

    // Enter Global Config Values & Instantiate ADAL AuthenticationContext
    window.config = {
        instance: 'https://login.microsoftonline.com/',
        tenant: tenantName,
        clientId: clientID,
        postLogoutRedirectUri: window.location.origin,
        cacheLocation: 'localStorage'
    };
    
    var authContext = new AuthenticationContext(config);

    // Get UI jQuery Objects
    var $panel = $(".panel-body");
    var $userDisplay = $(".app-user");
    var $signInButton = $(".app-login");
    var $signOutButton = $(".app-logout");
    var $errorMessage = $(".app-error");

    // Check For & Handle Redirect From AAD After Login
    var isCallback = authContext.isCallback(window.location.hash);
    authContext.handleWindowCallback();
    $errorMessage.html(authContext.getLoginError());

    if (isCallback && !authContext.getLoginError()) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
    }

    // Check Login Status, Update UI
    var user = authContext.getCachedUser();
    if (user) {
        $userDisplay.html(user.userName);
        $userDisplay.show();
        $signInButton.hide();
        $signOutButton.show();


        ////call rest endpoint
        //authContext.acquireToken(resource, function (error, token) {
        //    if (error || !token) {
        //        jQuery("#loginMessage").text('ADAL Error Occurred: ' + error);
        //        return;
        //    }

        //    // Get calendar events
        //    jQuery.ajax({
        //        type: 'GET',
        //        url: calenderEndpoint,
        //        headers: {
        //            'Accept': 'application/json',
        //            'Authorization': 'Bearer ' + token,
        //        },
        //    }).done(function (data) {
        //        jQuery("#restDataEventName").text(data.value[0].Subject);
        //        jQuery("#restDataEventStart").text(data.value[0].Start.DateTime + ' ' + data.value[0].Start.TimeZone);
        //        alert(JSON.stringify(data.value[0].Subject));

        //    });
        //});
        }else {
        $userDisplay.empty();
        $userDisplay.hide();
        $signInButton.show();
        $signOutButton.hide();
    }



    // Handle Navigation Directly to View
    window.onhashchange = function () {
        loadView(stripHash(window.location.hash));
    };
    window.onload = function () {
        $(window).trigger("hashchange");
    };

    // Register NavBar Click Handlers
    $signOutButton.click(function () {
        authContext.logOut();
    });
    $signInButton.click(function () {
        authContext.login();
    });

    // Route View Requests To Appropriate Controller
    function loadCtrl(view) {
        console.log(view); 
        switch (view.toLowerCase()) {
            case 'home':
                return homeCtrl;
            case 'filesapi':
                return filesApiCtrl;
            case 'userdata':
                return userDataCtrl;
        }
    }

    // Show a View
    function loadView(view) {

        $errorMessage.empty();
        var ctrl = loadCtrl(view);

        if (!ctrl)
            return;

        // Check if View Requires Authentication
        if (ctrl.requireADLogin && !authContext.getCachedUser()) {
            authContext.config.redirectUri = window.location.href;
            authContext.login();
            return;
        }

        // Load View HTML
        $.ajax({
            type: "GET",
            url: "App/Views/" + view + '.html',
            dataType: "html",
        }).done(function (html) {

            // Show HTML Skeleton (Without Data)
            var $html = $(html);
            $html.find(".data-container").empty();
            $panel.html($html.html());
            ctrl.postProcess(html);

        }).fail(function () {
            $errorMessage.html('Error loading page.');
        }).always(function () {

        });
    };

    function stripHash(view) {
        return view.substr(view.indexOf('#') + 1);
    }

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


