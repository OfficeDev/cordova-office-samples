// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.
var app = {
    outlookClient: null,

    initialize: function() {
        this.bindEvents();
    },

    renderMessages: function (messages) {
        var view = document.getElementById('messages-list');

        var html = messages.map(function(msg) {
            return '<li class="topcoat-list__item email">'
                // from
                + '<div class="recipient">' + ((msg.Sender.EmailAddress && msg.Sender.EmailAddress.Address) ? msg.Sender.EmailAddress.Address : '') + '</div>'
                // subject
                + '<div class="subject">' + msg.Subject + '</div>'
                // body
                + '<div class="body">' + msg.BodyPreview + '</div>';
        }).join('</br>');

        view.innerHTML = html;
    },

    loadMessages: function (folder) {

        // for Sent items correct folder name in outlook is SentItems
        folder = (folder == 'Sent') ? 'sentItems' : folder;

        var outlookFolder = app.outlookClient.me.folders.getFolder(folder);
        
        outlookFolder.messages.getMessages().fetchAll().then(app.renderMessages, app.onError);
    },
    bindEvents: function() {
        document.addEventListener('deviceready', this.onDeviceReady, false);

        function toggleMenu() {
            // menu must be always shown on desktop/tablet
            if (document.body.clientWidth > 480) return;
            var cl = document.body.classList;
            if (cl.contains('left-nav')) { cl.remove('left-nav'); }
            else {cl.add('left-nav');}
        }
        document.getElementById('slide-menu-button').addEventListener('click', toggleMenu);

        Array.prototype.forEach.call(document.getElementsByClassName('side-nav__button'), function (el) {
            el.addEventListener('click', function (e) {
                toggleMenu();
                app.loadMessages(e.srcElement.innerText);
            });
        });
    },

    onDeviceReady: function () {

        var resourceUrl = 'https://outlook.office365.com';
        var officeEndpointUrl = 'https://outlook.office365.com/ews/odata';

        var appId = '14b0c641-7fea-4e84-8557-25285eb86e43';
        var authUrl = 'https://login.windows.net/common/';
        var redirectUrl = 'http://localhost:4400/services/office365/redirectTarget.html';

        var authContext = new Microsoft.ADAL.AuthenticationContext(authUrl);

        app.outlookClient = new Microsoft.OutlookServices.Client(officeEndpointUrl, authContext, resourceUrl, appId, redirectUrl);
        app.loadMessages('Inbox');
    },

    onError: function (err) {
        console.log(err);
    }
};