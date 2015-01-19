/*
 * Licensed to the Apache Software Foundation (ASF) under one
 * or more contributor license agreements.  See the NOTICE file
 * distributed with this work for additional information
 * regarding copyright ownership.  The ASF licenses this file
 * to you under the Apache License, Version 2.0 (the
 * "License"); you may not use this file except in compliance
 * with the License.  You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied.  See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */
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

        var authContext = new O365Auth.Context(authUrl, redirectUrl);

        app.outlookClient = new Microsoft.OutlookServices.Client(officeEndpointUrl, authContext.getAccessTokenFn(resourceUrl, '', appId));
        app.loadMessages('Inbox');
    },

    onError: function (err) {
        console.log(err);
    }
};