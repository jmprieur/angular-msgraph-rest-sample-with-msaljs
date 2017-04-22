/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

"use strict";

(function () {
    angular
        .module('app')
        .service('GraphHelper', ['$http', 'config', function ($http, config) {
            var factory = {
                _userAgentApplication: null,
                get userAgentApplication() {
                    if (!this._userAgentApplication) {
                        this._userAgentApplication = new Msal.UserAgentApplication(config.clientID, null, null);
                    }
                    return this._userAgentApplication;
                },
                set userAgentApplication(val) {
                    this._userAgentApplication = val;
                },
                isAuthenticated: function () {
                    return this.userAgentApplication.getUser() != null;
                },
                // Sign in the user and get user profile
                login: function login() {
                    return this.userAgentApplication.loginPopup(config.graphScopes)
                        .then(function (idToken) {
                            return this.userAgentApplication.acquireTokenSilent(config.graphScopes);
                        }.bind(this))
                        .then(this.meInternal.bind(this));
                },
                logout: function logout() {
                    this.userAgentApplication.logout();
                    delete localStorage.token;
                    delete localStorage.user;
                },

                meInternal: function (token) {
                    this.setDefaultHeaders(token);

                    return $http.get('https://graph.microsoft.com/v1.0/me').then(function (response) {
                        return response.data;
                    });
                },

                setDefaultHeaders: function (token) {
                    // Add the required Authorization header with bearer token.
                    $http.defaults.headers.common.Authorization = 'Bearer ' + token;

                    // This header has been added to identify our sample in the Microsoft Graph service. If extracting this code for your project please remove.
                    $http.defaults.headers.common.SampleID = 'angular-connect-rest-sample';
                },

                // Get the profile of the current user.
                me: function me() {
                    return this.userAgentApplication.acquireTokenSilent(config.graphScopes)
                        .then(this.meInternal.bind(this));
                },

                // Send an email on behalf of the current user.
                sendMail: function sendMail(email) {
                    return this.userAgentApplication.acquireTokenSilent(config.graphScopes)
                        .then(function (token) {
                            this.setDefaultHeaders(token);

                            return $http.post('https://graph.microsoft.com/v1.0/me/sendMail',
                                { 'message': email, 'saveToSentItems': true });
                        }.bind(this));
                }
            }

            return factory;
        }]);
})();