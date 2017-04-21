/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

"use strict";

(function () {
    angular
        .module('app')
        .service('GraphHelper', ['$http', 'authenticationHandler', 'config', function ($http, authenticationHandler, config) {
            return {
                // Sign in and sign out the user.
                login: function login() {
                    return authenticationHandler.loginAndGetAccessToken(config.graphScopes)
                        .then(this.meInternal.bind(this));
                },
                logout: function logout() {
                    authenticationHandler.userAgentApplication.logout();
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
                    return authenticationHandler.getAccessTokenViaPopup(config.graphScopes).then(this.meInternal.bind(this));
                },

                // Send an email on behalf of the current user.
                sendMail: function sendMail(email) {
                    // Initialize the auth request.
                    // localStorage.user = userAgentApplication.user;
                    authenticationHandler.getAccessTokenViaPopup(config.graphScopes).then(function (token) {
                        setDefaultHeaders(token);

                        return $http.post('https://graph.microsoft.com/v1.0/me/sendMail',
                            { 'message': email, 'saveToSentItems': true });
                    });
                }
            }
        }]);
})();