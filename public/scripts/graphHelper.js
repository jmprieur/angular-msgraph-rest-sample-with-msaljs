/* 
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. 
*  See LICENSE in the source repository root for complete license information. 
*/

"use strict";

(function () {
    angular
        .module('app')
        .service('GraphHelper', ['$http', function ($http) {
            
            // Initialize the auth request.
            clientApplication = createApplication(APPLICATION_CONFIG, function ()
            {
                // localStorage.user = clientApplication.user;
                callWebApi(APPLICATION_CONFIG.graphScopes, function (token, error)
                {
                    if (error == null) {
                        localStorage.token = angular.toJson(token);

                        // Add the required Authorization header with bearer token.
//                        $http.defaults.headers.common.Authorization = 'Bearer ' + token;

                    }
                });
            });

            return {

                // Sign in and sign out the user.
                login: function login() {
                    clientApplication.login();
                },
                logout: function logout() {
                    clientApplication.logout();
                    delete localStorage.token;
                    delete localStorage.user;
                },

                // Get the profile of the current user.
                me: function me() {
                    return $http.get('https://graph.microsoft.com/v1.0/me');
                },

                // Send an email on behalf of the current user.
                sendMail: function sendMail(email) {
                    return $http.post('https://graph.microsoft.com/v1.0/me/sendMail', { 'message': email, 'saveToSentItems': true });
                }
            }
        }]);
})();