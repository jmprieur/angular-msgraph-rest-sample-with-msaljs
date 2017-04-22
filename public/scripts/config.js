/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

(function () {
    angular
        .module('app')
        .service('config', [function () {
            return {
                clientID: '3e9dc15e-8b28-4512-831e-7ded1276c4e8',
                redirectUri: "http://localhost:8080/",
                interactionMode: "popUp",
                graphEndpoint: "https://graph.microsoft.com/v1.0/me",
                graphScopes: ["user.read", "mail.send"]
            }
        }]);
})();