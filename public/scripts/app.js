/*
*  Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
*  See LICENSE in the source repository root for complete license information.
*/

(function () {
    angular.module('app', [
        'angular-loading-bar'
    ]);

    // Set up a callback on the logger to log to the console
    var logger = new Msal.Logger("4a2152af-37b7-431f-aca8-896f0b9a5b8d");
    logger.level = Msal.LogLevel.Verbose;
    logger.localCallback = console.log;
})();