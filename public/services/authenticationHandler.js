(function () {
	angular
        .module('app')
        .service('authenticationHandler', ['$q', 'config', function ($q, config) {
        	this.userAgentApplication = null;

        	this.isAuthenticated = function () {
        		var scope = config.graphScopes;

        		if (this.userAgentApplication) {
        			var userAgentApplication = this.userAgentApplication;
        		}
        		else {
        			var userAgentApplication = new Msal.UserAgentApplication(config.clientID, null, null);
        			this.userAgentApplication = userAgentApplication;
        		}

        		return this.userAgentApplication.getUser() != null;
        	};

        	this.loginAndGetAccessToken = function (scope) {
        		return this.loginViaPopup()
					.then(function () {
						return this.getAccessTokenViaPopup(scope);
					}.bind(this));
        	};

        	this.loginViaPopup = function () {
        		var deferred = $q.defer();
        		var userAgentApplication = this.userAgentApplication = new Msal.UserAgentApplication(config.clientID, null, function (error, token) {
        			if (token) {
        				deferred.resolve(token); // note, this is an id token, we cant call graph with it
        			}
        			else {
        				deferred.reject(error);
        			}
        		}.bind(this));

        		userAgentApplication.interactionMode = config.interactionMode;
        		userAgentApplication.redirectUri = config.redirectUri;

        		userAgentApplication.login("slice=testslice&uid=true");
        		return deferred.promise;
        	};

        	this.getAccessTokenViaPopup = function (scope) {
        		// assumes user is logged in

        		var userAgentApplication = this.userAgentApplication; // todo: create one if doesnt exist
        		var deferred = $q.defer();

        		userAgentApplication.acquireTokenSilent(scope, function (errorDesc, token, error) {
        			if (token) {
        				deferred.resolve(token);
        			} else {
        				userAgentApplication.interactionMode = 'popUp';
        				userAgentApplication.acquireToken(scope, function (error, token) {
        					if (token) {
        						deferred.resolve(token);
        					}
        					if (!token) {
        						deferred.reject(error);
        					}
        				});
        			}
        		});

        		return deferred.promise;
        	}
        }]);
})();