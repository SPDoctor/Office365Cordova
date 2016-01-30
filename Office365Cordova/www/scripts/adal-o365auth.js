// Override o365auth.js to replace with adal.js for web client applications

var O365Auth;
(function (O365Auth) {
	var Context = (function () {
		function Context(authUri, redirectUri) {
			this._redirectUri = 'http://localhost/';
			if (!authUri) {
				if (O365Auth.Settings.authUri) {
					this._authUri = O365Auth.Settings.authUri;
				} else {
					throw new Microsoft.Utility.Exception('No authUri provided nor found in O365Auth.authUri');
				}
			} else {
				this._authUri = authUri;
			}
			if (this._authUri.charAt(this._authUri.length - 1) !== '/') {
				this._authUri += '/';
			}
			if (!redirectUri) {
				if (O365Auth.Settings.redirectUri) {
					this._redirectUri = O365Auth.Settings.redirectUri;
				}
			} else {
				this._redirectUri = redirectUri;
			}
		}

		Context.prototype.getDeferred = function () {
			if (O365Auth.deferred) {
				return O365Auth.deferred();
			}

			return new Microsoft.Utility.Deferred();
		};

		Context.prototype.ajax = function (url, data, verb) {
			var deferred = new Microsoft.Utility.Deferred(), xhr = new XMLHttpRequest();

			if (!verb) {
				verb = 'GET';
			}

			xhr.open(verb.toUpperCase(), url, true);

			xhr.setRequestHeader('Content-Type', 'application/x-www-form-urlencoded; charset=UTF-8');
			xhr.setRequestHeader('Accept', '*/*');

			xhr.onreadystatechange = function (e) {
				if (xhr.readyState == 4) {
					if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 304) {
						deferred.resolve(xhr.responseText);
					} else {
						deferred.reject(xhr);
					}
				} else {
					deferred.notify(xhr.readyState);
				}
			};

			xhr.send(data);

			return deferred;
		};

		Context.prototype.post = function (url, data) {
			return this.ajax(url, data, 'POST');
		};

		Context.prototype.getAccessTokenFromRefreshToken = function (resourceId, refreshToken, clientId) {
			var deferred = this.getDeferred(), url = this._authUri + 'oauth2/token', data = 'grant_type=refresh_token&refresh_token=' + encodeURIComponent(refreshToken) + '&client_id=' + encodeURIComponent(clientId) + (resourceId ? '&resource=' + encodeURIComponent(resourceId) : '');

			this.post(url, data).then(function (result) {
				var jsonResult = JSON.parse(result), access_token = {
					token: jsonResult.access_token,
					expires_in: new Date((new Date()).getTime() + (jsonResult.expires_in - 300) * 1000)
				};


				// cache most recent refresh token if available.
				deferred.resolve(access_token.token);
			}.bind(this), function (xhr) {
				deferred.reject(new Microsoft.Utility.HttpException(xhr));
			});

			return deferred;
		};

		Context.prototype.isLoginRequired = function (resourceId, clientId) {
			if (!clientId) {
				if (O365Auth.Settings.clientId) {
					clientId = O365Auth.Settings.clientId;
				} else {
					throw new Microsoft.Utility.Exception('clientId was not provided nor found in O365Auth.clientId');
				}
			}

			if (resourceId) {
				var access_token;
				if (access_token && access_token.expires_in > new Date()) {
					return false;
				}
			}

			return true;
		};

		Context.prototype.getAccessToken = function (resourceId, loginHint, clientId, redirectUri) {
			return AuthenticationContext.acquireToken(resourceId);

		};

		Context.prototype.getAccessTokenFn = function (resourceId, loginHint, clientId, redirectUri) {
			return function () {
				return this.getAccessToken(resourceId, loginHint, clientId, redirectUri);
			}.bind(this);
		};

		Context.prototype.getIdToken = function (resourceId, loginHint, clientId, redirectUri) {
		};

		Context.prototype.logOut = function (clientId) {
		};

		return Context;
	})();
	O365Auth.Context = Context;
})(O365Auth || (O365Auth = {}));
