// For an introduction to the Blank template, see the following documentation:
// http://go.microsoft.com/fwlink/?LinkID=397704
// To debug code on page load in Ripple or on Android devices/emulators: launch your app, set breakpoints, 
// and then run "window.location.reload()" in the JavaScript Console.
(function () {
	"use strict";

	document.addEventListener('deviceready', onDeviceReady.bind(this), false);

	function onDeviceReady() {
		// Handle the Cordova pause and resume events
		document.addEventListener('pause', onPause.bind(this), false);
		document.addEventListener('resume', onResume.bind(this), false);

		// TODO: Cordova has been loaded. Perform any initialization that requires Cordova here.

	};

	function onPause() {
		// TODO: This application has been suspended. Save application state here.
	};

	function onResume() {
		// TODO: This application has been reactivated. Restore application state here.
	};
})();

function onLoad(content, type) {
	var endPointUri, resourceID;
	if (!O365Auth || !O365Auth.Context) { log("No O365Auth instantiated"); return; }
	if (document.getElementById("msaccount").checked) {
		O365Auth.Settings.clientId = "cb17c17a-87e3-40cd-9ab0-38cbe8c93cda";
		resourceID = "https://graph.windows.net/";
		endPointUri = "https://outlook.office.com/api/v2.0/";
	}
	else {
		O365Auth.Settings.clientId = "92f98787-c980-4c15-9be0-348ba4244408";
		resourceID = "https://outlook.office365.com/";
		endPointUri = "https://outlook.office365.com/api/v1.0/";
	}
	var authContext = new O365Auth.Context();
	if (content == "id") {
		authContext.getIdToken(resourceID).then(function (token) {
			document.getElementById("name").innerHTML = "Hello, " + token.givenName + " " + token.familyName + ". We know who you are ;-)";
		}, function (error) {
			log('Failed to get ID token. Error = ' + error.message);
		});
		return;
	}

	if (type == "REST") {
		authContext.getAccessToken(resourceID).then(function (accessToken) {
			if (content == "messages") {
				var requestUri = endPointUri + 'me/folders/inbox/messages?$top=20'; // or folders('inbox')/messages
				var bearerToken = "Bearer " + accessToken;
				$.ajax(requestUri, {
					headers: {
						"Authorization": bearerToken,
						"Accept": "application/json;odata.metadata=minimal"
					}
				}).then(function (response) {
					var messages = "";
					if (response.value.length == 0) messages = "You have no messages"
					for (var i = 0; i < response.value.length; i++) {
						messages += "<li>" + response.value[i].Subject + "</li>";
					};
					document.getElementById("messages").innerHTML = messages;
				}).fail(function (error) {
					log(error.statusText + ' - ' + error.status);
				});
			}
			if (content == "contacts") {
				var requestUri = endPointUri + 'me/contacts?$top=20';
				var bearerToken = "Bearer " + accessToken;
				$.ajax(requestUri, {
					headers: {
						"Authorization": bearerToken,
						"Accept": "application/json;odata.metadata=minimal"
					}
				}).then(function (response) {
					var contacts = "";
					if (response.value.length == 0) contacts = "You have no contacts"
					for (var i = 0; i < response.value.length; i++) {
						contacts += "<li>" + response.value[i].DisplayName + "</li>";
					};
					document.getElementById("contacts").innerHTML = contacts;
				}).fail(function (error) {
					log(error.statusText + ' - ' + error.status);
				});
			}
			if (content == "contacts2") {
				var requestUri = endPointUri + 'me/contacts?$top=20';
				var bearerToken = "Bearer " + accessToken;
				var xhr = new XMLHttpRequest();
				xhr.open('GET', requestUri);
				xhr.setRequestHeader("Authorization", bearerToken);
				xhr.setRequestHeader("Accept", "application/json;odata.metadata=minimal");
				xhr.onload = function () {
					if (xhr.status === 200) {
						var response = JSON.parse(xhr.responseText);
						var contacts = "";
						for (var i = 0; i < response.value.length; i++) {
							contacts += "<li>" + response.value[i].DisplayName + "</li>";
						};
						document.getElementById("contacts").innerHTML = contacts;
					}
					else {
						log('Request failed.  Returned status of ' + xhr.status);
					}
				};
				xhr.send();
			}
		}, function (error) {
			log('Failed to get access token. Error = ' + error.message);
		});
	}
	else { // Office365 wrapper API

		if (!Microsoft.OutlookServices) { log("Couldn't get Outlook Services."); return; }
		var outlookClient = new Microsoft.OutlookServices.Client(endPointUri, authContext.getAccessTokenFn(resourceID));

		if (content == "contacts") {
			outlookClient.me.contacts.getContacts().fetch(50).then(function (result) {
				var contacts = "";
				result.currentPage.forEach(function (contact) {
					contacts += "<li>" + contact.fileAs + "</li>";
				});
				document.getElementById("contacts").innerHTML = contacts;
			}, function (error) {
				log('Failed to get contacts. Error = ' + error.message);
			});
		}
		if (content == "messages") {
			outlookClient.me.folders.getFolder("Inbox").fetch()
		.then(function (folder) {
			// Retrieve all the messages
			folder.messages.getMessages().fetch()
			.then(function (mails) {
				// mails.currentPage contains all the mails in Inbox
				var messages = "";
				mails.currentPage.forEach(function (message) {
					messages += "<li>" + message.subject + "</li>";
				});
				document.getElementById("messages").innerHTML = messages;
			}, function (error) {
				console.log(error);
			});
		}, function (error) {
			console.log(error);
		});
		}
	}
}

function onClear() {
	document.getElementById("name").innerHTML = "";
	document.getElementById("messages").innerHTML = "";
	document.getElementById("contacts").innerHTML = "";
	document.getElementById("error").innerHTML = "";
}

function onLogout() {
	var authContext = new O365Auth.Context();
	authContext.logOut();
}

function log(error) {
	document.getElementById("error").innerHTML = error;
}
