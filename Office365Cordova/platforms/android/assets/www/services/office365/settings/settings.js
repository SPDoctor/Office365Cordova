
var O365Auth;
(function (O365Auth) {
    (function (Settings) {
        Settings.clientId = '92f98787-c980-4c15-9be0-348ba4244408';
        Settings.authUri = 'https://login.microsoftonline.com/common/';
        Settings.redirectUri = 'http://localhost:4400/services/office365/redirectTarget.html';
        Settings.domain = 'flosim.net';
    })(O365Auth.Settings || (O365Auth.Settings = {}));
    var Settings = O365Auth.Settings;
})(O365Auth || (O365Auth = {}));
