(function() {
    window.config = {
        instance: 'https://login.microsoftonline.com/',
        tenant: 'common',
        clientId: 'cfc8d4ef-105f-4d4c-82ef-56a74624197b',
        postLogoutRedirectUri: window.location.origin,
        cacheLocation: 'localStorage',  //enable this for IE, as sessionStorage does not work for localhost
        resource: 'https://graph.microsoft.com'
    }

    var authContext = new AuthenticationContext(config);

    // Get UI objects
    var errorMessage = document.getElementById("app-error");

    wireUpCommands();

    // check for and handle redirect from AAD after login
    var isCallback = authContext.isCallback(window.location.hash);
    authContext.handleWindowCallback();
    errorMessage.innerText = authContext.getLoginError();

    if (isCallback && !authContext.getLoginError()) {
        window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
    }

    // Check login status, update UI
    var user = authContext.getCachedUser();
    setLoggedInUser(user);


    // Handle Navigation Directly to View
    window.onhashchange = function () {
        loadView(stripHash(window.location.hash));
    };
    window.onload = function () {
        $(window).trigger("hashchange");
    };


    function wireUpCommands() {
        var loginButton = document.getElementById("buttonSignIn");
        loginButton.onclick = signInButtonClicked;
    }

    function setLoggedInUser(user) {
        var loginButton = document.getElementById("buttonSignIn");
        var loginLabel = document.getElementById("labelSignIn");

        if (user) {
            loginLabel.innerText = "Sign out " + user.userName;
            loginButton.action = "sign_out";
        }
        else
        {
            loginLabel.innerText = "Sign in";
            loginButton.action = "sign_in";
        }
    }

    function signInButtonClicked() {
        var loginButton = document.getElementById("buttonSignIn");
        if (loginButton.action == "sign_out") {
            authContext.logOut();
        } else {
            authContext.login();
        }
    }

    function loadView(view) {

    }

    function stripHash(view) {
        return view.substr(view.indexOf('#') + 1);
    }

}())