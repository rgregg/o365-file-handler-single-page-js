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

    // Load the view, if one was specified
    var viewName = getParameterByName("view");
    if (viewName && viewName.length > 0)
    {
        loadView(viewName);
    }
    else
    {
        loadView(null);
    }

    function wireUpCommands() {
        document.getElementById("buttonSignIn").onclick = signInButtonClicked;
        document.getElementById("saveButton").onclick = saveButtonClicked;
        document.getElementById("renameButton").onclick = renameButtonClicked;
        document.getElementById("shareButton").onclick = shareButtonClicked;
        document.getElementById("closeButton").onclick = closeButtonClicked;
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

    function saveButtonClicked() {

    }

    function renameButtonClicked() {
        
    }

    function shareButtonClicked() {

    }

    function closeButtonClicked() {
        window.close();
    }

    function loadView(view) {
        var panel = document.getElementById("panel-body");
        var panelDefault = document.getElementById("panel-default");
        var cookieData = getCookie("fileHandlerActivation");
        var activationParameters = JSON.parse(cookieData);

        configureUxForView(view);

        if (view == "preview" || view == "open" || view == "newFile")
        {
            panelDefault.style.display = "none";

            // Launch the markdown previewer
            fetchFileFromMSGraph(activationParameters, view, function(view, text) {
                if (view == "preview")
                {
                    loadTextInPreview(text);
                }
                else if (view == "open" || view == "newFile")
                {
                    loadTextInEditor(text);
                }
                finishedLoading();
            });
        }
        else
        {
            panelDefault.style.display = "initial";
            panel.innerText = "";
            finishedLoading();
        }
    }

    function configureUxForView(view) {
        if (view == "preview")
        {
            // Hide the command bar for preview view
            var commandBar = document.getElementById("commandBar");
            commandBar.style.display = "none";
        }
        else if (view == "open" || view == "newfile")
        {
            // enable editor commands in toolbar
            let editorCommands = document.getElementById("editorActionButtons");
            editorCommands.style.display = "initial";
        }
    }

    function fetchFileFromMSGraph(activationParameters, view, callback)
    {
        var client = MicrosoftGraph.Client.init({
            authProvider: (done) => {
                authContext.acquireToken("https://graph.microsoft.com", (function(msg, token) {
                    done(null, token);
                }));
            }            
        });

        var itemUrl = activationParameters.items[0];
        client.api(itemUrl).get().then((res) => {
            var downloadUrl = res["@microsoft.graph.downloadUrl"];
            client.api(downloadUrl).getStream((err, req) => {
                // remove custom headers since it causes CORS issues and isn't required
                req.set('Authorization', null);
                req.set('SdkVersion', null);
                req.responseType('blob');
                req.end((err, res)=>{
                    var fileBlob = res.body;
                    var reader = new FileReader();
                    reader.onload = function() {
                        var text = reader.result;
                        callback(view, text);
                    }
                    reader.readAsText(fileBlob);
                })
            });
        });
    }

    function loadTextInEditor(text) {
        let textArea = document.getElementById("markdownContent");
        let editor = new SimpleMDE({
             element: textArea,
             previewRender: function(plainText) { return convertMarkdownToHtml(plainText); }
            });
        editor.value(text);
    }

    function loadTextInPreview(text) {
        // Load the HTML into body-panel
        let bodyPanel = document.getElementById("panel-body");
        bodyPanel.innerHTML = convertMarkdownToHtml(text);
    }

    function convertMarkdownToHtml(text) {
        let  converter = new showdown.Converter();
        converter.setFlavor('github');
        return converter.makeHtml(text);
    }

    function finishedLoading() {
        var loadingSpinner = document.getElementById("app-loading");
        loadingSpinner.style.display = "none";
    }

    // from https://www.w3schools.com/js/js_cookies.asp
    function getCookie(cname) {
        var name = cname + "=";
        var decodedCookie = decodeURIComponent(document.cookie);
        var ca = decodedCookie.split(';');
        for(var i = 0; i <ca.length; i++) {
            var c = ca[i];
            while (c.charAt(0) == ' ') {
                c = c.substring(1);
            }
            if (c.indexOf(name) == 0) {
                return c.substring(name.length, c.length);
            }
        }
        return "";
    }


    // from http://stackoverflow.com/questions/901115/how-can-i-get-query-string-values-in-javascript
    function getParameterByName(name, url) {
        if (!url) { url = window.location.href; }
        name = name.replace(/[\[\]]/g, "\\$&");
        var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
          results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, " "));
    }
}())