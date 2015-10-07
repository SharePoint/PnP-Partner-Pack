(function (window, undefined) {

    "use strict";

    var $ = window.jQuery;
    var document = window.document;

    // SPHostUrl parameter name
    var SPHostUrlKey = "SPHostUrl";

    // SPAppWebUrl parameter name
    var SPAppWebUrlKey = "SPAppWebUrl";

    // Gets SPHostUrl from the current URL and appends it as query string to the links which point to current domain in the page.
    $(document).ready(function () {
        ensureSPHasRedirectedToSharePointRemoved();
        
        var spHostUrl = getSPItemFromQueryString(window.location.search, SPHostUrlKey);
        var spAppWebUrl = getSPItemFromQueryString(window.location.search, SPAppWebUrlKey);
        var currentAuthority = getAuthorityFromUrl(window.location.href).toUpperCase();

        $("#mainChrome").hide();

        if (spHostUrl && currentAuthority) {
            appendSPHostUrlToLinks(spHostUrl, spAppWebUrl, currentAuthority);

            // The SharePoint js files URL are in the form:
            // web_url/_layouts/15/resource
            var scriptbase = decodeURIComponent(spHostUrl) + "/_layouts/15/";

            // Load the js file and continue to the 
            //   success handler
            $.getScript(scriptbase + "SP.UI.Controls.js", renderChrome);
        }
    });

    // Appends SPHostUrl as query string to all the links which point to current domain.
    function appendSPHostUrlToLinks(spHostUrl, spAppWebUrl, currentAuthority) {
        $("a")
            .filter(function () {
                var authority = getAuthorityFromUrl(this.href);
                if (!authority && /^#|:/.test(this.href)) {
                    // Filters out anchors and urls with other unsupported protocols.
                    return false;
                }
                return authority.toUpperCase() == currentAuthority;
            })
            .each(function () {
                if (!getSPItemFromQueryString(this.search, SPHostUrlKey)) {
                    if (this.search.length > 0) {
                        this.search += "&" + SPHostUrlKey + "=" + spHostUrl + "&SPAppWebUrl=" + spAppWebUrl;
                    }
                    else {
                        this.search = "?" + SPHostUrlKey + "=" + spHostUrl + "&SPAppWebUrl=" + spAppWebUrl;
                    }
                }
            });
        $("form")
            .filter(function () {
                var authority = getAuthorityFromUrl(this.action);
                if (!authority && /^#|:/.test(this.action)) {

                    return false;
                }
                if (authority == null)
                    return false;
                return authority.toUpperCase() == currentAuthority;
            })
            .each(function () {
                if (this.action.indexOf("?") == -1) {
                    this.action += "?" + SPHostUrlKey + "=" + spHostUrl + "&SPAppWebUrl=" + spAppWebUrl;
                } else {
                    if (this.action.indexOf(SPHostUrlKey) == -1) {
                        this.action += "&" + SPHostUrlKey + "=" + spHostUrl + "&SPAppWebUrl=" + spAppWebUrl;
                    }
                }
            });
    }

    // Gets SP* item from the given query string.
    function getSPItemFromQueryString(queryString, key) {
        if (queryString) {
            if (queryString[0] === "?") {
                queryString = queryString.substring(1);
            }

            var keyValuePairArray = queryString.split("&");

            for (var i = 0; i < keyValuePairArray.length; i++) {
                var currentKeyValuePair = keyValuePairArray[i].split("=");

                if (currentKeyValuePair.length > 1 && currentKeyValuePair[0] == key) {
                    return currentKeyValuePair[1];
                }
            }
        }

        return null;
    }

    // Gets authority from the given url when it is an absolute url with http/https protocol or a protocol relative url.
    function getAuthorityFromUrl(url) {
        if (url) {
            var match = /^(?:https:\/\/|http:\/\/|\/\/)([^\/\?#]+)(?:\/|#|$|\?)/i.exec(url);
            if (match) {
                return match[1];
            }
        }
        return null;
    }

    // If SPHasRedirectedToSharePoint exists in the query string, remove it.
    // Hence, when user bookmarks the url, SPHasRedirectedToSharePoint will not be included.
    // Note that modifying window.location.search will cause an additional request to server.
    function ensureSPHasRedirectedToSharePointRemoved() {
        var SPHasRedirectedToSharePointParam = "&SPHasRedirectedToSharePoint=1";

        var queryString = window.location.search;

        if (queryString.indexOf(SPHasRedirectedToSharePointParam) >= 0) {
            window.location.search = queryString.replace(SPHasRedirectedToSharePointParam, "");
        }
    }

    //Function to prepare the options and render the control
    function renderChrome() {
        // The Help, Account and Contact pages receive the 
        //   same query string parameters as the main page
        var options = {
            "appIconUrl": "/AppIcon.png",
            "appTitle": "PnP Partner Pack - Site Provisioning",
            "appHelpPageUrl": "https://github.com/OfficeDev/PnP-Partner-Pack/",
            // The onCssLoaded event allows you to 
            // specify a callback to execute when the
            // chrome resources have been loaded.
            "onCssLoaded": "chromeLoaded()",
            "settingsLinks": [
                {
                    "linkUrl": "http://aka.ms/OfficeDevPnP",
                    "displayName": "OfficeDev PnP"
                }
            ],
        };

        var nav = new SP.UI.Controls.Navigation(
                                "chrome_placeholder",
                                options
                            );
        nav.setVisible(true);
        // Fixed scrolling with Chrome control
        $("body").css("overflow", "auto");
        $("#chromeLoader").fadeOut();
        $("#mainChrome").delay(500).fadeIn(400, function () {
            if (typeof (onChromeLoaded) != "undefined") {
                onChromeLoaded();
            }
        });
    }

})(window);

// Callback for the onCssLoaded event defined
//  in the options object of the chrome control
function chromeLoaded() {
    // When the page has loaded the required
    //  resources for the chrome control,
    //  display the page body.
    $("body").show();
}