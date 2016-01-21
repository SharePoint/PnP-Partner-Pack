
// variable used for cross site CSOM calls
var context;
// peoplePicker variable needs to be globally scoped as the generated html contains JS that will call into functions of this class
var peoplePicker;
var businessOwnerPrimaryPicker;
var businessOwnerSecondaryPicker;

//Wait for the page to load
$(document).ready(function () {

    //Get the URI decoded SharePoint site url from the SPHostUrl parameter.
    var spHostUrl = window.location.protocol + "//" + window.location.host;
    var appWebUrl = spHostUrl + _spPageContextInfo.siteServerRelativeUrl + "/";
    var spLanguage = 'en-US';

    //Build absolute path to the layouts root with the spHostUrl
    var layoutsRoot = spHostUrl + '/_layouts/15/';

    //load all appropriate scripts for the page to function
    $.getScript(layoutsRoot + 'SP.Runtime.js',
        function () {
            $.getScript(layoutsRoot + 'SP.js',
                function () {
                    //Load the SP.UI.Controls.js file to render the App Chrome
                    //$.getScript(layoutsRoot + 'SP.UI.Controls.js', renderSPChrome);

                    //load scripts for cross site calls (needed to use the people picker control in an IFrame)
                    $.getScript(layoutsRoot + 'SP.RequestExecutor.js', function () {
                        context = new SP.ClientContext(appWebUrl);
                        var factory = new SP.ProxyWebRequestExecutorFactory(appWebUrl);
                        context.set_webRequestExecutorFactory(factory);
                        businessOwnerPrimaryPicker = getPeoplePickerInstance(context, $('#spanAdministrators'), $('#inputAdministrators'), $('#divAdministratorsSearch'), $('#hdnAdministrators'), "businessOwnerPrimaryPicker", spLanguage);
                        businessOwnerSecondaryPicker = getPeoplePickerInstance(context, $('#spanAdministrators1'), $('#inputAdministrators1'), $('#divAdministratorsSearch1'), $('#hdnAdministrators1'), "businessOwnerSecondaryPicker", spLanguage);
                    });
                });
        });
});

function getPeoplePickerInstance(context, spanControl, inputControl, searchDivControl, hiddenControl, variableName, spLanguage) {
    var newPicker;
    newPicker = new CAMControl.PeoplePicker(context, spanControl, inputControl, searchDivControl, hiddenControl);
    // required to pass the variable name here!
    newPicker.InstanceName = variableName;
    // Pass current language, if not set defaults to en-US. Use the SPLanguage query string param or provide a string like "nl-BE"
    // Do not set the Language property if you do not have foreseen javascript resource file for your language
    newPicker.Language = spLanguage;
    // optionally show more/less entries in the people picker dropdown, 4 is the default
    newPicker.MaxEntriesShown = 5;
    // Can duplicate entries be selected (default = false)
    newPicker.AllowDuplicates = false;
    // Show the user loginname
    newPicker.ShowLoginName = true;
    // Show the user title
    newPicker.ShowTitle = true;
    // Set principal type to determine what is shown (default = 1, only users are resolved).
    // See http://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.client.utilities.principaltype.aspx for more details
    // Set ShowLoginName and ShowTitle to false if you're resolving groups
    newPicker.PrincipalType = 1;
    // start user resolving as of 2 entered characters (= default)
    newPicker.MinimalCharactersBeforeSearching = 2;

    //Adding maxusers
    newPicker.MaxUsers = 1;

    // Hookup everything
    newPicker.Initialize();

    return newPicker;
}