var SPOCustomBar = SPOCustomBar || {};

SPOCustomBar.addCustomButton = function (id, icon, title, onclick) {
    if ($("#" + id).length)
        return;

    var btnContainer = $($.parseHTML('<div id="' + id + '" class="o365cs-nav-topItem"></div>'));
    var btn = $($.parseHTML(
        '<button class="o365button o365cs-nav-item o365cs-nav-button o365cs-topnavText ms-bgc-tdr-h" type="button" aria-haspopup="true" role="menuitem" title="' + title + '">' +
            '<i class="ms-Icon ms-Icon--' + icon + '" role="presentation"></span>' +
        '</button>'));
    btn.click(onclick);
    btnContainer.append(btn);

    $("#SPOResponsiveBar > div").prepend(btnContainer);

    return btnContainer;
}
SPOCustomBar.addCustomLinkButton = function (id, icon, title, url) {
    return this.addCustomButton(id, icon, title, function () {
        window.location = url;
    });
}

SPOCustomBar.notifCount = 0;
SPOCustomBar.setNotificationCount = function (count) {
    SPOCustomBar.notifCount = count;
    $("#ResponsiveNotifCount").text(SPOCustomBar.notifCount);
}

SPOCustomBar.matchTopBarColor = function () {
    if ($("#O365_NavHeader").length < 1) {
        // we need the header bar to be fully loaded to match its color. Wait.
        setTimeout(this.matchTopBarColor, 50);
        return;
    }
    var notifBar = $("#SPOResponsiveBar > div");
    notifBar.css("background-color", $("#O365_NavHeader").css("background-color"));
}

SPOCustomBar.setUpCustomBar = function () {
    if ($("#SPOResponsiveBar").length)
        return;

    var notifBarContainer = $("<div>");
    notifBarContainer.attr("id", "SPOResponsiveBar");
    notifBarContainer.addClass("o365cs-nav-header16 o365cs-base o365cs-coremin");
    $("#s4-workspace").before(notifBarContainer);

    var notifBar = $("<div>");
    notifBar.addClass("o365cs-nav-rightAlign");
    notifBarContainer.append(notifBar);
    // this.matchTopBarColor();
}

SPOCustomBar.setUpNotifications = function () {
    if ($("#SPONotificationBtn").length)
        return;

    var notifBtnContainer = this.addCustomButton('SPONotificationBtn', 'globe', 'Notifications', function () {
        $(this).toggleClass("ms-bgc-w ms-fcl-b o365cs-spo-topbarMenuOpen");
        $(this).toggleClass("o365cs-topnavText ms-bgc-tdr-h");
        $(this).siblings("div.o365cs-nav-contextMenu").toggle();
    });
    notifBtnContainer.find("button").append($.parseHTML('<span id="ResponsiveNotifCount">' + this.notifCount + '</span>'));
    var notifDropdown = $($.parseHTML(
        '<div class="o365cs-nav-contextMenu o365spo contextMenuPopup"></div>'));
    notifBtnContainer.append(notifDropdown);
}

SPOCustomBar.addNotification = function (id, text, link, separatorBefore) {
    SPOCustomBar.setNotificationCount(this.notifCount + 1);

    var newNotificationHtml = '<a id="' + id + '" class="o365button o365cs-contextMenuItem ms-fcl-b" role="link" href="' + link + '" style="text-decoration: none;">' + text + '</a>';

    if (separatorBefore) {
        newNotificationHtml = '<div id="' + id + '_separator" class="o365cs-contextMenuSeparator ms-bcl-nl"></div>' +
            newNotificationHtml;
    }

    var newNotification = $($.parseHTML(newNotificationHtml));
    $("#SPONotificationBtn  > div.contextMenuPopup").append(newNotification);
}

SPOCustomBar.removeNotification = function (id) {
    if ($("#" + id).length > 0) {
        $("#" + id).remove();
        $("#" + id + "_separator").remove();

        SPOCustomBar.setNotificationCount(this.notifCount - 1);
    }
}


SPOCustomBar.setUpCustomFooter = function () {
    if ($("#SPOCustomFooter").length)
        return;

    var footerContainer = $("<div>");
    footerContainer.attr("id", "SPOCustomFooter");

    footerContainer.append("<ul>");

    $("#s4-workspace").append(footerContainer);
}

SPOCustomBar.addCustomFooterText = function (id, text) {
    if ($("#" + id).length)
        return;

    var customElement = $("<div>");
    customElement.attr("id", id);
    customElement.html(text);

    $("#SPOCustomFooter > ul").before(customElement);

    return customElement;
}
SPOCustomBar.addCustomFooterLink = function (id, text, url) {
    if ($("#" + id).length)
        return;

    var customElement = $("<a>");
    customElement.attr("id", id);
    customElement.attr("href", url);
    customElement.html(text);

    $("#SPOCustomFooter > ul").append($("<li>").append(customElement));

    return customElement;
}



SPOCustomBar.init = function (whenReadyDoFunc) {
    // avoid executing inside iframes (used by Sharepoint for dialogs)
    if (self !== top) return;

    if (!window.jQuery) {
        // jQuery is needed for Custom Bar to run
        setTimeout(function () { SPOCustomBar.init(whenReadyDoFunc); }, 50);
    } else {
        $(function () {
            SPOCustomBar.setUpCustomBar();
            SPOCustomBar.setUpCustomFooter();
            whenReadyDoFunc();
        });
    }
}



// The following initializes the custom bars, passing the customization to be executed when it is ready.
SPOCustomBar.init(function () {
    SPOCustomBar.setUpNotifications();
    SPOCustomBar.addNotification('SPOSampleNotification01', 'Add #2', 'javascript:SPOCustomBar.addNotification(\'SPOSampleNotification02\',\'Generic notification\', \'javascript:alert(\\\'Generic notification\\\');\');', false);
    SPOCustomBar.addNotification('SPOSampleNotification02', 'Generic notification', 'javascript:alert(\'Generic notification!\');', false);
    SPOCustomBar.addNotification('SPOSampleNotification03', 'Remove #2', 'javascript:SPOCustomBar.removeNotification(\'SPOSampleNotification02\');', true);

    SPOCustomBar.addCustomLinkButton('SPOSearchBtn', 'search', 'Search Center', 'SearchCenter.aspx');
    SPOCustomBar.addCustomLinkButton('SPOCRMBtn', 'person', 'CRM', 'CRM.aspx');

    SPOCustomBar.addCustomFooterText('SPOFooterCopyright', '&copy; 2015, Contoso Inc.');
    SPOCustomBar.addCustomFooterLink('SPOFooterCRMLink', 'CRM', 'CRM.aspx');
    SPOCustomBar.addCustomFooterLink('SPOFooterSearchLink', 'Search Center', 'SearchCenter.aspx');
    SPOCustomBar.addCustomFooterLink('SPOFooterPrivacyLink', 'Privacy Policy', 'Privacy.aspx');
});
