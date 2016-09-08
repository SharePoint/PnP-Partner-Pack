var delayKeyup = (function () {
    var timer = 0;
    return function (callback, ms) {
        clearTimeout(timer);
        timer = setTimeout(callback, ms);
    };
})();

function removePersona(e, searchId) {
    // Remove the selected person parent li
    $(e.currentTarget).parent("li").remove();

    var ul = $("#" + searchId + " ul[class = ms-PeoplePicker-selectedPeople]");
    var count = $("li", ul).length;
    $("#" + searchId + " .ms-PeoplePicker-selectedCount").text(count);

    e.preventDefault();
    return false;
}

var searchXhr;
function searchPeopleOrGroups(searchId, searchName, searchGroups, maxSelection) {
    // 4 stands for DONE
    if (searchXhr && searchXhr.readyState != 4) {
        searchXhr.abort();
    }

    var searchText = $("#" + searchId + " .ms-PeoplePicker-searchField").val();
    var d = { "searchText": searchText, "searchGroups": searchGroups };

    if (searchText && searchText != "") {
        searchXhr = $.ajax({
            url: '/Persona/SearchPeopleOrGroups',
            type: 'GET',
            data: d,
            success: function (result) {
                $("#" + searchId + " .ms-PeoplePicker-peopleList").empty();
                if (result) {
                    $.each(result.Principals, function (i, e) {
                        // Create a new li element to display the search result
                        // Root
                        var li = $("<li></li>").addClass("ms-PeoplePicker-peopleListItem").addClass("ms-PeoplePicker-result");
                        var div = $("<div></div>").addClass("ms-PeoplePicker-peopleListBtn").attr("tabindex", 0).attr("role", "button");
                        var selectableDiv = $("<div></div>").addClass("ms-Persona").addClass("ms-Persona--selectable").addClass("ms-Persona--sm");

                        // Image area
                        var imageDiv = $("<div></div>").addClass("ms-Persona-imageArea");
                        var initialsDiv = $("<div></div>").addClass("ms-Persona-initials").addClass(e.BadgeColor);
                        $(initialsDiv).text(e.Abbreviation);
                        $(imageDiv).append(initialsDiv);
                        if (e.Mail != undefined && e.Mail != "") {
                            var photoImg = $("<img src='/Persona/GetPhoto?upn=" + e.Mail + "&height=64&width=64'></img>").addClass("ms-Persona-image");
                            $(imageDiv).append(photoImg);

                            $(photoImg).error(function () {
                                $(this).hide();
                            });
                        }

                        // Details area
                        var detailsDiv = $("<div></div>").addClass("ms-Persona-details");
                        var displayName = $("<div></div>").addClass("ms-Persona-primaryText");
                        if (e.FirstName != null && e.LastName != null) {
                            $(displayName).text(e.FirstName + " " + e.LastName);
                        }
                        else {
                            $(displayName).text(e.DisplayName);
                        }
                        var secondaryText = $("<div></div>").addClass("ms-Persona-secondaryText");
                        $(secondaryText).text(e.UserPrincipalName);
                        $(detailsDiv).append(displayName);
                        $(detailsDiv).append(secondaryText);

                        // Add image area and details area to the parent
                        $(selectableDiv).append(imageDiv);
                        $(selectableDiv).append(detailsDiv);

                        $(div).append(selectableDiv);
                        $(li).append(div);

                        var mail = e.Mail;

                        $(li).click(function (e) {
                            var currentSelected = $("#" + searchId + " ul[class=ms-PeoplePicker-selectedPeople] > li").length;
                            if (currentSelected >= maxSelection) {
                                alert("You cannot select more profiles.");
                                return false;
                            }

                            var selectedLi = $("<li></li>").addClass("ms-PeoplePicker-selectedPerson");
                            var selectedDiv = $("<div></div>").addClass("ms-Persona").addClass("ms-Persona--sm");
                            var removeButton = $("<button></button>").addClass("ms-PeoplePicker-resultAction").addClass("js-selectedRemove");
                            $(removeButton).append($("<i></i>").addClass("ms-Icon").addClass("ms-Icon--x"));

                            $(removeButton).click(function (e) {
                                removePersona(e, searchId);
                            });

                            $(selectedDiv).append(imageDiv);
                            $(selectedDiv).append(detailsDiv);

                            $(selectedLi).append(selectedDiv);
                            $(selectedLi).append(removeButton);

                            var ul = $("#" + searchId + " ul[class = ms-PeoplePicker-selectedPeople]");
                            var count = $("li", ul).length;
                            selectedLi.append($("<input type=hidden name=" + searchName + ".Principals[" + count + "].Mail />").val(mail));

                            $("#" + searchId + " .ms-PeoplePicker-selectedCount").text(count + 1);

                            ul.append(selectedLi);
                            // Select the parent li if not already selected
                            if ($(e.currentTarget).parent("li").length == 0) {
                                $(e.currentTarget).remove();
                            }
                            else {
                                $(e.currentTarget).parent("li").remove();
                            }
                        });

                        $("#" + searchId + " .ms-PeoplePicker-peopleList").append(li);
                    });

                    $("#" + searchId + " .ms-PeoplePicker-results").show();
                }
            },
            error: function (e) {
                if (e.statusText != "abort") {
                    alert("error");
                }
            }
        });
        setTimeout(searchXhr, 500);
    }
}
