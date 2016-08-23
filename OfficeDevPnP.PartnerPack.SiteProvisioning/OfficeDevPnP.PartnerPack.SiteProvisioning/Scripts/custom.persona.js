$().ready(function () {
    $(".ms-PeoplePicker-searchField").on("keyup" , function (e) {
        if (e) {
            e.preventDefault();
            var keycode = (event.keyCode ? event.keyCode : event.which);

            SearchPeopleOrGroups();
        }

        e.preventDefault();
        return false;

    });
    
    $(".js-selectedRemove").click(function (e) {
        removePersona(e);
    });

    FixPersonaImages();
});

function removePersona(e) {
    // Remove the selected person parent li
    $(e.currentTarget).parent("li").remove();
    e.preventDefault();
    return false;
}

var searchXhr;
function SearchPeopleOrGroups() {
    // 4 stands for DONE
    if (searchXhr && searchXhr.readyState != 4) {
        searchXhr.abort();
    }

    var searchText = $(".ms-PeoplePicker-searchField").val();    
    var d = { "searchText": searchText };

    if (searchText && searchText != "") {
        searchXhr = $.ajax({
            url: '/Persona/SearchPeopleOrGroups',
            type: 'GET',
            data: d,
            success: function (result) {
                $(".ms-PeoplePicker-peopleList").empty();
                if (result) {
                    $.each(result, function (i, e) {
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
                        $(secondaryText).text(e.JobTitle);
                        $(detailsDiv).append(displayName);
                        $(detailsDiv).append(secondaryText);

                        // Add image area and details area to the parent
                        $(selectableDiv).append(imageDiv);
                        $(selectableDiv).append(detailsDiv);

                        $(div).append(selectableDiv);
                        $(li).append(div);


                        $(li).click(function (e) {
                            // TODO: copy the selected item to the $(".ms-PeoplePicker-selectedPeople")
                            var selectedLi = $("<li></li>").addClass("ms-PeoplePicker-selectedPerson");
                            var selectedDiv = $("<div></div>").addClass("ms-Persona").addClass("ms-Persona--sm");
                            var removeButton = $("<button></button>").addClass("ms-PeoplePicker-resultAction").addClass("js-selectedRemove");
                            $(removeButton).append($("<i></i>").addClass("ms-Icon").addClass("ms-Icon--x"));

                            $(removeButton).click(function (e) {
                                removePersona(e);
                            });

                            $(selectedDiv).append(imageDiv);
                            $(selectedDiv).append(detailsDiv);

                            $(selectedLi).append(selectedDiv);
                            $(selectedLi).append(removeButton);

                            $("ul[class = ms-PeoplePicker-selectedPeople]").append(selectedLi);
                            // Select the parent li if not already selected
                            if ($(e.currentTarget).parent("li").length == 0) {
                                $(e.currentTarget).remove();
                            }
                            else {
                                $(e.currentTarget).parent("li").remove();
                            }
                        });

                        $(".ms-PeoplePicker-peopleList").append(li);
                    });

                    $(".ms-PeoplePicker-results").show();
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