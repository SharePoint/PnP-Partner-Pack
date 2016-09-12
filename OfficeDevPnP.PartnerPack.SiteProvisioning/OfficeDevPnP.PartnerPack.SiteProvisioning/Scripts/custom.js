function fixPersonaImages() {
    $(".ms-Persona-image").error(function () {
        $(this).hide();
    });
}

function applyOfficeUIFabricStyles() {
    if ($.fn.NavBar) {
        $('.ms-NavBar').NavBar();
    }

    if ($.fn.Dropdown) {
        $(".ms-Dropdown").Dropdown();
    }

    if ($.fn.SearchBox) {
        $(".ms-SearchBox").SearchBox();
    }

    if ($.fn.DatePicker) {
        $(".ms-DatePicker").DatePicker();
    }

    if ($.fn.CommandBar) {
        $(".ms-CommandBar").CommandBar();
    }

    if ($.fn.Dialog) {
        $(".ms-Dialog").Dialog();
    }

    //if ($.fn.ContextualMenu) {
    //    $(".ms-ContextualMenu").ContextualMenu();
    //}

    if ($.fn.Facepile) {
        $(".ms-Facepile").Facepile();
    }

    if ($.fn.ListItem) {
        $(".ms-ListItem").ListItem();
    }

    if ($.fn.Panel) {
        $(".ms-Panel").Panel();
    }

    if ($.fn.PeoplePicker) {
        $(".ms-PeoplePicker").PeoplePicker();
    }

    if ($.fn.PersonaCard) {
        $(".ms-PersonaCard").PersonaCard();
    }

    if ($.fn.Pivot) {
        $(".ms-Pivot").Pivot();
    }

    if ($.fn.TextField) {
        $(".ms-TextField").TextField();
    }
}

$(document).ready(function () {

    // Make sure all the Office UI Fabric elements are properly rendered
    applyOfficeUIFabricStyles();

    // Handle selection of current NavBar item
    $("li.ms-NavBar-item > a").each(function (index) {
        if (window.location.href.indexOf($(this).attr("href")) > 0) {
            $(this).parent().addClass("is-selected");
        }
        else {
            $(this).parent().removeClass("is-selected");
        }
    });

    // Hide any unavailable persona image
    fixPersonaImages();
});
