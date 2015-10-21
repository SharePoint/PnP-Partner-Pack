$(document).ready(function () {
    // Check to make sure the NavBar plugin is available, then run it
    if ($.fn.NavBar) {
        $('.ms-NavBar').NavBar();
    }

    // Handle selection of current NavBar item
    $("li.ms-NavBar-item > a").each(function (index) {
        if (window.location.href.indexOf($(this).attr("href")) > 0) {
            $(this).parent().addClass("is-selected");
        }
        else {
            $(this).parent().removeClass("is-selected");
        }
    });
});
