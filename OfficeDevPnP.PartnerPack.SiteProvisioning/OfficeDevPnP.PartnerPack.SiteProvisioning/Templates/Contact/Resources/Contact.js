/*global window, document, $, _spPageContextInfo */
$(document).ready(function () {
    "use strict";
    var PnPContact,
        ContactSearchResults;
    PnPContact = {
        items: [],
        FIRSTNAME_ID: "txtfirstName",
        LASTNAME_ID: "txtlastName",
        MOBILENUMBER_ID: "txtmobileNumber",
        CONTACT_LIST_NAME: "Contacts%20List",
        TOPN_RECORDS_DISPLAY: 3000,

        ContactSearch: function () {
            var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl + "/";
            ContactSearchResults.init($('#searchResults'), siteURL);
            ContactSearchResults.GetContactsFromSPList();
        },
        FilterContact: function (collection, valueToCompare, propertyName) {
            var searchResultItem = "";
            if (collection !== null && collection.length > 0) {
                if (propertyName === PnPContact.FIRSTNAME_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.FirstName === valueToCompare;
                    });
                } else if (propertyName === PnPContact.LASTNAME_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.Title === valueToCompare;
                    });
                } else if (propertyName === PnPContact.MOBILENUMBER_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.CellPhone === valueToCompare;
                    });
                }
            }
            return searchResultItem;
        },
        FilterContactByRegex: function (collection, valueToCompare, propertyName) {
            var regex = new RegExp("^" + valueToCompare, "ig");
            var searchResultItem = "";
            if (collection !== null && collection.length > 0) {
                if (propertyName === PnPContact.FIRSTNAME_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.FirstName.match(regex);
                    });
                } else if (propertyName === PnPContact.LASTNAME_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.Title.match(regex);
                    });
                } else if (propertyName === PnPContact.MOBILENUMBER_ID) {
                    searchResultItem = collection.filter(function (el) {
                        return el.CellPhone.match(regex);
                    });
                }
            }
            return searchResultItem;
        },
        ReturnAutoCompleteResponse: function (response, searchresultcollection, propertyName) {
            if (searchresultcollection !== null && searchresultcollection.length > 0) {
                if (propertyName === PnPContact.FIRSTNAME_ID) {
                    response($.map(searchresultcollection, function (item) {
                        return {
                            label: item.FirstName,
                            value: item.FirstName
                        };
                    }));
                } else if (propertyName === PnPContact.LASTNAME_ID) {
                    response($.map(searchresultcollection, function (item) {
                        return {
                            label: item.Title,
                            value: item.Title
                        };
                    }));
                } else if (propertyName === PnPContact.MOBILENUMBER_ID) {
                    response($.map(searchresultcollection, function (item) {
                        return {
                            label: item.CellPhone,
                            value: item.CellPhone
                        };
                    }));
                }
            } else {
                response([{label: 'No results found.', val: -1}]);
            }
        },
        RenderDiv: function (resultitem) {
            var html = "<div class='result-row' style='padding-bottom:5px; border-bottom: 1px solid #c0c0c0;'>";
            var clickableLink = "<span class='bold-class'>Full Name:</span> <a href='#'>" + resultitem.FirstName + " " + resultitem.Title + "</a><br/><span class='bold-class'>Mobile:</span><span> " + resultitem.CellPhone + "</span><br/><span class='bold-class'>Email:</span><span> " + resultitem.EMail + "</span><br/>";
            html += clickableLink;
            html += "</div>";
            return html;
        },
        GetValueById: function (elementId) {
            return $("#" + elementId).val();
        },
        ClearTextbox: function (elementId) {
            $('#' + elementId).val("");
        },
        RenderContacts: function () {
            var html = "<div class='results'>";
            $.each(PnPContact.items, function (key) {
                var value = PnPContact.items[key];
                html += PnPContact.RenderDiv(value);
            });
            html += "</div>";
            $("#searchResults").html(html);
        }
    };

    ContactSearchResults = {
        element: '',
        url: '',
        init: function (searchTerm, projectUrl) {
            ContactSearchResults.element = searchTerm;
            ContactSearchResults.url = projectUrl;
        },
        GetContactsFromSPList: function () {
            $.ajax({
                url: ContactSearchResults.url + "/_api/lists/getbytitle('" + PnPContact.CONTACT_LIST_NAME + "')/items?$select=FirstName,Title,CellPhone,EMail&$top=" + PnPContact.TOPN_RECORDS_DISPLAY + "&$orderby=FirstName asc",
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                },
                success: ContactSearchResults.onSuccess,
                error: ContactSearchResults.onError
            });
        },
        onSuccess: function (data) {
            var results = data.d.results;
            var html = "<div class='results'>";
            if (results.length === 0) {
                var clickableLink = "<div class='result-row' style='padding-bottom:5px; border-bottom: 1px solid #c0c0c0;'><span class='bold-class'>No Contacts Available</span><br/>";
                html += clickableLink;
                html += "</div>";
            } else {
                $.each(results, function (key) {
                    var value = results[key];
                    html += PnPContact.RenderDiv(value);
                    var item = {
                        FirstName: value.FirstName,
                        Title: value.Title,
                        CellPhone: value.CellPhone,
                        EMail: value.EMail
                    };
                    PnPContact.items.push(item);
                });
            }

            html += "</div>";
            $("#searchResults").html(html);
        },
        onError: function (err) {
            $("#searchResults").html("<h3>An error occured</h3><br/>" + JSON.stringify(err));
        }
    };

    PnPContact.ContactSearch();
    var searchtextboxId;

    $('input[type="text"]').on('change keyup', function () {
        var value = $(this).val();
        if (value === "") {
            PnPContact.RenderContacts();
        }
    });

    $('input[type="text"]').autocomplete({
        source: function (request, response) {
            searchtextboxId = $(this.element).prop("id");
            var searchresultcollection = [];
            if (searchtextboxId === PnPContact.FIRSTNAME_ID) {
                PnPContact.ClearTextbox(PnPContact.LASTNAME_ID);
                PnPContact.ClearTextbox(PnPContact.MOBILENUMBER_ID);
                searchresultcollection = PnPContact.FilterContactByRegex(PnPContact.items, PnPContact.GetValueById(PnPContact.FIRSTNAME_ID), PnPContact.FIRSTNAME_ID);
                PnPContact.ReturnAutoCompleteResponse(response, searchresultcollection, PnPContact.FIRSTNAME_ID);
            } else if (searchtextboxId === PnPContact.LASTNAME_ID) {
                PnPContact.ClearTextbox(PnPContact.FIRSTNAME_ID);
                PnPContact.ClearTextbox(PnPContact.MOBILENUMBER_ID);
                searchresultcollection = PnPContact.FilterContactByRegex(PnPContact.items, PnPContact.GetValueById(PnPContact.LASTNAME_ID), PnPContact.LASTNAME_ID);
                PnPContact.ReturnAutoCompleteResponse(response, searchresultcollection, PnPContact.LASTNAME_ID);
            } else if (searchtextboxId === PnPContact.MOBILENUMBER_ID) {
                PnPContact.ClearTextbox(PnPContact.FIRSTNAME_ID);
                PnPContact.ClearTextbox(PnPContact.LASTNAME_ID);
                searchresultcollection = PnPContact.FilterContactByRegex(PnPContact.items, PnPContact.GetValueById(PnPContact.MOBILENUMBER_ID), PnPContact.MOBILENUMBER_ID);
                PnPContact.ReturnAutoCompleteResponse(response, searchresultcollection, PnPContact.MOBILENUMBER_ID);
            }
        },
        select: function (event, ui) {
            if (ui.item) {
                var searchitem;
                if (searchtextboxId === PnPContact.FIRSTNAME_ID) {
                    searchitem = PnPContact.FilterContact(PnPContact.items, ui.item.value, PnPContact.FIRSTNAME_ID);
                } else if (searchtextboxId === PnPContact.LASTNAME_ID) {
                    searchitem = PnPContact.FilterContact(PnPContact.items, ui.item.value, PnPContact.LASTNAME_ID);
                } else if (searchtextboxId === PnPContact.MOBILENUMBER_ID) {
                    searchitem = PnPContact.FilterContact(PnPContact.items, ui.item.value, PnPContact.MOBILENUMBER_ID);
                }
                if (searchitem !== null && searchitem.length > 0) {
                    var resultitem = searchitem[0];
                    var htmlText = "<div class='results'>";
                    htmlText = htmlText + PnPContact.RenderDiv(resultitem) + "</div>";
                    $("#searchResults").html(htmlText);
                }

            }
        }
    });

});
