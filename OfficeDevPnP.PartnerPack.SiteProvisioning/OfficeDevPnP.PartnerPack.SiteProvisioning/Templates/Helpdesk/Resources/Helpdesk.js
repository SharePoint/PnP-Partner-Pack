/*global window, document, console, $, _spPageContextInfo*/
$(document).ready(function () {
    "use strict";
    var PnPHelpDesk,
        HelpdeskTicketResults;
    PnPHelpDesk = {
        editDialog: '',
        viewDialog: '',
        HELPDESK_LIST_NAME: "Helpdesk",
        EDIT_ATTACHMENT: "editattachment",
        SiteURL: window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl + "/",
        displayTicketsList: [],

        BindDropDownByIdListName: function (id, listName) {
            $.ajax({
                url: PnPHelpDesk.SiteURL + "_api/web/lists/getbytitle('" + listName + "')/items",
                type: "GET",
                contentType: "application/json;odata=verbose",
                headers: {
                    "Accept": "application/json;odata=verbose"
                },
                success: function (data) {
                    var results = data.d.results;
                    $.each(results, function (key) {
                        var value = results[key];
                        $("#" + id).append($("<option></option>").val(value.ID).html(value.Title));
                    });
                },
                error: function (data) {
                    console.log(data);
                    console.log('Operataion failed');
                }
            });
        },
        GetItemTypeForListName: function (name) {
            return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
        },
        GenerateTicket: function () {
            var num = Math.floor(Math.random() * 900000) + 100000;
            return num;
        },
        CheckValidations: function () {
            var subjectlength = $("#subject").val().length;
            var descriptionlength = $("#ticket_description").val().length;
            var workstationlength = $("#workstation").val().length;

            if (subjectlength === 0 || descriptionlength === 0 || workstationlength === 0) {
                return false;
            } else {
                return true;
            }
        },
        SaveNewTicketInfo: function () {
            var ticketSubject = $("#subject").val();
            var ticketTypeID = $("#ticket_type").val();
            var ticketTargetDeptID = $("#ticket_to").val();
            var ticketDescription = $("#ticket_description").val();
            var ticketPriorityID = $("#priority").val();
            var ticketWorkstation = $("#workstation").val();
            var randomticketID = PnPHelpDesk.GenerateTicket();

            var itemType = PnPHelpDesk.GetItemTypeForListName(PnPHelpDesk.HELPDESK_LIST_NAME);
            var userid = _spPageContextInfo.userId;
            var item = {
                "__metadata": { "type": itemType },
                "Title": ticketSubject,
                "Subject": ticketSubject,
                "Description": ticketDescription,
                "Workstation": ticketWorkstation,
                "TicketNumber": String(randomticketID),
                "TicketTypeId": ticketTypeID,
                "TicketToId": ticketTargetDeptID,
                "PriorityId": ticketPriorityID,
                "StatusId": 1,
                "TicketFromId": userid
            };
            $.ajax({
                url: PnPHelpDesk.SiteURL + "_api/web/lists/getbytitle('" + PnPHelpDesk.HELPDESK_LIST_NAME + "')/items",
                type: "POST",
                contentType: "application/json;odata=verbose",
                data: JSON.stringify(item),
                async: false,
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val()
                },
                success: function (data) {
                    PnPHelpDesk.UploadFile(randomticketID, PnPHelpDesk.HELPDESK_LIST_NAME, PnPHelpDesk.SiteURL, "attachment");
                    var ticketFieldsArray = [];
                    ticketFieldsArray.push(item.TicketNumber);
                    ticketFieldsArray.push(item.Subject);
                    ticketFieldsArray.push('New');
                    ticketFieldsArray.push(item.TicketNumber);
                    PnPHelpDesk.displayTicketsList.push(ticketFieldsArray);
                    console.log('Item is added successfully');
                },
                error: function (data) {
                    console.log('Operataion failed');
                }
            });
        },
        UploadFile: function (randomticketID, listName, siteURL, fileControlId) {
            var ticketObject = PnPHelpDesk.GetTicketByID(randomticketID);
            var listitemid = ticketObject.ID;
            var file = $("#" + fileControlId)[0].files[0];
            if (file != null && file.name.length > 0) {
                var fileUploadurl = siteURL + "_api/web/lists/getByTitle(@TargetLibrary)/Items(" + listitemid + ")/AttachmentFiles/add(FileName=@TargetFileName)?@TargetLibrary='" + listName + "'" + "&@TargetFileName='" + file.name + "'";
                var metaData = 'This is the content of the file added through REST API';
                $.ajax({
                    url: fileUploadurl,
                    async: false,
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: metaData,
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-RequestDigest": $("#__REQUESTDIGEST").val()
                    },
                    success: function (data) {
                        console.log('File uploaded successfully');
                    },
                    error: function (data) {
                        console.log("File upload failed. Please try again");
                    }
                });
            }
        },
        GetListItemWithId: function (itemId, listName, siteurl, success, failure) {
            var url = siteurl + "/_api/web/lists/getbytitle('" + listName + "')/items?$filter=Id eq " + itemId;
            $.ajax({
                url: url,
                method: "GET",
                headers: { "Accept": "application/json; odata=verbose" },
                success: function (data) {
                    if (data.d.results.length === 1) {
                        success(data.d.results[0]);
                    } else {
                        failure("Multiple results obtained for the specified Id value");
                    }
                },
                error: function (data) {
                    failure(data);
                }
            });
        },
        UpdateListItemDetails: function (itemId, listName, siteUrl, success, failure) {
            var isvalid = PnPHelpDesk.CheckDialgValidations();
            if (isvalid) {
                $("#div_dialg_error").hide();
                var itemType = PnPHelpDesk.GetItemTypeForListName(listName);
                // Prepare the ListItem
                var item = {
                    "__metadata": { "type": itemType },
                    "Subject": $("#subject_dialg").val(),
                    "Description": $("#ticket_dialg_description").val(),
                    "Workstation": $("#workstation_dialg").val(),
                    "TicketTypeId": $("#ticket_dialg_type").val(),
                    "TicketToId": $("#ticket_dialg_to").val(),
                    "PriorityId": $("#priority_dialg").val(),
                    "TicketNumber": $("#dialg_ticketNumber").val()
                };

                PnPHelpDesk.GetListItemWithId(itemId, listName, siteUrl,
                        function (data) {
                            $.ajax({
                                url: data.__metadata.uri,
                                type: "POST",
                                contentType: "application/json;odata=verbose",
                                data: JSON.stringify(item),
                                headers: {
                                    "Accept": "application/json;odata=verbose",
                                    "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                    "X-HTTP-Method": "MERGE",
                                    "If-Match": data.__metadata.etag
                                },
                                success: function (data) {
                                    success(data, item);
                                },
                                error: function (data) {
                                    failure(data);
                                }
                            });
                        },
                        function (data) {
                            failure(data);
                        });

            } else {
                $("#div_dialg_error").show();
            }
        },
        UpdateListItem: function () {
            var siteURL = window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl + "/";
            PnPHelpDesk.UpdateListItemDetails($("#dialg_ticketId").val(), PnPHelpDesk.HELPDESK_LIST_NAME, siteURL, function (data, item) {//success callback
                PnPHelpDesk.UploadFile($("#dialg_ticketNumber").val(), PnPHelpDesk.HELPDESK_LIST_NAME, siteURL, PnPHelpDesk.EDIT_ATTACHMENT);
                console.log("Item updated, refreshing avilable items");

                $.each(PnPHelpDesk.displayTicketsList, function (index, value) {
                    if (value[0] === item.TicketNumber) {
                        var temprecord = PnPHelpDesk.displayTicketsList[index];
                        temprecord[1] = item.Subject;
                        PnPHelpDesk.displayTicketsList[index] = temprecord;
                    }
                });
                PnPHelpDesk.DisplayMyTickets(PnPHelpDesk.displayTicketsList);
                PnPHelpDesk.editDialog.dialog("close");
            },
                    function () {//Failure callback
                        $("#" + PnPHelpDesk.EDIT_ATTACHMENT).val('');
                        console.log("Oops, an error occured. Please try again");
                    });
        },
        RetrieveTickets: function () {

            // Call our Init-function
            HelpdeskTicketResults.init($('#tblMyTickets'), PnPHelpDesk.SiteURL);

            // Call our Load-function which will post the actual query
            HelpdeskTicketResults.GetTicketsFromSPList();
        },
        DisplayMyTickets: function (displayTicketsList) {
            var displayTickets = displayTicketsList;
            $('#tblMyTickets').DataTable({
                destroy: true,
                "lengthMenu": [5, 10],
                data: displayTickets,
                columns: [
                    {
                        title: "Ticket#",
                        "render": function (data, type) {
                            if (type === 'display') {
                                return $('<a>')
                                    .attr('href', "#")
                                    .attr('id', data)
                                    .text(data)
                                    .wrap('<div></div>')
                                    .parent()
                                    .html();

                            } else {
                                return data;
                            }
                        }
                    },
                    { title: "Title" },
                    { title: "State" },
                    {
                        "title": "",
                        "render": function (data, type) {
                            if (type === 'display') {
                                return $('<a>')
                                    .attr('href', "#")
                                    .attr('id', data)
                                    .text("Edit")
                                    .wrap('<div></div>')
                                    .parent()
                                    .html();

                            } else {
                                return data;
                            }
                        }
                    }
                ]
            });
            $("#tblMyTickets").on("click", "a", function (event) {
                var ticketId = $(this).attr('id');
                var ticketObj = PnPHelpDesk.GetTicketByID(ticketId);
                var linkText = $(this).text();
                if (linkText !== "Edit") {
                    $("#view-ticket-type").html(ticketObj.TicketType.Title);
                    $("#view-ticket-To").html(ticketObj.TicketTo.Title);
                    $("#view-ticket-Subject").html(ticketObj.Subject);
                    $("#view-ticket-Description").html(ticketObj.Description);
                    $("#view-ticket-Priority").html(ticketObj.Priority.Title);
                    $("#view-ticket-Workstation").html(ticketObj.Workstation);
                    PnPHelpDesk.viewDialog.dialog("open");
                } else {
                    $("#ticket_dialg_type").val(ticketObj.TicketTypeId);
                    $("#ticket_dialg_to").val(ticketObj.TicketToId);
                    $("#subject_dialg").val(ticketObj.Subject);
                    $("#ticket_dialg_description").val(ticketObj.Description);
                    $("#priority_dialg").val(ticketObj.PriorityId);
                    $("#workstation_dialg").val(ticketObj.Workstation);
                    $("#dialg_ticketNumber").val(ticketObj.TicketNumber);
                    $("#dialg_ticketId").val(ticketObj.ID);
                    if (ticketObj !== null && ticketObj.Attachments === true) {
                        $("#" + PnPHelpDesk.EDIT_ATTACHMENT).attr('disabled', true);
                    } else {
                        $("#" + PnPHelpDesk.EDIT_ATTACHMENT).attr('disabled', false);
                        $("#" + PnPHelpDesk.EDIT_ATTACHMENT).val('');
                    }
                    PnPHelpDesk.editDialog.dialog("open");
                }
            });
        },
        CheckDialgValidations: function () {
            var subjectlength = $("#subject_dialg").val().length;
            var descriptionlength = $("#ticket_dialg_description").val().length;
            var workstationlength = $("#workstation_dialg").val().length;

            if (subjectlength === 0 || descriptionlength === 0 || workstationlength === 0) {
                return false;
            } else {
                return true;
            }
        },
        GetTicketByID: function (ticketNumber) {
            var ticketObject = "";
            var ticketurl = PnPHelpDesk.SiteURL + "_api/lists/getbytitle('" + PnPHelpDesk.HELPDESK_LIST_NAME + "')/items?$select=ID,TicketType/Title,TicketTo/Title,Subject,Description,PriorityId,TicketToId,TicketTypeId,Priority/Title,Workstation,TicketNumber,StatusId,Title,Attachments&$expand=TicketType,Priority,TicketTo&$filter=TicketNumber eq " + ticketNumber;
            $.ajax(
                {
                    url: ticketurl,
                    async: false,
                    method: "GET",
                    headers:
                            {
                                "accept": "application/json;odata=verbose"
                            },
                    success: function TicketonSuccess(data) {
                        ticketObject = data.d.results[0];
                    },
                    error: function onError(error) {
                        console.log('Error occurred while retrieving the ticket');
                    }
                }
            );
            return ticketObject;
        },
        FormReset: function () {
            $("#subject").val('');
            $("#ticket_description").val('');
            $("#attachment").val('');
            $("#workstation").val('');
            $("#ticket_type").val("1");
            $("#ticket_to").val("1");
            $("#priority").val("1");
            return false;
        }
    };

    HelpdeskTicketResults = {
        element: '',
        url: '',
        init: function (ticketsElement, siteURL) {
            HelpdeskTicketResults.element = ticketsElement;
            HelpdeskTicketResults.url = siteURL + "_api/lists/getbytitle('" + PnPHelpDesk.HELPDESK_LIST_NAME + "')/items?$select=TicketNumber,Subject,Status/Title&$expand=Status";
        },
        GetTicketsFromSPList: function () {
            $.ajax(
                {
                    url: HelpdeskTicketResults.url,
                    method: "GET",
                    headers:
                            {
                                "accept": "application/json;odata=verbose"
                            },
                    success: HelpdeskTicketResults.onSuccess,
                    error: HelpdeskTicketResults.onError
                }
            );
        },
        onSuccess: function (data) {
            var results = data.d.results;
            $.each(results, function (key) {
                var value = results[key];
                var ticketFieldsArray = [];
                ticketFieldsArray.push(value.TicketNumber);
                ticketFieldsArray.push(value.Subject);
                ticketFieldsArray.push(value.Status.Title);
                ticketFieldsArray.push(value.TicketNumber);
                PnPHelpDesk.displayTicketsList.push(ticketFieldsArray);
            });
            PnPHelpDesk.DisplayMyTickets(PnPHelpDesk.displayTicketsList);
        },
        onError: function (err) {
            $("#tblMyTickets tbody").html("<h3>An error occured</h3><br/>" + JSON.stringify(err));
        }
    };
    // Load the Drop Downs
    PnPHelpDesk.BindDropDownByIdListName("ticket_type", "TicketType");
    PnPHelpDesk.BindDropDownByIdListName("ticket_to", "TicketTo");
    PnPHelpDesk.BindDropDownByIdListName("priority", "Priority");
    PnPHelpDesk.BindDropDownByIdListName("ticket_dialg_type", "TicketType");
    PnPHelpDesk.BindDropDownByIdListName("ticket_dialg_to", "TicketTo");
    PnPHelpDesk.BindDropDownByIdListName("priority_dialg", "Priority");

    $("#create-ticket").click(function () {
        var isvalid = PnPHelpDesk.CheckValidations();
        if (isvalid) {
            $("#div_error").hide();
            PnPHelpDesk.SaveNewTicketInfo();
            PnPHelpDesk.DisplayMyTickets(PnPHelpDesk.displayTicketsList);
            PnPHelpDesk.FormReset();
            return false;
        } else {
            $("#div_error").show();
            return false;
        }
    });

    PnPHelpDesk.RetrieveTickets();

    $("#reset-ticket").click(function () {
        PnPHelpDesk.FormReset();
        return false;
    });

    PnPHelpDesk.editDialog = $("#dialog-form").dialog({
        autoOpen: false,
        height: 600,
        width: 425,
        modal: true,
        buttons: {
            "Save": PnPHelpDesk.UpdateListItem,
            Cancel: function () {
                PnPHelpDesk.editDialog.dialog("close");
            }
        },
        close: function () {
            PnPHelpDesk.editDialog.dialog("close");
        }
    });

    PnPHelpDesk.viewDialog = $("#dialog-view").dialog({
        autoOpen: false,
        height: 400,
        width: 425,
        modal: true,
        buttons: {
            Close: function () {
                PnPHelpDesk.viewDialog.dialog("close");
            }
        },
        close: function () {
            PnPHelpDesk.viewDialog.dialog("close");
        }
    });

    //Code to hide validation message while loading page
    $("#div_error").hide();
    $("#div_dialg_error").hide();
});
