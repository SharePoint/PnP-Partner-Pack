$(document).ready(function () {
    var PnPExpense;

    PnPExpense = {
        expenseItemDialog: '',
        EXPENSE_LIST_NAME: "Expense",
        EXPENSEITEMS_LIST_NAME: "ExpenseItems",
        EXPENSE_FINANCE_MANAGERWORKFLOW_LIST_NAME: "FinanceManagerWorkflow",
        SiteURL: window.location.protocol + "//" + window.location.host + _spPageContextInfo.siteServerRelativeUrl + "/",
        myExpenses: [],

        CheckExpenseListValidation: function () {
            var billTypeValid, billNumberValid, billDateValid, amountValid;

            billTypeValid = $("#bill_type").val();
            billNumberValid = $("#bill_no").val();
            billDateValid = $("#bill_date").val();
            amountValid = $("#bill_amount_rupees").val();

            if (jQuery.trim(billTypeValid).length == 0 || jQuery.trim(billNumberValid).length == 0 || jQuery.trim(billDateValid).length == 0 || jQuery.trim(amountValid).length == 0) {
                return false;
            }
            else {
                return true;
            }

        },

        CheckFormDataValidation: function () {
            var nameAsPerBankValid, empCodeValid, bankAccountNoValid, clientNameValid, projName, deptName, claimSubject;
            nameAsPerBankValid = $("#user_name_bank_account").val();
            empCodeValid = $("#emp_code").val();
            bankAccountNoValid = $("#bank_account_number").val();            
            clientNameValid = $("#client_name").val();
            projName = $("#project").val();
            deptName = $("#dept").val();
            claimSubject = $("#claim_subject").val();

            if (jQuery.trim(nameAsPerBankValid).length == 0 || jQuery.trim(empCodeValid).length == 0 || jQuery.trim(projName).length == 0 || jQuery.trim(deptName).length == 0 ||
                jQuery.trim(bankAccountNoValid).length == 0 || jQuery.trim(clientNameValid).length == 0 || $(".cam-entity-resolved").text().length <= 0 || jQuery.trim(claimSubject).length == 0) {
                return false;
            }
            else {
                return true;
            }
        },

        AddExpense: function () {
            var sno, billtype, billnumber, billdate, vendorname, particulars, amountinrupees;
            sno = $("#users tbody").children().length + 1;
            billtype = $("#bill_type");
            billnumber = $("#bill_no");
            billdate = $("#bill_date");
            vendorname = $("#vendor_name");
            particulars = $("#bill_particulars");
            amountinrupees = $("#bill_amount_rupees");

            var isExpenseValid = PnPExpense.CheckExpenseListValidation();

            if (isExpenseValid) {
                $("#div_error").hide();

                $("#users tbody").append("<tr>" +
                    "<td>" + sno + "</td>" +
                    "<td>" + billtype.val() + "</td>" +
                    "<td>" + billnumber.val() + "</td>" +
                    "<td>" + billdate.val() + "</td>" +
                    "<td>" + vendorname.val() + "</td>" +
                    "<td>" + particulars.val() + "</td>" +
                    "<td>" + amountinrupees.val() + "</td>" +
                "</tr>");

                PnPExpense.expenseItemDialog.dialog("close");
            }
            else {
                $("#div_error").show();
            }
        },

        GenerateExpenseReportId: function () {
            var num = Math.floor(Math.random() * 900000) + 100000;
            return num;
        },

        GetItemTypeForListName: function (name) {
            return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
        },

        GetEmailFromText: function (text) {
            var emails = text.match(/([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/gi);
            if (emails != null && emails.length > 0)
                return emails[0];
            return "";
        },

        EnsureUser: function (peoplepickerUser) {
            var userId = 0;
            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl + "/_api/web/ensureUser('" + encodeURIComponent(peoplepickerUser) + "')",
                type: "POST",
                async: false,
                contentType: "application/json;odata=verbose",
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val()
                },
                success: function (data) {
                    if (data.d.Id > 0)
                        userId = data.d.Id;
                },
                error: function (data) {
                    console.log(JSON.stringify(data));
                }
            });
            return userId;
        },

        SaveNewExpenseInfo: function () {
            var claimSubject = $("#claim_subject").val();
            var empNumber = $("#emp_code").val();
            var empNameAsPerBankAcc = $("#user_name_bank_account").val();
            var empBankAccNumber = $("#bank_account_number").val();
            var isBillToClient = $("#bill_to_client").val();
            var empDept = $("#dept").val();
            var empProject = $("#project").val();
            var clientName = $("#client_name").val();
            var randomclaimID = PnPExpense.GenerateExpenseReportId();
            var expenseTableItemData = PnPExpense.GetTableData();
            var claimAmount = expenseTableItemData.sum("amount");

            var itemType = PnPExpense.GetItemTypeForListName(PnPExpense.EXPENSE_LIST_NAME);
            var userid = _spPageContextInfo.userId;
            //"[{"Login":"i:0#.f|membership|acd.def@imaginea.com","Name":"abc  def","Email":"abc.def@xxx.com"}]"
            var peoplePickerEmail = PnPExpense.GetEmailFromText($("#hdnAdministrators").val().split(',')[2]);
            var peoplePickerClaimId = 'i:0#.f|membership|' + peoplePickerEmail;
            var item = {
                "__metadata": { "type": itemType },
                "Title": claimSubject,
                "ClaimNumber": String(randomclaimID),
                "EmployeeNumber": empNumber,
                "EmployeeName": empNameAsPerBankAcc,
                "BankAccountNumber": empBankAccNumber,
                "ReimburseFromClient": isBillToClient,
                "Department": empDept,
                "Project": empProject,
                "ClientName": clientName,
                "ClaimStatus": "New",
                "ClaimAmount": claimAmount,
                "ApproverId": PnPExpense.EnsureUser(peoplePickerClaimId)
            };
            //var digestValue = PnPExpense.GetDigest();
            $.ajax({
                url: PnPExpense.SiteURL + "_api/web/lists/getbytitle('" + PnPExpense.EXPENSE_LIST_NAME + "')/items",
                type: "POST",
                contentType: "application/json;odata=verbose",
                data: JSON.stringify(item),
                async: false,
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val()
                    //"X-RequestDigest": digestValue

                },
                success: function (data) {
                    $('#spanStatus').html('Expense report created successfully').removeClass('ticketerror').addClass('ticketsuccess');
                    $('.cam-peoplepicker-delImage').click();
                    PnPExpense.UploadFile(item.ClaimNumber, PnPExpense.EXPENSE_LIST_NAME, PnPExpense.SiteURL, "attachments");
                    PnPExpense.SaveNewExpenseItemInfo(expenseTableItemData, item.ClaimNumber);
                    console.log('Expense report created successfully')
                },
                error: function (data) {
                    $('#spanStatus').html('Expense report creation failed').removeClass('ticketsuccess').addClass('ticketerror');
                    $('.cam-peoplepicker-delImage').click();
                    console.log('Expense report creation failed')
                }
            });
        },

        SaveNewExpenseItemInfo: function (expenseTableItemData, claimNumber) {
            var itemType = PnPExpense.GetItemTypeForListName(PnPExpense.EXPENSEITEMS_LIST_NAME);
            for (expenseitemcounter = 0; expenseitemcounter < expenseTableItemData.length; expenseitemcounter++) {
                var expenseitem = {
                    "__metadata": { "type": itemType },
                    "ClaimNumber": claimNumber,
                    "BillNumber": expenseTableItemData[expenseitemcounter].billnumber,
                    "BillDate": expenseTableItemData[expenseitemcounter].billdate,
                    "BillTypeId": expenseTableItemData[expenseitemcounter].billtype,
                    "VendorName": expenseTableItemData[expenseitemcounter].vendorname,
                    "BillDescription": expenseTableItemData[expenseitemcounter].billdescription,
                    "BillAmount": expenseTableItemData[expenseitemcounter].amount
                };

                $.ajax({
                    url: PnPExpense.SiteURL + "_api/web/lists/getbytitle('" + PnPExpense.EXPENSEITEMS_LIST_NAME + "')/items",
                    type: "POST",
                    contentType: "application/json;odata=verbose",
                    data: JSON.stringify(expenseitem),
                    async: false,
                    headers: {
                        "Accept": "application/json;odata=verbose",
                        "X-RequestDigest": $("#__REQUESTDIGEST").val()
                    },
                    success: function (data) {
                        console.log('ExpenseItem created successfully');
                    },
                    error: function (data) {
                        console.log('ExpenseItem creation failed');
                    }
                });
            }
        },

        GetTableData: function () {
            var data = [];
            var i = 0;
            $('#users tbody tr').each(function (index, tr) {
                var tds = $(tr).find('td');
                if (tds.length > 1) {
                    data[i++] = {
                        sno: tds[0].textContent,
                        billtype: parseInt(tds[1].textContent),
                        billnumber: tds[2].textContent,
                        billdate: tds[3].textContent,
                        vendorname: tds[4].textContent,
                        billdescription: tds[5].textContent,
                        amount: parseFloat(tds[6].textContent)
                    }
                }
            });
            return data;
        },

        DisplayMyExpenses: function (expensesItems) {

            if (expensesItems.length > 0) {
                $("#Refresh-tickets").prop('disabled', false);
            } else {
                $("#Refresh-tickets").prop('disabled', true);
            }

            $('#tblMyExpenses').DataTable({
                destroy: true,
                "lengthMenu": [5, 10],
                data: expensesItems,
                columns: [
                    { title: "Claim Number#" },
                    { title: "Claim Description" },
                    { title: "Claim Date" },
                    { title: "Total Amount" },
                    { title: "Claim Status" },
                      {
                          "title": "",
                          "render": function (data, type) {
                              if (type === 'display' && data !== "") {
                                  return $('<a>')
                                      .attr('href', "#")
                                      .attr('id', data)
                                      .text("Display")
                                      .wrap('<div></div>')
                                      .parent()
                                      .html();

                              } else {
                                  return "Display";
                              }
                          }
                      }
                ]
            });
        },

        GetAllExpenses: function () {
            PnPExpense.myExpenses = [];
            $.ajax(
               {
                   url: PnPExpense.SiteURL + "_api/lists/getbytitle('" + PnPExpense.EXPENSE_LIST_NAME + "')/items?$select=Title,ClaimNumber,Created,ClaimStatus,ClaimAmount&$filter=AuthorId eq " + _spPageContextInfo.userId,
                   method: "GET",
                   headers:
                           {
                               "accept": "application/json;odata=verbose"
                           },
                   success: function (data) {
                       var results = data.d.results;
                       $.each(results, function (key) {
                           var value = results[key];
                           var expenseItem = [];
                           expenseItem.push(value.ClaimNumber);
                           expenseItem.push(value.Title);
                           expenseItem.push(new Date(value.Created).toLocaleDateString());
                           expenseItem.push(value.ClaimAmount);
                           expenseItem.push(value.ClaimStatus);
                           if (value.ClaimStatus !== "New") {
                               expenseItem.push(value.ClaimNumber);
                           } else {
                               expenseItem.push("");
                           }
                           PnPExpense.myExpenses.push(expenseItem);
                       });
                       PnPExpense.DisplayMyExpenses(PnPExpense.myExpenses);
                   },
                   error: function (err) {
                       $("#tblMyExpenses tbody").html("<h3>An error occured</h3><br/>" + JSON.stringify(err));
                   }
               }
           );
        },

        ExpenseDialogReset: function () {
            $("#bill_type").val(2); // 2 for dining
            $("#bill_no").val('');
            $("#bill_date").val('');
            $("#vendor_name").val('');
            $("#bill_particulars").val('');
            $("#bill_amount_rupees").val('');
        },

        GetFinanceManagerWorkFlowDetails: function () {
            var url = PnPExpense.SiteURL + "/_api/web/lists/getbytitle('" + PnPExpense.EXPENSE_FINANCE_MANAGERWORKFLOW_LIST_NAME + "')/ItemCount";
            $.ajax({
                url: url,
                method: "GET",
                async: false,
                headers: { "Accept": "application/json; odata=verbose" },
                success: function (data) {
                    if (data.d.ItemCount >= 1) {
                        $('#finaceManagerWorkFlow').hide();
                        $('#submit-expense').prop('disabled', false);
                    } else {
                        $('#finaceManagerWorkFlow').show();
                        $('#submit-expense').prop('disabled', true);
                    }
                },
                error: function (data) {
                    console.log('Error while retrieving information from FinanceManagerWorkflow');
                }
            });
        },

        SaveFinanceManagerWorkflowDetails: function () {
            var itemType = PnPExpense.GetItemTypeForListName(PnPExpense.EXPENSE_FINANCE_MANAGERWORKFLOW_LIST_NAME);
            //"[{"Login":"i:0#.f|membership|acd.def@imaginea.com","Name":"abc  def","Email":"abc.def@xxx.com"}]"
            var financePeoplePickerEmail = PnPExpense.GetEmailFromText($("#hdnAdministrators1").val().split(',')[2]);
            var financepeoplePickerClaimId = 'i:0#.f|membership|' + financePeoplePickerEmail;
            var item = {
                "__metadata": { "type": itemType },
                "FinanceManagerId": PnPExpense.EnsureUser(financepeoplePickerClaimId)
            };

            $.ajax({
                url: PnPExpense.SiteURL + "_api/web/lists/getbytitle('" + PnPExpense.EXPENSE_FINANCE_MANAGERWORKFLOW_LIST_NAME + "')/items",
                type: "POST",
                contentType: "application/json;odata=verbose",
                data: JSON.stringify(item),
                headers: {
                    "Accept": "application/json;odata=verbose",
                    "X-RequestDigest": $("#__REQUESTDIGEST").val()
                },
                success: function (data) {
                    $('#spanStatus').html('Finance Manager Workflow record created successfully').addClass('ticketsuccess');
                    $('#submit-expense').prop('disabled', false);
                    console.log('Finance Manager Workflow record created successfully');
                },
                error: function (data) {
                    $('#spanStatus').html('Finance Manager Workflow record creation failed').addClass('ticketerror');
                    console.log('Finance Manager Workflow record creation failed');
                }
            });
        },

        BindDropDownByIdListName: function (id, listName) {
            $.ajax({
                url: PnPExpense.SiteURL + "_api/web/lists/getbytitle('" + listName + "')/items",
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
                    $('#bill_type').val(2); // 2 for dining
                },
                error: function (data) {
                    console.log(data);
                    console.log('Operataion failed');
                }
            });
        },

        FormReset: function () {
            $("#emp_code").val('');
            $("#user_name_bank_account").val('');
            $("#bank_account_number").val('');
            $("#bill_to_client").val('false');
            $("#dept").val('');
            $("#project").val('');
            $("#client_name").val('');
            $("#attachments").val('');
            $('.cam-peoplepicker-delImage').click();
            $("#claim_subject").val('');

            $("#users").find("tbody tr").remove();
            PnPExpense.ExpenseDialogReset();

            return false;
        },

        GetExpenseByClaimNumber: function (claimNumber) {
            var expenseObject;
            var expenseurl = PnPExpense.SiteURL + "_api/lists/getbytitle('" + PnPExpense.EXPENSE_LIST_NAME + "')/items?$select=ID&$filter=ClaimNumber eq " + claimNumber;
            $.ajax(
                {
                    url: expenseurl,
                    async: false,
                    method: "GET",
                    headers:
                            {
                                "accept": "application/json;odata=verbose"
                            },
                    success: function ExpenseOnSuccess(data) {
                        expenseObject = data.d.results[0];
                    },
                    error: function onError(error) {
                        console.log('Error occurred while retrieving the ticket');
                    }
                }
            );
            return expenseObject;
        },

        WriteFileToList: function (file, digValue, siteURL, listitemid, listName) {
            var reader = new FileReader();
            reader.onload = function (e) {
                // get file content  
                if (file != null && file.name.length > 0) {
                    var fileUploadurl = siteURL + "_api/web/lists/getByTitle(@TargetLibrary)/Items(" + listitemid + ")/AttachmentFiles/add(FileName=@TargetFileName)?@TargetLibrary='" + listName + "'" + "&@TargetFileName='" + file.name + "'";
                    var filedata = e.target.result;
                    $.ajax({
                        url: fileUploadurl,
                        async: false,
                        type: "POST",
                        data: filedata,
                        processData: false,
                        headers: {
                            "Accept": "application/json;odata=verbose",
                            //"X-RequestDigest": $("#__REQUESTDIGEST").val()
                            "X-RequestDigest": digValue
                        },
                        success: function (data) {
                            console.log('File uploaded successfully');
                        },
                        error: function (data) {
                            console.log("File upload failed. Please try again");
                        }
                    });
                }
            }
            reader.readAsArrayBuffer(file);
        },

        UploadFile: function (claimNumber, listName, siteURL, fileControlId) {
            var digValue = PnPExpense.GetDigest();
            var expenseObject = PnPExpense.GetExpenseByClaimNumber(claimNumber);
            var listitemid = expenseObject.ID;
            var noOfFiles = $("#" + fileControlId)[0].files.length;
            if (noOfFiles > 0) {
                for (fileindex = 0; fileindex < noOfFiles; fileindex++) {
                    var file = $("#" + fileControlId)[0].files[fileindex];
                    PnPExpense.WriteFileToList(file, digValue, siteURL, listitemid, listName);
                }
            }
        },

        GetExpenseByID: function (claimNumber) {
            var expenseObject = "";
            var expenseurl = PnPExpense.SiteURL + "_api/lists/getbytitle('" + PnPExpense.EXPENSE_LIST_NAME + "')/items?$select=ManagerComments,FinanceManagerComments&$filter=ClaimNumber eq " + claimNumber;
            $.ajax(
                {
                    url: expenseurl,
                    async: false,
                    method: "GET",
                    headers:
                            {
                                "accept": "application/json;odata=verbose"
                            },
                    success: function onSuccess(data) {
                        expenseObject = data.d.results[0];
                    },
                    error: function onError(error) {
                        console.log('Error occurred while retrieving the Expense details');
                    }
                }
            );
            return expenseObject;
        },

        GetDigest: function () {
            var digestValue;
            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl + "/_api/contextinfo",
                method: "POST",
                async: false,
                headers: { "Accept": "application/json; odata=verbose" },
                success: function (data) {
                    //$('#__REQUESTDIGEST').val(data.d.GetContextWebInformation.FormDigestValue)
                    digestValue = data.d.GetContextWebInformation.FormDigestValue;
                },
                error: function (data, errorCode, errorMessage) {
                    alert(errorMessage)
                }
            });
            return digestValue;
        }
    }

    $("#bill_date").datepicker({ maxDate: '0' });

    PnPExpense.BindDropDownByIdListName("bill_type", "BillType");

    PnPExpense.GetFinanceManagerWorkFlowDetails();

    PnPExpense.GetAllExpenses();

    PnPExpense.expenseItemDialog = $("#dialog-form").dialog({
        autoOpen: false,
        height: 354,
        width: 500,
        modal: true,
        buttons: {
            "Add an expense": PnPExpense.AddExpense,
            Cancel: function () {
                PnPExpense.expenseItemDialog.dialog("close");
                PnPExpense.ExpenseDialogReset();
            }
        },
        close: function () {
            PnPExpense.expenseItemDialog.dialog("close");
            PnPExpense.ExpenseDialogReset();
        }
    });

    PnPExpense.itmanagerWorkflowDialog = $("#it-dialog-form").dialog({
        autoOpen: false,
        height: 200,
        width: 450,
        modal: true,
        Open: function () {
            $('.cam-peoplepicker-delImage').click();
        },
        buttons: {
            "Save": function () {
                if ($(".cam-entity-resolved").text().length > 0) {
                    PnPExpense.SaveFinanceManagerWorkflowDetails();
                    $('#finaceManagerWorkFlow').hide();
                    $('.cam-peoplepicker-delImage').click();
                    PnPExpense.itmanagerWorkflowDialog.dialog("close");
                }
                else {
                    return false;
                }
            },
            Cancel: function () {
                $('.cam-peoplepicker-delImage').click();
                PnPExpense.itmanagerWorkflowDialog.dialog("close");
            }
        },
        close: function () {
            PnPExpense.itmanagerWorkflowDialog.dialog("close");
        }
    });

    PnPExpense.displayCommentsDialog = $("#comments-dialog-form").dialog({
        autoOpen: false,
        height: 250,
        width: 540,
        modal: true,
        Open: function () {
            $('.cam-peoplepicker-delImage').click();
        },
        buttons: {            
            Cancel: function () {
                $('.cam-peoplepicker-delImage').click();
                PnPExpense.displayCommentsDialog.dialog("close");
            }
        },
        close: function () {
            PnPExpense.displayCommentsDialog.dialog("close");
        }
    });

    $("#add-expense").on("click", function () {
        PnPExpense.expenseItemDialog.dialog("open");
    });

    $("#reset").on("click", function () {
        PnPExpense.FormReset();
        return false;
    });

    $("#submit-expense").click(function () {

        var isDataValid = PnPExpense.CheckFormDataValidation();
        var noOfExpenses = $("#users tbody").children().length;
        if (isDataValid) {
            if (noOfExpenses > 0) {
                $("#div_error_report").hide();
                PnPExpense.SaveNewExpenseInfo();
                PnPExpense.GetAllExpenses();
                PnPExpense.FormReset();
                return false;
            }
            else {
                alert("You do not have any expenses to be submitted!!!");
                return false;
            }
        }
        else {
            $("#div_error_report").show();
            return false;
        }
    });

    $('#Refresh-tickets').click(function () {
        PnPExpense.myExpenses.length = 0;
        var dataTable = $('#tblMyExpenses').DataTable();
        dataTable.clear().draw();
        PnPExpense.GetAllExpenses();
        return false;
    });

    $('#finaceManagerWorkFlow').click(function () {
        PnPExpense.itmanagerWorkflowDialog.dialog("open");
    });

    $("#tblMyExpenses").on("click", "a", function (event) {
        var claimNumber = $(this).attr('id');
        var expenseObj = PnPExpense.GetExpenseByID(claimNumber);
        $("#mngr_comments").html(expenseObj.ManagerComments);
        $("#it_mngr_comments").html(expenseObj.FinanceManagerComments);
        PnPExpense.displayCommentsDialog.dialog("open");

    });

    Array.prototype.sum = function (prop) {
        var total = 0.0;
        for (var i = 0, _len = this.length; i < _len; i++) {
            total += this[i][prop]
        }
        return total
    }
});