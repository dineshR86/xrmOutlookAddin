/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import 'bootstrap';
//import 'bootstrap-select';

var queryobj = {
    sitecollection: "",
    list: "",
    statusfilter: "",
    clientfilter: "",
    stakeholderfilter: "",
    filterfield:""
}

var mailitem = {
    subject: "",
    to: "",
    from: "",
    conversation: "",
    created: ""
}

$(document).ready(() => {
    fetchConfigData();
    fetchContractFilterData();
    loadData();
});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

function getToAddress(asyncResult) {
    mailitem.to = asyncresult.value;
}

async function run() {
    /**
         * Insert your Outlook code here
         */
}


function getListItems(querydata) {
    $("#xrmitems").css("display", "block");
    var querystring = "sc="+querydata.sitecollection+"&list="+querydata.list+"&ff="+querydata.filterfield+"&val="+querydata.statusfilter;
    console.log(querystring);
    fetchListItems(querystring);
    // //Office.context.mailbox.item.to.getAsync(getToAddress);
    // mailitem.subject=Office.context.mailbox.item.subject;
    // mailitem.from=Office.context.mailbox.item.sender.emailAddress;
    // mailitem.created=Office.context.mailbox.item.dateTimeCreated;
    // mailitem.conversation=Office.context.mailbox.item.conversationId;
    // mailitem.to=Office.context.mailbox.item.to;
}


function loadData() {
    $('#run').click(run);
    $.fn.selectpicker.Constructor.BootstrapVersion = '4';
    //Event handler for site collection dropdown
    $("#sitecollections").on("change", function (event) {
        var optionselected = $(this).find("option:selected");
        if (optionselected.text() == "-select-") {
            $("#lists").css("display", "none");
        } else {
            $("#lists").css("display", "block");
            queryobj.sitecollection = optionselected.val();
        }
        console.log(optionselected.text());
    });

    //Event handler for lists change event
    $("#lists").on("change", function (event) {
        var optionselected = $(this).find("option:selected");
        if (optionselected.text() == "-select-") {
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "none");
        } else if (optionselected.val().indexOf("Cases") > -1) {
            $("#casefilter").css("display", "block");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "none");
        } else if (optionselected.val().indexOf("Projects") > -1) {
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "block");
            $("#contractfilter").css("display", "none");
        } else if (optionselected.val().indexOf("Contracts") > -1) {
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "block");
        }
        queryobj.list = optionselected.val();
    });

    //event handler for filter change event
    $("#casestatus,#projectstatus,#relatedClient,#relatedStakeholder").on("change", function (event) {
        var optionselected = $(this).find("option:selected");
        var parentselect = optionselected.prevObject;
        if (parentselect[0].id == "casestatus") {
            queryobj.statusfilter = optionselected.val();
            queryobj.filterfield="StatusLookupId";
        } else if (parentselect[0].id == "projectstatus") {
            queryobj.statusfilter = optionselected.val();
            queryobj.filterfield="";
        }
        else if (parentselect[0].id == "relatedClient") {
            queryobj.clientfilter = optionselected.val();
        } else if (parentselect[0].id == "relatedStakeholder") {
            queryobj.stakeholderfilter = optionselected.val();
        }

        console.log(queryobj);
        console.log(mailitem);
        getListItems(queryobj);
    });
}

function fetchConfigData() {
    console.log("Fetching Config list data");
    $.ajax({
        url: "https://xrmoutlookaddin.azurewebsites.net/api/GetXRMAddInConfiguration?code=nzUUuX1DObCOn5GTzvoLGR/nRDU6Pog08RY6jMHNvpBp/zz0dgd/DQ==",
        method: "Get",
        headers: { "Accept": "application/json;odata=verbose" },
        success: function (data) {
            //var configdata=JSON.parse(data);
            $.each(data.SiteCollectionUrls.split(";"), (index, value) => {
                $("#sitecollections").append('<option value="' + value + '">' + value + '</option>')
            });

            $.each(data.Lists.split(";"), (index, value) => {
                $("#listsdd").append('<option value="' + value + '">' + value + '</option>')
            });

            $.each(data.ProjectStatusFilter.split(";"), (index, value) => {
                $("#projectstatus").append('<option value="' + value + '">' + value + '</option>')
            });
        },
        error: function (data) { console.log(data); }
    });
}

function fetchContractFilterData() {
    console.log("Fetching Config list data");
    $.ajax({
        url: "https://xrmoutlookaddin.azurewebsites.net/api/GetContractFilters?code=JwRjIrMznRj4r4XwPKb1ERaTX7rrjaz7qp/YUAyrj7K2PEr8129EMw==",
        method: "Get",
        headers: { "Accept": "application/json;odata=verbose" },
        success: function (data) {
            $.each(data.Clients, (index, value) => {
                var clientoptions = value.split(",");
                $("#relatedClient").append('<option value="' + clientoptions[1] + '">' + clientoptions[0] + '</option>')
            });

            $.each(data.Stakeholders, (index, value) => {
                var stakeholderoptions = value.split(",");
                $("#relatedStakeholder").append('<option value="' + stakeholderoptions[1] + '">' + stakeholderoptions[0] + '</option>')
            });

            $.each(data.Status, (index, value) => {
                var statusOptions = value.split(",");
                $("#casestatus").append('<option value="' + statusOptions[1] + '">' + statusOptions[0] + '</option>')
            });
        },
        error: function (data) { console.log(data); }
    });
}

function fetchListItems(queryString){
    console.log("Fetching list item data");
    $.ajax({
        url: "https://xrmoutlookaddin.azurewebsites.net/api/GetListItems?code=nL0I4H0QhnTBUU7fXOMrY4WB0oJ3tZc5TMk0mtBpxM168KGJUJthng==&"+queryString,
        method: "Get",
        headers: { "Accept": "application/json;odata=verbose" },
        success: function (data) {
            console.log(data);
            $.each(data, (index, value) => {
                $("#xrmitemsDD").append('<option value="' + value.ID + '">' + value.Title + '</option>')
            });
            $('#xrmitemsDD').selectpicker();
            $('#xrmitemsDD').addClass("selectpicker");
        },
        error: function (data) { console.log(data); }
    });
}

