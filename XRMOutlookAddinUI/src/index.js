/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
//import 'bootstrap';
//import 'bootstrap-select';

var queryobj = {
    sitecollection: "",
    list: "",
    statusfilter: "",
    clientfilter: "",
    clientfield: "",
    stakeholderfilter: "",
    stakeholderfield: "",
    filterfield: ""
}

var mailitem = {
    Subject: "",
    To: "",
    From: "",
    ConversationId: "",
    Received: "",
    Message: "",
    ConversationTopic:"",
    itemid:"",
    listid:"",
    sitecollectionUrl:"",
    listname:""
}

// $(document).ready(function () {
//     fetchConfigData();
//    fetchContractFilterData();
//    loadData();
//     //getMailData(Office.context.mailbox.item);
// });

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    //when you browse the page outside outlook load the document.ready outside the this method.
    $(document).ready(function () {
        fetchConfigData();
       fetchContractFilterData();
       loadData();
        getMailData(Office.context.mailbox.item);
    });
};


function getListItems(querydata) {

    var querystring = "";
    if (querydata.clientfield.length > 0 && querydata.stakeholderfield.length > 0) {
        querystring = "sc=" + querydata.sitecollection + "&list=" + querydata.list + "&ff=" + querydata.clientfield + "&val=" + querydata.clientfilter + "&ff1=" + querydata.stakeholderfield + "&val1=" + querydata.stakeholderfilter;
    } else if (querydata.clientfield.length > 0) {
        querystring = "sc=" + querydata.sitecollection + "&list=" + querydata.list + "&ff=" + querydata.clientfield + "&val=" + querydata.clientfilter;
    } else if (querydata.stakeholderfield.length > 0) {
        querystring = "sc=" + querydata.sitecollection + "&list=" + querydata.list + "&ff=" + querydata.stakeholderfield + "&val=" + querydata.stakeholderfilter;
    }
    else {
        querystring = "sc=" + querydata.sitecollection + "&list=" + querydata.list + "&ff=" + querydata.filterfield + "&val=" + querydata.statusfilter;
    }

    console.log(querystring);
    fetchListItems(querystring);

}


function loadData() {
    //$('#run').click(run);
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
        $(this).attr("disabled", "disabled");
        //console.log(optionselected.text());
    });

    //Event handler for lists change event
    $("#listsdd").on("change", function (event) {
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
        $("#listsdd").attr("disabled", "disabled");
    });

    //event handler for filter change event
    $("#casestatus,#projectstatus,#relatedClient,#relatedStakeholder").on("change", function (event) {
        var optionselected = $(this).find("option:selected");
        var parentselect = optionselected.prevObject;
        if (parentselect[0].id == "casestatus") {
            queryobj.statusfilter = optionselected.val();
            queryobj.filterfield = "StatusLookupId";
        } else if (parentselect[0].id == "projectstatus") {
            queryobj.statusfilter = optionselected.val();
            queryobj.filterfield = "Status";
        } else if (parentselect[0].id == "relatedClient") {
            queryobj.clientfilter = optionselected.val();
            queryobj.clientfield = "Client_x0020_Contract_x0020_PartLookupId";
        } else if (parentselect[0].id == "relatedStakeholder") {
            queryobj.stakeholderfilter = optionselected.val();
            queryobj.stakeholderfield = "Stakeholder_x0020_Contract_x0020LookupId";
        }

        $("#btnFetch").css("display", "block");
    });

    // event handler for fetch
    $("#btnFetch").click(function (event) {
        $("#xrmitems").css("display", "block");
        getListItems(queryobj);
    });
}

function fetchConfigData() {
    $(".loader").css("display", "block");
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
            $(".loader").css("display", "none");
        },
        error: function (data) { console.log(data); }
    });
}

function fetchContractFilterData() {
    $(".loader").css("display", "block");
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
            $(".loader").css("display", "none");
        },
        error: function (data) { console.log(data); }
    });
}

function fetchListItems(queryString) {
    $(".loader").css("display", "block");
    console.log("Fetching list item data");
    $("#ddsaveemail").css("display", "block");
    $("#ddsaveattachments").css("display", "block");
    $("#btnSave").css("display", "block");
    $.ajax({
        url: "https://xrmoutlookaddin.azurewebsites.net/api/GetListItems?code=nL0I4H0QhnTBUU7fXOMrY4WB0oJ3tZc5TMk0mtBpxM168KGJUJthng==&" + queryString,
        method: "Get",
        headers: { "Accept": "application/json;odata=verbose" },
        success: function (data) {
            console.log(data);
            $.each(data, (index, value) => {
                $("#xrmitemsDD").append('<option value="' + value.ID + '">' + value.Title + '</option>')
            });
            $('#xrmitemsDD').selectpicker();
            $('#xrmitemsDD').addClass("selectpicker");
            $("#btnFetch").css("display", "none");
            $(".loader").css("display", "none");
        },
        error: function (data) { console.log(data); }
    });
}

function getMailData(item) {
    $(".loader").css("display", "block");
    // //Office.context.mailbox.item.to.getAsync(getToAddress);
    mailitem.subject = item.subject;
    mailitem.from = buildEmailAddressString(item.from);
    mailitem.created = item.dateTimeCreated;
    mailitem.conversation = item.conversationId;
    Office.context.mailbox.item.body.getAsync('text', function (result) {
        if (result.status === 'succeeded') {
            mailitem.body = result.value;
        }
    });

    mailitem.to=buildToEmailAddressesString(item.to);

    //   Office.context.mailbox.item.body.getAsync('html', function(result){
    //     if (result.status === 'succeeded') {
    //         console.log(result.value);
    //     }
    //   });

    console.log(mailitem);
    $(".loader").css("display", "none");
}

// Format an EmailAddressDetails object as
  // GivenName Surname <emailaddress>
  function buildEmailAddressString(address) {
    return address.displayName + "," + address.emailAddress + ";";
  }
  
  // Take an array of EmailAddressDetails objects and
  // build a list of formatted strings, separated by a line-break
  function buildToEmailAddressesString(addresses) {
    if (addresses && addresses.length > 0) {
      var returnString = "";
      
      for (var i = 0; i < addresses.length; i++) {
        if (i > 0) {
          returnString = returnString + "<br/>";
        }
        returnString = returnString + buildEmailAddressString(addresses[i]);
      }
      
      return returnString;
    }
    
    return "None";
  }

