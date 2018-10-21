/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import 'bootstrap';
//import 'bootstrap-select';

var queryobj={
    sitecollection:"",
    list:"",
    statusfilter:"",
    clientfilter:"",
    stakeholderfilter:""
}

var mailitem={
    subject:"",
    to:"",
    from:"",
    conversation:"",
    created:""
}



// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
    $(document).ready(() => {
        
        loadData();
    });
    
};

function getToAddress(asyncResult){
    mailitem.to=asyncresult.value;
}

async function run() {
    /**
         * Insert your Outlook code here
         */
}

function getListItems(querydata){
    $("#xrmitems").css("display","block");
        //Office.context.mailbox.item.to.getAsync(getToAddress);
        mailitem.subject=Office.context.mailbox.item.subject;
        mailitem.from=Office.context.mailbox.item.sender.emailAddress;
        mailitem.created=Office.context.mailbox.item.dateTimeCreated;
        mailitem.conversation=Office.context.mailbox.item.conversationId;
        mailitem.to=Office.context.mailbox.item.to;
}

function loadData(){
    $('#run').click(run);
    $.fn.selectpicker.Constructor.BootstrapVersion = '4';
    //Event handler for site collection dropdown
    $("#sitecollections").on("change", function (event) {
        var optionselected=$(this).find("option:selected");
        if(optionselected.text()=="-select-"){
            $("#lists").css("display", "none");
        }else{
            $("#lists").css("display", "block");
            queryobj.sitecollection=optionselected.val();
        }
        console.log(optionselected.text());
    });

    //Event handler for lists change event
    $("#lists").on("change", function (event) {
        var optionselected=$(this).find("option:selected");
        if(optionselected.text()=="-select-"){
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "none");
        }else if(optionselected.val()=="list1"){
            $("#casefilter").css("display", "block");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "none");
        }else if(optionselected.val()=="list2"){
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "block");
            $("#contractfilter").css("display", "none");
        }else if(optionselected.val()=="list3"){
            $("#casefilter").css("display", "none");
            $("#projectfilter").css("display", "none");
            $("#contractfilter").css("display", "block");
        }
        queryobj.list=optionselected.val();
    });

    //event handler for filter change event
    $("#casestatus,#projectstatus,#relatedClient,#relatedStakeholder").on("change", function (event) {
        var optionselected=$(this).find("option:selected");
        var parentselect=optionselected.prevObject;
        if(parentselect[0].id=="casestatus"){
            queryobj.statusfilter=optionselected.val();
        }else if(parentselect[0].id=="projectstatus"){
            queryobj.statusfilter=optionselected.val();
        }
        else if(parentselect[0].id=="relatedClient"){
            queryobj.clientfilter=optionselected.val();
        }else if(parentselect[0].id=="relatedStakeholder"){
            queryobj.stakeholderfilter=optionselected.val();
        }

        console.log(queryobj);
        console.log(mailitem);
        getListItems(queryobj);
    });
}

