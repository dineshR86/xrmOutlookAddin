/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import 'bootstrap';
//import 'bootstrap-select';

$(document).ready(() => {
    $('#run').click(run);
    $.fn.selectpicker.Constructor.BootstrapVersion = '4';
    $("#sitecollections").on("change", function (event) {
        // if ($(this).find("option:selected").val() > 1) {
        //     $("#lists").css("display", "block");
        // }
        console.log($(this).find("option:selected"));
        $("#lists").css("display", "block");    
    });

});

// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run() {
    /**
         * Insert your Outlook code here
         */
}

