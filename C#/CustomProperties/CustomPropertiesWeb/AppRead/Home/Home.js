/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// Add any initialization logic to this function.
var customProperties;
var customPropertiesAreLoaded = false;
var imageId = null;
var item;

Office.initialize = function (reason) {
    initApp();
}

// Initializes the mail app for Outlook.
function initApp() {    
    Office.context.mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    item = Office.context.mailbox.item;
    itemBody();
}

function itemBody(){
    item.body.getTypeAsync(function (result) {
        if (result.status == Office.AsyncResultStatus.Failed)
            console.log("Error");
        else {            
            if (result.value == Office.MailboxEnums.BodyType.Html && Office.context.roamingSettings.get("signature") != null) {
                console.log("Here");
                item.body.prependAsync(
                        '<br><br><img src="http://i.imgur.com/' + Office.context.roamingSettings.get("signature") + '.jpg"/>',
                        {
                            coercionType: Office.CoercionType.Html,
                            asyncContext: { var3: 1, var4: 2 }
                        },
                        function (asyncResult) {
                            if (asyncResult.status == Office.AsyncResultStatus.Failed)
                                console.log("Error in the prepend");
                            else
                                console.log("Everything is ok");
                        }
                );
            }
            else
                console.log(result.value);
        }
    });
}

// Function called after the local custom properties are loaded from the Exchange server.
function customPropsCallback(asyncResult) {
    customProperties = asyncResult.value;
    customPropertiesAreLoaded = true;

    console.log(Office.context.roamingSettings.get("signature"));
    if (Office.context.roamingSettings.get("signature") != null)
    {
        document.getElementById('output').innerHTML = Office.context.roamingSettings.get("signature");
    }

}

function storeData(asyncResult) {       

    html2canvas($(".signature-generated"), {
        onrendered: function (canvas) {
            var imageData = canvas.toDataURL('image/jpeg', '1.0');
            imageData = imageData.split(',')[1];

            var form = new FormData();
            form.append("image", imageData);
            form.append("album", "vCf0c");
            form.append("type", "");
            form.append("name", "");
            form.append("title", "signatureImg");
            form.append("description", "Signature image");

            var settings = {
                "async": true,
                "type": "POST",
                "crossDomain": true,
                "url": "https://api.imgur.com/3/image",
                "data": form,
                "mimeType": "multipart/form-data",
                "processData": false,
                "contentType": false,
                "headers": {
                    "authorization": "Bearer 9e62c64c3ee7452983d72756d39535bcae010230"
                }
            }
            $.ajax(settings).done(function (response) {
                console.log(response);
                var json = JSON.parse(response);
                imageId = json.data.id;
                showToast("Success", "Signature stored!");
                Office.context.roamingSettings.set("signature", imageId);
                Office.context.roamingSettings.saveAsync();
            })
            .fail(function (err) {
                console.log(err);
            });
        }
    });
    
}

function getData() {
    console.log(Office.context.roamingSettings.get("signature"));
    showToast("Data", Office.context.roamingSettings.get("signature"));
}

// Displays the toast.
function showToast(title, message) {
    var notice = document.getElementById("notice");
    var output = document.getElementById('output');

    notice.innerHTML = title;
    output.innerHTML = message;
    
}

