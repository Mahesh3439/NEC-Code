var CustomJS = {
    load: function () {
        $('#Attachments').MultiFile({
            onFileChange: function () {                
            }
        });
    },

    fnAlert: function (text) {
        alert(text);
    },

    fnUploadAttachments: function (siteURL,itemID) {
        var fileArray = [];
        if ($("#Attachments input:file")[0].files.length > 0) {
            $("#Attachments input:file").each(function () {
                var file = $(this)[0].files[0];
                var fileName = file.name;
                getFileBuffer(file).then(
                    function (buffer) {
                        var bytes = new Uint8Array(buffer);
                        var binary = '';
                        for (var b = 0; b < bytes.length; b++) {
                            binary += String.fromCharCode(bytes[b]);
                        }
                        console.log(' File size:' + bytes.length);

                        var URL = siteURL + "/_api/web/lists/getbytitle('Projects')/items(" + itemID + ")/AttachmentFiles/add(FileName='" + fileName + "')";
                        $.ajax({
                            url: URL,
                            type: "POST",
                            processData: false,
                            binaryStringRequestBody: true,
                            contentType: "application/json;odata=verbose",
                            data: binary,
                            headers: {
                                "accept": "application/json;odata=verbose",
                                "X-RequestDigest": $("#__REQUESTDIGEST").val(),
                                "content-length": binary.byteLength
                            },
                            success: function (data) {
                                console.log(data + ' uploaded successfully');
                                //deferred.resolve(data);
                            },
                            error: function (data) {
                                var error =  data.responseText;
                                console.log(fileName + "not uploaded error");
                              //  deferred.reject(data);
                            }
                        });

                    },
                    function (err) {
                        deferred.reject(err);
                    });
                //return deferred.promise();           
            });
        }



    }

};

function getFileBuffer(file) {
    var deferred = $.Deferred();
    var reader = new FileReader();
    reader.onload = function (e) {
        deferred.resolve(e.target.result);
    }
    reader.onerror = function (e) {
        deferred.reject(e.target.error);
    }
    reader.readAsArrayBuffer(file);
    return deferred.promise();
}
