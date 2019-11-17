var CustomJS = {
    load: function () {
        $('#Attachments').MultiFile({
            onFileChange: function () { 
                console.log(this, arguments);               
            }
        });
    },

    fnclear: function () {
        
        
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
