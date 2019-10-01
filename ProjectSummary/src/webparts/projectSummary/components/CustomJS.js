var CustomJS = {
    load: function () {
        $('#Attachments').MultiFile({
            onFileChange: function () {
                alert('hello mahesh');
            }
        });
    },

    fnAlert: function (text) {
        alert(text);
    }
};


window['jsLoadInit'] = () =>
{
    $('#Attachments').MultiFile({
        onFileChange: function () {
            alert('hello mahesh');
        }
    });

    alert('hi');
}