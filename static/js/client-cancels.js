function generateReport() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var range_start = formData.get('start_date')
    var provider = formData.get('provider');
    var client = formData.get('client');
    var cancelReasons = [];

    var checkboxes = form.querySelectorAll('input[name="cancel_reason"]:checked');
    checkboxes.forEach(function (checkbox) {
        cancelReasons.push(checkbox.value);
    });

    if (cancelReasons.length === 0) {
        alert("Please select at least one cancellation reason.");
        return;
    }

    if (!range_start) {
        alert("Please select a start date.");
        return;
    }

    var jsonData = JSON.stringify({
        range_start: range_start,
        provider: provider,
        client: client,
        cancel_reasons: cancelReasons
    });
    console.log('JSON Data:', jsonData);

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/client-cancels/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Client_Cancel_Report.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
            } else {
                var reader = new FileReader();
                reader.onload = function () {
                    var errorMessage = reader.result;
                    console.error('Error generating report:', errorMessage);
                };
                reader.readAsText(xhr.response);
            }
        }
    };

    console.log(jsonData);
    xhr.send(jsonData);
}