function generateReport() {
    var messageDiv = document.getElementById('loading-message');
    if (!messageDiv) {
        messageDiv = document.createElement('div');
        messageDiv.id = 'loading-message';
        messageDiv.style.position = 'fixed';
        messageDiv.style.top = '50%';
        messageDiv.style.left = '50%';
        messageDiv.style.transform = 'translate(-50%, -50%)';
        messageDiv.style.padding = '20px';
        messageDiv.style.backgroundColor = '#666';
        messageDiv.style.border = '1px solid #ccc';
        messageDiv.style.zIndex = '1000';
        messageDiv.style.textAlign = 'center';
        messageDiv.innerHTML = '<p>Generating the report, please be patient. This might take a few minutes...</p>';
        document.body.appendChild(messageDiv);
    }

    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var start_date = formData.get('start_date');
    var end_date = formData.get('end_date');
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

    if (!start_date) {
        alert("Please select a start date.");
        return;
    }

    var jsonData = JSON.stringify({
        start_date: start_date,
        end_date: end_date,
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
                messageDiv.style.display = 'none';
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