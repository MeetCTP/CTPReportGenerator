function generateReport() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var status = [];
    var pg_type = [];
    var service_type = [];

    formData.forEach((value, key) => {
        if (key === 'status') status.push(value);
        if (key === 'pg_type') pg_type.push(value);
        if (key === 'service_type') service_type.push(value);
    });

    var jsonData = JSON.stringify({ status: status, pg_type: pg_type, service_type: service_type });
    console.log('JSON Data:', jsonData);

    // Send an AJAX request to the Flask backend
    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/active-contacts/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';  // Set response type to blob

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Active_Contacts_Report.xlsx';
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

    console.log(jsonData)

    // Send the selected filters as JSON data in the request body
    xhr.send(jsonData);
}