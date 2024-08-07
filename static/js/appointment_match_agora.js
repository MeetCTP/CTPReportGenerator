function generateReport() {
    // Get the start and end date values from the input fields
    var startDate = document.getElementById('start-date').value;
    var endDate = document.getElementById('end-date').value;

    var jsonData = JSON.stringify({ start_date: startDate, end_date: endDate });
    console.log('JSON Data:', jsonData);

    // Send an AJAX request to the Flask backend
    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/agora-match/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';  // Set response type to blob

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Agora_Appointment_Match_Report.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
            } else {
                // Error: Report generation failed, handle error (e.g., display an error message)
                var reader = new FileReader();
                reader.onload = function () {
                    var errorMessage = reader.result;
                    console.error('Error generating report:', errorMessage);
                };
                reader.readAsText(xhr.response);
            }
        }
    };

    // Send the start and end date values as JSON data in the request body
    xhr.send(JSON.stringify({ start_date: startDate, end_date: endDate }));
}