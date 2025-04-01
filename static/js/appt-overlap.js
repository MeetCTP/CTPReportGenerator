document.addEventListener('DOMContentLoaded', function () {
    var generateButton = document.getElementById('generate');
    if (generateButton) {
        generateButton.addEventListener('click', function () {
            generateReport();
        });
    }

    function generateReport() {
        var startDate = document.getElementById('start-date').value;
        var endDate = document.getElementById('end-date').value;
        var provider = document.getElementById('provider').value.trim();
        var client = document.getElementById('client').value.trim();

        if (!startDate || !endDate) {
            alert("Please fill in both the start and end dates.");
            return;
        }

        var jsonData = {
            start_date: startDate,
            end_date: endDate,
            provider: provider ? provider : null,
            client: client ? client : null
        };

        console.log('JSON Data:', jsonData);

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/appt-overlap/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = 'Overlapping_Appointment_Report.xlsx';
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                } else {
                    var reader = new FileReader();
                    reader.onload = function () {
                        var errorMessage = reader.result;
                        console.error('Error generating report:', errorMessage);
                        alert('Error generating report. Please try again later.');
                    };
                    reader.readAsText(xhr.response);
                }
            }
        };

        xhr.send(JSON.stringify(jsonData));
    }
});