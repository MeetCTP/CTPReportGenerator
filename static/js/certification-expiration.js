document.addEventListener('DOMContentLoaded', function() {
    var generateButton = document.getElementById('generate');
    if (generateButton) {
        generateButton.addEventListener('click', function() {
            generateReport();
        });
    }

    function generateReport() {
        var form = document.getElementById('report-form');
        var formData = new FormData(form);

        var status = formData.get('status');
        var timeframe = formData.get('timeframe');
        var provider = formData.get('provider');

        var jsonData = JSON.stringify({
            status: status,
            timeframe: timeframe,
            provider: provider
        });
        console.log('JSON Data:', jsonData);

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/certification-expiration/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = 'Certification-Expiration.xlsx';
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

        xhr.send(jsonData);
    }
});