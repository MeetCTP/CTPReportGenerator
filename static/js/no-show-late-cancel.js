document.addEventListener('DOMContentLoaded', function() {
    var generateButton = document.getElementById('generate');
    if (generateButton) {
        generateButton.addEventListener('click', function() {
            generateReport();
        });
    }
});

function generateReport() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var app_start = formData.get('app-start');
    var app_end = formData.get('app-end');
    var provider = formData.get('provider');
    var client = formData.get('client');
    var school = formData.get('school');

    var jsonData = JSON.stringify({
        app_start: app_start,
        app_end: app_end,
        provider: provider,
        client: client,
        school: school
    });
    console.log('JSON Data:', jsonData);

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/no-show-late-cancel/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'No_Show_Late_Cancel_Report.xlsx';
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