document.addEventListener('DOMContentLoaded', function() {
    var generateButton = document.getElementById('generate');
    if (generateButton) {
        generateButton.addEventListener('click', function() {
            generateReport();
        });
    }
});

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

    const schoolSplit = school.split(': ')
    const schoolName = schoolSplit[1]

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
                a.download = `No_Show_Late_Cancel_Report_${schoolName}.xlsx`;
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