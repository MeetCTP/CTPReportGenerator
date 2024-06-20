function generateReport() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var app_start = formData.get('app-start');
    var app_end = formData.get('app-end');
    var converted_after = formData.get('converted-after');

    var jsonData = JSON.stringify({
        app_start: app_start,
        app_end: app_end,
        converted_after: converted_after
    });
    console.log('JSON Data:', jsonData);

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/late-conversions/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Late_Conversions_Report.xlsx';
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