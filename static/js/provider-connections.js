function generateReport() {
    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/provider-connections/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Provider_Connections_Report.xlsx';
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

    xhr.send();
}