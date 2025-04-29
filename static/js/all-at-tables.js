document.addEventListener('DOMContentLoaded', function () {
    const generateButton = document.getElementById('generate');

    if (generateButton) {
        generateButton.addEventListener('click', function () {
            generateReport();
        });
    }

    function generateReport() {
        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;

        if (!startDate || !endDate) {
            alert("Please select both start and end dates.");
            return;
        }

        const jsonData = JSON.stringify({
            start_date: startDate,
            end_date: endDate
        });

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/monthly-goals-report/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = `All_Airtables_${startDate}-${endDate}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    console.log("Response Received")
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

        console.log("Request sent")
        console.log(jsonData)
        xhr.send(jsonData);
    }
});