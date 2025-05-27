document.addEventListener('DOMContentLoaded', function () {
    const generateButton = document.getElementById('generate');

    if (generateButton) {
        generateButton.addEventListener('click', function () {
            generateReport();
        });
    }

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
        var selectedTable = formData.get('table');

        let selectedTables = []
        if (selectedTable === "Paraprofessional") {
            selectedTables = [
                "Paraprofessional",
                "Archived Para Apps 2021-2022",
                "Archived Para Apps 2019-2021",
                "Archived Para Apps 08.15.2022",
                "Simple Tracker (Not to use)"
            ]
        } else {
            selectedTables = [selectedTable]
        }

        var jsonData = JSON.stringify({
            tables: selectedTables
        });

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/valid-emails/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = `Valid_Email_Addresses.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    messageDiv.style.display = 'none';
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