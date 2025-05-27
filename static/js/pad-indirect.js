document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('report-form');
    const generateButton = document.getElementById('generate');

    generateButton.addEventListener('click', function () {
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

        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;

        const formData = {
            start_date: startDate,
            end_date: endDate
        };
        console.log('JSON Data:', formData);

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/pad-indirect/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = `PAD_Indirect_Time_Report_${startDate}_${endDate}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
                    messageDiv.style.display = 'none';
                } else {
                    var reader = new FileReader();
                    reader.onload = function () {
                        var errorMessage = JSON.parse(reader.result).error || 'Error generating the report';
                        alert(`Failed to generate report: ${errorMessage}`);
                    };
                    reader.readAsText(xhr.response);
                }
            }
        };

        // Send the request
        xhr.send(JSON.stringify(formData));
    });
});