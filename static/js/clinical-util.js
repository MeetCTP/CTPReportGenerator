document.addEventListener('DOMContentLoaded', function () {
    const form = document.getElementById('report-form');
    const generateButton = document.getElementById('generate');

    generateButton.addEventListener('click', function () {
        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;
        const provider = document.getElementById('provider').value;
        const client = document.getElementById('client').value;

        if (!startDate || !endDate) {
            alert('Please fill in required fields');
            return;
        }

        if (!provider && !client) {
            alert('Please choose at least one provider or client');
            return;
        }

        const formData = {
            start_date: startDate,
            end_date: endDate,
            provider: provider,
            client: client
        };
        console.log('JSON Data:', formData);

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/clinical-util-tracker/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                    var url = window.URL.createObjectURL(blob);
                    var a = document.createElement('a');
                    a.href = url;
                    a.download = `Clinical_Util_Tracker_${provider}_${startDate}_${endDate}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    document.body.removeChild(a);
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