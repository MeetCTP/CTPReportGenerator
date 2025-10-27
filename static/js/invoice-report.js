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
            const logo = document.createElement('img');
            logo.src = logoUrl;
            logo.alt = 'Loading...';
            logo.id = 'company-logo';
            logo.style.width = '50px';  // Adjust size if needed
            logo.style.marginBottom = '10px';
            messageDiv.appendChild(logo);
            messageDiv.innerHTML = '<p>Generating the report, please be patient. This might take a few minutes...</p>';
            document.body.appendChild(messageDiv);
        }

        const rawFileInput = document.getElementById('input_file');
        const rawFile = rawFileInput.files[0];
        const lcnsFileInput = document.getElementById('lcns_file');
        const lcnsFile = lcnsFileInput.files[0];
        const selectedSchool = document.getElementById('school').value;

        const formData = new FormData();
        formData.append('input_file', rawFile);
        formData.append('school', selectedSchool);
        if (lcnsFile) {
            formData.append('lcns_file', lcnsFile);
        }

        console.log("FormData Contents:");
        formData.forEach((value, key) => {
            console.log(key, value);
        });

        // Send request
        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/invoice-report/generate-report', true);
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    const blob = new Blob([xhr.response], {
                        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    });
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `Formatted_Invoice_Data_${selectedSchool}.xlsx`;
                    document.body.appendChild(a);
                    a.click();
                    window.URL.revokeObjectURL(url);
                    messageDiv.style.display = 'none';
                } else {
                    const reader = new FileReader();
                    reader.onload = function () {
                        const errorMessage = reader.result;
                        alert(`Error generating report: ${errorMessage}`);
                        console.error('Error:', errorMessage);
                    };
                    reader.readAsText(xhr.response);
                }
            }
        };
        xhr.send(formData);
    }
});