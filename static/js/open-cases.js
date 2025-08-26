document.addEventListener('DOMContentLoaded', function () {
    const generateButton = document.getElementById('generate');

    if (generateButton) {
        generateButton.addEventListener('click', function () {
            generateReport();
        });
    }

    function generateReport() {
        let messageDiv = document.getElementById('loading-message');
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
            logo.style.width = '100px';
            logo.style.marginBottom = '50px';

            // Add logo first, then message
            messageDiv.appendChild(logo);
            const messageText = document.createElement('p');
            messageText.textContent = "Generating the report, please be patient. This might take a few minutes...";
            messageDiv.appendChild(messageText);

            document.body.appendChild(messageDiv);
        }

        const excelFileInput = document.getElementById('excel-file');
        const file = excelFileInput.files[0];

        const formData = new FormData();
        if (file) {
            formData.append('file', file);
        }

        console.log("FormData Contents:");
        formData.forEach((value, key) => {
            console.log(key, value);
        });

        // Send request
        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/open-cases/generate-report', true);
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                if (xhr.status === 200) {
                    // Detect response file type
                    const contentType = xhr.getResponseHeader("Content-Type");
                    let fileExtension = "xlsx"; // default
                    if (contentType && contentType.includes("csv")) {
                        fileExtension = "csv";
                    }

                    const blob = new Blob([xhr.response], { type: contentType });
                    const url = window.URL.createObjectURL(blob);

                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `Formatted_Open_Cases.xlsx`;
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