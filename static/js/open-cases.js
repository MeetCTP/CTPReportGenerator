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

        const ccaFile = document.getElementById('cca-file').files[0];
        const agoraFile = document.getElementById('agora-file').files[0];
        const insightFile = document.getElementById('insight-file').files[0];
        const otherFile = document.getElementById('other-file').files[0];

        if (!ccaFile || !agoraFile || !insightFile) {
            alert("Please upload all required report files before continuing.");
            messageDiv.style.display = 'none';
            return;
        }

        const formData = new FormData();
        formData.append('cca_file', ccaFile);
        formData.append('agora_file', agoraFile);
        formData.append('insight_file', insightFile);
        formData.append('other_file', otherFile);

        console.log("FormData Contents:");
        formData.forEach((value, key) => {
            console.log(key, value);
        });

        // Send request
        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/open-cases/generate-report', true);
        xhr.responseType = 'json'; // Expect JSON now, not blob

        xhr.onreadystatechange = function () {
            if (xhr.readyState === XMLHttpRequest.DONE) {
                messageDiv.style.display = 'none'; // Hide loading spinner either way

                if (xhr.status === 200 && xhr.response && xhr.response.message) {
                    alert(xhr.response.message); // ✅ Success
                    console.log("✅ Success:", xhr.response.message);
                } else if (xhr.response && xhr.response.error) {
                    alert(`Error: ${xhr.response.error}`); // ❌ API error message
                    console.error("Error:", xhr.response.error);
                } else {
                    alert("Unexpected error occurred. Please try again.");
                    console.error("Unexpected response:", xhr.response);
                }
            }
        };

        xhr.onerror = function () {
            messageDiv.style.display = 'none';
            alert("Network error: Could not reach server.");
        };
        
        xhr.send(formData);
    }
});