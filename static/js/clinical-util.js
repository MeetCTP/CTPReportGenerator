document.addEventListener('DOMContentLoaded', function () {
    // Get form and button elements
    const form = document.getElementById('report-form');
    const generateButton = document.getElementById('generate');

    // Set up event listener for the generate report button
    generateButton.addEventListener('click', function () {
        // Get the form values
        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;
        const provider = document.getElementById('provider').value;

        // Validate form data
        if (!startDate || !endDate || !provider) {
            alert('Please fill in all fields');
            return;
        }

        // Prepare the data to send in the POST request
        const formData = {
            start_date: startDate,
            end_date: endDate,
            provider: provider
        };

        // Send the form data to the backend
        fetch('/report-generator/clinical-util-tracker/generate-report', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(formData)
        })
        .then(response => {
            if (response.ok) {
                return response.blob(); // Get the file as a blob
            } else {
                return response.json().then(errData => {
                    throw new Error(errData.error || 'Error generating the report');
                });
            }
        })
        .then(blob => {
            // Create a link to download the file
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Clinical_Util_Tracker_${provider}_${startDate}_${endDate}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url); // Clean up the object URL
            document.body.removeChild(a);
        })
        .catch(error => {
            alert(`Failed to generate report: ${error.message}`);
        });
    });
});