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

        const startDate = document.getElementById('start-date').value;
        const endDate = document.getElementById('end-date').value;
        const excelFileInput = document.getElementById('excel-file');
        const excelFile = excelFileInput.files[0];

        const pgTypeCheckboxes = document.querySelectorAll('input[name="pg_type"]:checked');
        let pg_type = Array.from(pgTypeCheckboxes).map(cb => cb.value);

        const schoolChoice = document.querySelector('input[name="school-choice"]:checked');
        if (!schoolChoice) {
            alert('Please select a school.');
            return;
        }
        const selectedSchool = schoolChoice.value;

        if (!startDate || !endDate) {
            alert('Start and end dates are required.');
            return;
        }

        const formData = new FormData();
        formData.append('start_date', startDate);
        formData.append('end_date', endDate);
        formData.append('school', selectedSchool);
        if (pg_type) {
            pg_type.forEach(val => formData.append('pg_type', val))
        }
        if (excelFile) {
            formData.append('excel_file', excelFile);
        }

        console.log("FormData Contents:");
        formData.forEach((value, key) => {
            console.log(key, value);
        });

        // Send request
        const xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/school-matching/generate-report', true);
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
                    a.download = `Appointment_Match_Report_${selectedSchool}_${startDate}_${endDate}.xlsx`;
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