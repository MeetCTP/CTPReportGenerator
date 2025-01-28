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
        const excelFileInput = document.getElementById('excel-file');
        const excelFile = excelFileInput.files[0];

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