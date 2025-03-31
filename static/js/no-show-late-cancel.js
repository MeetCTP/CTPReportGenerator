document.addEventListener('DOMContentLoaded', function () {
    var generateButton = document.getElementById('generate');
    var schoolSelect = document.getElementById('school');

    // Handle the display of the select dropdown based on radio button selection
    var radioButtons = document.querySelectorAll('input[name="reportType"]');
    radioButtons.forEach(radio => {
        radio.addEventListener('change', function () {
            if (document.getElementById('select').checked) {
                // Show the select dropdown if 'select' is chosen
                schoolSelect.style.display = 'block';
            } else {
                // Hide the select dropdown if anything else is selected
                schoolSelect.style.display = 'none';
            }
        });
    });

        if (generateButton) {
        generateButton.addEventListener('click', async function () {
            // Check which radio button is selected
            var reportType = document.querySelector('input[name="reportType"]:checked').value;

            if (reportType === 'multiple') {
                // If 'multiple' is selected, generate separate reports for each school
                await generateReportsForAllSchools();
            }
            else if (reportType === 'select') {
                // If 'select' is selected, generate a single report for the school that was selected
                await generateReportForSelectSchool();
            }
            else {
                // If 'single' is selected, generate one combined report for all schools
                await generateSingleReport();
            }
        });
    }
});

async function generateReportsForAllSchools() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var app_start = formData.get('app-start');
    var app_end = formData.get('app-end');
    var provider = formData.get('provider');
    var client = formData.get('client');

    var schools = [
        'School: Agora Cyber',
        'School: Commonwealth Charter Academy',
        'School: Achievement House Cyber Charter School',
        'School: Elwyn',
        'School: PA Distance Learning Charter',
        'School: Insight',
        'School: Reach Cyber',
        'School: PA Virtual Charter',
        'School: PA Leadership Charter School',
        'School: Delaware Co Intermediate Unit',
        'School: Central PA Digital Learning Foundation',
        'School: PA Cyber',
        'School: Gettysburg Montessori Charter'
    ];

    // Loop through each school, but wait for the report for one school to finish before moving to the next
    for (let school of schools) {
        await generateReportForSchool(school, app_start, app_end, provider, client, formData); // Wait for each school before moving to the next
    }
}

function generateReportForSchool(school, app_start, app_end, provider, client, formData) {
    return new Promise((resolve, reject) => {
        formData.set('school', school);  // Set the school dynamically

        var jsonData = JSON.stringify({
            app_start: app_start,
            app_end: app_end,
            provider: provider,
            client: client,
            school: school,
            single: 0
        });

        var xhr = new XMLHttpRequest();
        xhr.open('POST', '/report-generator/no-show-late-cancel/generate-report', true);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.responseType = 'blob';

        xhr.onreadystatechange = function () {
    if (xhr.readyState === XMLHttpRequest.DONE) {
        if (xhr.status === 200) {
            // Check if the response contains content
            if (xhr.response) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

                // Log the response content for debugging
                console.log("Received blob:", blob);

                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = school.replace(/[^a-zA-Z0-9]/g, '_') + '_No_Show_Late_Cancel_Report.xlsx'; // Ensure file name is valid
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                console.log('Report generated and downloaded for:', school);
                resolve(); // Resolve the promise once the report is downloaded
            } else {
                console.error("Received empty response or no data from the server.");
                reject(new Error('No data received from the server'));
            }
        } else {
            var reader = new FileReader();
            reader.onload = function () {
                var errorMessage = reader.result;
                console.error('Error generating report for', school, ':', errorMessage);
                reject(new Error('Error generating report for ' + school));
            };
            reader.readAsText(xhr.response);
        }
    }
};

        console.log('Generating report for school:', school);
        xhr.send(jsonData);
    });
}

async function generateReportForSelectSchool() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var app_start = formData.get('app-start');
    var app_end = formData.get('app-end');
    var provider = formData.get('provider');
    var client = formData.get('client');
    var selectedSchool = formData.get('school'); // Get selected school from the dropdown

    var jsonData = JSON.stringify({
        app_start: app_start,
        app_end: app_end,
        provider: provider,
        client: client,
        school: selectedSchool,
        single: 0
    });

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/no-show-late-cancel/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = selectedSchool.replace(/[^a-zA-Z0-9]/g, '_') + '_No_Show_Late_Cancel_Report.xlsx'; // Ensure file name is valid
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                console.log('Report generated and downloaded for:', selectedSchool);
            } else {
                var reader = new FileReader();
                reader.onload = function () {
                    var errorMessage = reader.result;
                    console.error('Error generating report for selected school:', errorMessage);
                };
                reader.readAsText(xhr.response);
            }
        }
    };

    console.log('Generating report for selected school:', selectedSchool);
    xhr.send(jsonData);
}

async function generateSingleReport() {
    var form = document.getElementById('report-form');
    var formData = new FormData(form);

    var app_start = formData.get('app-start');
    var app_end = formData.get('app-end');
    var provider = formData.get('provider');
    var client = formData.get('client');

    var schools = [
        'School: Agora Cyber',
        'School: Commonwealth Charter Academy',
        'School: Achievement House Cyber Charter School',
        'School: Elwyn',
        'School: PA Distance Learning Charter',
        'School: Insight',
        'School: Reach Cyber',
        'School: PA Virtual Charter',
        'School: PA Leadership Charter School',
        'School: Delaware Co Intermediate Unit',
        'School: Central PA Digital Learning Foundation',
        'School: PA Cyber',
        'School: Gettysburg Montessori Charter'
    ];

    var combinedSchools = schools.join(', ');

    var jsonData = JSON.stringify({
        app_start: app_start,
        app_end: app_end,
        provider: provider,
        client: client,
        school: combinedSchools,
        single: 1
    });

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/no-show-late-cancel/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'No_Show_Late_Cancel_Report_For_All_Schools.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                console.log('Single report generated and downloaded.');
            } else {
                var reader = new FileReader();
                reader.onload = function () {
                    var errorMessage = reader.result;
                    console.error('Error generating single report:', errorMessage);
                };
                reader.readAsText(xhr.response);
            }
        }
    };

    console.log(jsonData);
    xhr.send(jsonData);
}