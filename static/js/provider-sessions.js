document.addEventListener('DOMContentLoaded', function() {
    const supervisors = ['Service Supervisor: Cari.Tomczyk', 'Service Supervisor: Nicole M. Nies', 'Un-Assigned'];
    const statuses = ['Deleted', 'Un-Converted', 'Converted/Canceled', 'Cancelled', 'Converted/Deleted', 'Converted'];

    const supervisorsList = document.getElementById('supervisors-list');
    const statusList = document.getElementById('status-list');

    // Populate supervisors checkboxes
    supervisors.forEach((supervisor, index) => {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `supervisor-${index}`;
        checkbox.className = 'checkboxes';
        checkbox.value = supervisor;
        const label = document.createElement('label');
        label.htmlFor = `supervisor-${index}`;
        label.textContent = supervisor;

        supervisorsList.appendChild(checkbox);
        supervisorsList.appendChild(label);
        supervisorsList.appendChild(document.createElement('br'));
    });

    // Populate status checkboxes
    statuses.forEach((status, index) => {
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.id = `status-${index}`;
        checkbox.className = 'checkboxes';
        checkbox.value = status;
        const label = document.createElement('label');
        label.htmlFor = `status-${index}`;
        label.textContent = status;

        statusList.appendChild(checkbox);
        statusList.appendChild(label);
        statusList.appendChild(document.createElement('br'));
    });
});

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

    const form = document.getElementById('report-form');
    const formData = new FormData(form);

    const range_start = formData.get('range-start');
    const range_end = formData.get('range-end');
    const selectedSupervisors = Array.from(document.querySelectorAll('#supervisors-list input:checked')).map(el => el.value);
    const selectedStatuses = Array.from(document.querySelectorAll('#status-list input:checked')).map(el => el.value);

    const jsonData = JSON.stringify({
        range_start: range_start,
        range_end: range_end,
        supervisors: selectedSupervisors,
        status: selectedStatuses
    });

    var xhr = new XMLHttpRequest();
    xhr.open('POST', '/report-generator/provider-sessions/generate-report', true);
    xhr.setRequestHeader('Content-Type', 'application/json');
    xhr.responseType = 'blob';

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                var blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                var url = window.URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = 'Provider_Sessions_Report.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                messageDiv.style.display = 'none';
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

    console.log(jsonData);
    xhr.send(jsonData);
}