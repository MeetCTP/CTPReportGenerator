document.addEventListener('DOMContentLoaded', function () {
    const roleForm = document.getElementById('report-form');
    const reportSection = document.getElementById('report-section');
    const mailingListContainer = document.getElementById('mailing-list-container');

    roleForm.addEventListener('submit', async function (event) {
        event.preventDefault();

        const selectedRoles = Array.from(roleForm.elements['company_roles'])
            .filter(checkbox => checkbox.checked)
            .map(checkbox => checkbox.value);

        const startDate = roleForm.elements['start_date'].value;
        const endDate = roleForm.elements['end_date'].value;

        if (selectedRoles.length === 0) {
            alert('Please select at least one company role.');
            return;
        }

        if (!startDate || !endDate) {
            alert('Please select a valid date range.');
            return;
        }

        try {
            // Request to generate the report
            const response = await fetch('/report-generator/forty-eight-hour-warning/generate-report', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ company_roles: selectedRoles, start_date: startDate, end_date: endDate }),
            });

            if (!response.ok) {
                const errorResponse = await response.json();
                throw new Error(errorResponse.error || 'Network response was not ok');
            }

            const blob = await response.blob();
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.style.display = 'none';
            a.href = url;
            a.download = `48_Hour_Late_Conversions_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);

            // Request to get the mailing list
            const mailingListResponse = await fetch('/report-generator/forty-eight-hour-warning/get-mailing-list', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ company_roles: selectedRoles, start_date: startDate, end_date: endDate }),
            });

            if (!mailingListResponse.ok) {
                const errorResponse = await mailingListResponse.json();
                throw new Error(errorResponse.error || 'Network response was not ok');
            }

            const mailingList = await mailingListResponse.json();
            displayMailingList(mailingList);
            reportSection.style.display = 'block';

        } catch (error) {
            console.error('Error generating report:', error);
            alert(`Error generating report: ${error.message}`);
        }
    });

    function displayMailingList(mailingList) {
        mailingListContainer.innerHTML = '';

        for (const [email, info] of Object.entries(mailingList)) {
            const div = document.createElement('div');
            div.innerHTML = `
                <input type="checkbox" id="${email}" name="providers" value="${email}">
                <label for="${email}">${info.name} (${email})</label>
            `;
            mailingListContainer.appendChild(div);
        }
    }

    const mailingListForm = document.getElementById('mailing-list-form');
    mailingListForm.addEventListener('submit', async function (event) {
        event.preventDefault();

        const selectedProviders = Array.from(mailingListForm.elements['providers'])
            .filter(checkbox => checkbox.checked)
            .reduce((obj, checkbox) => {
                const email = checkbox.value;
                obj[email] = { name: checkbox.nextElementSibling.textContent };
                return obj;
            }, {});

        if (Object.keys(selectedProviders).length === 0) {
            alert('Please select at least one provider.');
            return;
        }

        try {
            await fetch('/send_emails', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ selected_providers: selectedProviders }),
            });

            alert('Emails sent successfully!');
        } catch (error) {
            console.error('Error sending emails:', error);
        }
    });
});