document.addEventListener('DOMContentLoaded', function () {
    const roleForm = document.getElementById('report-form');
    const reportSection = document.getElementById('report-section');
    const mailingListContainer = document.getElementById('mailing-list-container');
    const warningListContainer = document.getElementById('warning-list-container');
    const nonPaymentListContainer = document.getElementById('non-payment-list-container');

    roleForm.addEventListener('submit', async function (event) {
        event.preventDefault();

        const selectedRoles = Array.from(roleForm.elements['company_roles'])
            .filter(checkbox => checkbox.checked)
            .map(checkbox => checkbox.value);

        const startDate = roleForm.elements['start_date'].value;
        const endDate = roleForm.elements['end_date'].value;

        if (!startDate || !endDate) {
            alert('Please select a valid date range.');
            return;
        }

        try {
            // Request to generate the report and download it
            const reportResponse = await fetch('/report-generator/forty-eight-hour-warning/generate-report', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ company_roles: selectedRoles, start_date: startDate, end_date: endDate }),
            });

            console.log(JSON.stringify({ company_roles: selectedRoles, start_date: startDate, end_date: endDate }))

            if (!reportResponse.ok) {
                const errorResponse = await reportResponse.json();
                throw new Error(errorResponse.error || 'Network response was not ok');
            }

            const reportBlob = await reportResponse.blob();
            const reportUrl = window.URL.createObjectURL(reportBlob);
            const reportLink = document.createElement('a');
            reportLink.style.display = 'none';
            reportLink.href = reportUrl;
            reportLink.download = `48_Hour_Late_Conversions_Report_${new Date().toISOString().split('T')[0]}.xlsx`;
            document.body.appendChild(reportLink);
            reportLink.click();
            window.URL.revokeObjectURL(reportUrl);

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

            const mailingListData = await mailingListResponse.json();
            displayMailingList(mailingListData.mailing_list);
            displayWarningList(mailingListData.warning_list);
            displayNonPaymentList(mailingListData.non_payment_list);
            reportSection.style.display = 'block';

        } catch (error) {
            console.error('Error generating report:', error);
            alert(`Error generating report: ${error.message}`);
        }
    });

    function parseDate(dateString) {
        try {
            const trimmedDateString = dateString.trim();
            const [datePart, timePartWithPeriod] = trimmedDateString.split(/ +(?=\d)/); // Split on one or more spaces before time
    
            if (!datePart || !timePartWithPeriod) {
                console.error(`Invalid date string format: ${dateString}`);
                return 'Invalid Date';
            }
    
            const [timePart, period] = timePartWithPeriod.match(/(\d+:\d+)(AM|PM)/).slice(1, 3);
            const [month, day, year] = datePart.split('/').map(Number);
            let [hours, minutes] = timePart.split(':').map(Number);
    
            if (isNaN(month) || isNaN(day) || isNaN(year) || isNaN(hours) || isNaN(minutes)) {
                console.error(`Invalid date values in string: ${dateString}`);
                return 'Invalid Date';
            }
    
            if (period === 'PM' && hours < 12) {
                hours += 12;
            }
            if (period === 'AM' && hours === 12) {
                hours = 0;
            }
    
            const parsedDate = new Date(year, month - 1, day, hours, minutes);
            if (isNaN(parsedDate.getTime())) {
                console.error(`Parsed date is invalid: ${parsedDate}`);
                return 'Invalid Date';
            }
    
            return parsedDate;
        } catch (error) {
            console.error(`Error parsing date string "${dateString}": ${error.message}`);
            return 'Invalid Date';
        }
    }

    function displayMailingList(mailingList) {
        mailingListContainer.innerHTML = '';

        for (const [email, info] of Object.entries(mailingList)) {
            const div = document.createElement('div');
            div.innerHTML = `
                <input class="checkboxes" type="checkbox" id="${email}" name="providers" value="${email}">
                <label for="${email}">${info.name} (${email})</label>
            `;
            mailingListContainer.appendChild(div);
        }
    }

    function displayWarningList(warningList) {
        warningListContainer.innerHTML = '';

        warningList.forEach(item => {
            const div = document.createElement('div');
            const parsedDate = parseDate(item.AppStart);
            const formattedDate = parsedDate !== 'Invalid Date' ? parsedDate.toLocaleString() : 'Invalid Date';
            div.textContent = `${item.Provider} (${item.ProviderEmail}): ${formattedDate}`;
            warningListContainer.appendChild(div);
        });
    }

    function displayNonPaymentList(nonPaymentList) {
        nonPaymentListContainer.innerHTML = '';

        nonPaymentList.forEach(item => {
            const div = document.createElement('div');
            const parsedDate = parseDate(item.AppStart);
            const formattedDate = parsedDate !== 'Invalid Date' ? parsedDate.toLocaleString() : 'Invalid Date';
            div.textContent = `${item.Provider} (${item.ProviderEmail}): ${formattedDate}`;
            nonPaymentListContainer.appendChild(div);
        });
    }

    const mailingListForm = document.getElementById('mailing-list-form');
    mailingListForm.addEventListener('submit', async function (event) {
        event.preventDefault();

        const selectedProviders = Array.from(mailingListForm.elements['providers'])
            .filter(checkbox => checkbox.checked)
            .map(checkbox => {
                const email = checkbox.value;
                const name = checkbox.nextElementSibling.textContent;
        
                return {
                    Name: name.split(' (')[0],
                    Email: email,
                    Appointments: []  // Initialize as empty, will be populated by the backend
                };
            });

        if (Object.keys(selectedProviders).length === 0) {
            alert('Please select at least one provider.');
            return;
        }

        try {
            const response = await fetch('/report-generator/forty-eight-hour-warning/send-emails', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ selected_providers: selectedProviders }),
            });
            console.log('Selected Providers: ', selectedProviders)

            if (response.ok) {
                alert('Emails sent successfully!');
            } else {
                console.error('Failed to send emails:', response.statusText);
            }
        } catch (error) {
            console.error('Error sending emails:', error);
        }
    });
});