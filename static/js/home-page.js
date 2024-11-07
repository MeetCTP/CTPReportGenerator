function submitResponse(button) {
    var form = button.closest('form');
    var questionId = form.getAttribute('data-question-id');
    var responseBody = form.querySelector('textarea[name="response"]').value;

    fetch('/submit-response', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            questionId: questionId,
            responseBody: responseBody
        }),
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            // Update the responses section
            var responsesDiv = form.previousElementSibling;
            var newResponse = document.createElement('p');
            newResponse.textContent = `${responseBody} - ${data.createdBy} at ${new Date().toLocaleString()}`;
            responsesDiv.appendChild(newResponse);
            form.reset();
        } else {
            console.error('Error submitting response:', data.message);
        }
    })
    .catch(error => console.error('Error submitting response:', error));
}

document.addEventListener('DOMContentLoaded', function() {
    document.querySelectorAll('.delete-btn').forEach(button => {
        button.addEventListener('click', async function(event) {
            const itemId = this.getAttribute('data-id');
            const tableName = this.getAttribute('name');

            const confirmed = confirm('Are you sure you want to delete this item?');
            if (!confirmed) return;

            try {
                const response = await fetch(`/delete-home-item`, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify({ id: itemId, table: tableName }),
                });

                console.log(JSON.stringify({ id: itemId, table: tableName}))

                if (response.ok) {
                    this.closest('.item-container').remove();
                    alert('Item deleted successfully.');
                } else {
                    alert('Failed to delete the item.');
                }
            } catch (error) {
                console.error('Error:', error);
                alert('An error occurred. Please try again.');
            }
        });
    });
});

document.addEventListener('DOMContentLoaded', function() {
    var searchButton = document.getElementById('search');
    if (searchButton) {
        searchButton.addEventListener('click', function() {
            serciveCodeSearch();
        });
    }
});

function serciveCodeSearch() {
    var query = document.getElementById('query').value;

    fetch('/submit-search', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({
            query: query
        }),
    })
    .then(response => response.json())
    .then(data => {
        var resultsContainer = document.getElementById('results');
        resultsContainer.innerHTML = '';

        if (data.results && data.results.length > 0) {
            data.results.forEach(result => {
                var resultItem = document.createElement('div');
                resultItem.classList.add('result-item');
                resultItem.innerHTML = `
                    <p style="font-size: 1rem;"><strong>Code:</strong> ${result[0]}</p>
                    <pstyle="font-size: 1rem;"><strong>Description:</strong> ${result[1]}</p>
                `;
                resultsContainer.appendChild(resultItem);
            });
        } else {
            var noResultsMessage = document.createElement('p');
            noResultsMessage.textContent = data.message || 'No results found.';
            resultsContainer.appendChild(noResultsMessage);
        }
    })
    .catch(error => {
        console.error('Error:', error);
        var resultsContainer = document.getElementById('results');
        resultsContainer.innerHTML = '<p>An error occurred while fetching results.</p>';
    });
}