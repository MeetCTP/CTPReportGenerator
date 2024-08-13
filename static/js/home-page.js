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