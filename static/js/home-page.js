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
                    if (tableName == 'Notifications') {
                        this.closest('.item-container-other').remove();
                    } else if (tableName == 'WeeklyQAResponses') {
                        this.closest('.response-item').remove();
                    } else {
                        this.closest('.item-container').remove();
                    }
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

document.addEventListener("DOMContentLoaded", function() {
    const modal = document.getElementById("imageModal");
    const modalImg = document.getElementById("modalImage");
    const closeModalButton = document.getElementById("closeModal");
    let zoomLevel = 1;
    let isDragging = false;
    let startX, startY, imgPosX = 0, imgPosY = 0;

    function openModal(imageSrc) {
        modal.style.display = "block";
        modalImg.src = imageSrc;
        zoomLevel = 1;
        imgPosX = 0;
        imgPosY = 0;
        modalImg.style.transform = `scale(${zoomLevel}) translate(0px, 0px)`;
    }

    function closeModal() {
        modal.style.display = "none";
    }

    function zoom(event) {
        const zoomStep = 0.1;
        if (event.deltaY < 0) {
            zoomLevel += zoomStep;
        } else if (event.deltaY > 0 && zoomLevel > 0.4) {
            zoomLevel -= zoomStep;
        }
        updateTransform();
        event.preventDefault();
    }

    function updateTransform() {
        modalImg.style.transform = `scale(${zoomLevel}) translate(${imgPosX}px, ${imgPosY}px)`;
    }

    document.querySelectorAll(".news-image").forEach(image => {
        image.addEventListener("click", () => openModal(image.src));
    });

    closeModalButton.addEventListener("click", closeModal);
    modal.addEventListener("click", function(event) {
        if (event.target === modal) closeModal();
    });

    modalImg.addEventListener("wheel", zoom);

    modalImg.addEventListener("mousedown", (event) => {
        if (zoomLevel > 1) {
            isDragging = true;
            startX = event.clientX - imgPosX;
            startY = event.clientY - imgPosY;
            modalImg.style.cursor = "grab";
            event.preventDefault();
        }
    });

    document.addEventListener("mousemove", (event) => {
        if (isDragging) {
            imgPosX = event.clientX - startX;
            imgPosY = event.clientY - startY;
            updateTransform();
        }
    });

    document.addEventListener("mouseup", () => {
        isDragging = false;
        modalImg.style.cursor = "default";
    });
});