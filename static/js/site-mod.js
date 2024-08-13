function submitForm(formId) {
    var form = document.getElementById(formId);
    var formData = new FormData(form);

    var xhr = new XMLHttpRequest();
    xhr.open("POST", "/set/submit-form-data", true);

    xhr.onreadystatechange = function () {
        if (xhr.readyState === XMLHttpRequest.DONE) {
            if (xhr.status === 200) {
                console.log("Form submitted successfully.");
                console.log("Response:", xhr.responseText);
                form.reset();
            } else {
                console.error("Error submitting form:", xhr.responseText);
            }
        }
    };

    formData.forEach((value, key) => {
        console.log(key, value);
    });
    xhr.send(formData);
}