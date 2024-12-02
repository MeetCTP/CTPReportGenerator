document.addEventListener("DOMContentLoaded", () => {
    function toggleFolder(event) {
        const folderButton = event.target;
        const folderContents = folderButton.nextElementSibling;

        const isAlreadyOpen = folderContents.classList.contains("show");

        const allContents = document.querySelectorAll(".folder-content");
        const allFolders = document.querySelectorAll(".folder");
        allContents.forEach(content => content.classList.remove("show"));
        allFolders.forEach(folder => folder.classList.remove("expanded"));

        if (!isAlreadyOpen) {
            folderContents.classList.add("show");
            folderButton.classList.add("expanded");
        }
    }

    const folderButtons = document.querySelectorAll(".folder");
    folderButtons.forEach(button => {
        button.addEventListener("click", toggleFolder);
    });
});