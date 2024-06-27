function openTab(tabName, event) {
    var i;
    var tabContents = document.getElementsByClassName("tab-contents");
    var tabLinks = document.getElementsByClassName("tab-links");
    for (i = 0; i < tabContents.length; i++) {
        tabContents[i].classList.remove("active-tab");
    }
    for (i = 0; i < tabLinks.length; i++) {
        tabLinks[i].classList.remove("active-link");
    }
    document.getElementById(tabName).classList.add("active-tab");
    event.currentTarget.classList.add("active-link");
}

function openmenu() {
    document.getElementById("sidemenu").style.right = "0";
}

function closemenu() {
    document.getElementById("sidemenu").style.right = "-200px";
}

function handleSubmit(event) {
    event.preventDefault();
    var msg = document.getElementById("msg");
    msg.textContent = "Thank you for your message. We will get back to you soon.";
}

window.onload = function() {
    document.getElementById("see-more-btn").addEventListener("click", function() {
        var hiddenProjects = document.querySelectorAll(".project-item.hidden");
        hiddenProjects.forEach(function(project) {
            project.classList.remove("hidden");
        });
        this.style.display = "none";
    });
};
