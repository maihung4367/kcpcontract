const sidebar = document.querySelector(".sidebar");
const maincontent = document.querySelector(".main-content");
const toggle = document.querySelector(".sidebar .toggle");

toggle.addEventListener("click", () => {
    if (sidebar.className === "sidebar") sidebar.classList.add("open");
    else sidebar.classList.remove("open");
});
toggle.addEventListener("click", () => {
    if (maincontent.className === "maincontent")
        maincontent.classList.add("open");
    else maincontent.classList.remove("open");
});

//Dropdown Info
function dropDownInfo() {
    document.getElementById("navDropdown").classList.toggle("show");
}

// Close the dropdown if the user clicks outside of it
window.onclick = function(event) {
    if (!event.target.matches(".dropbtn")) {
        var dropdowns = document.getElementsByClassName("dropdown-content");
        var i;
        for (i = 0; i < dropdowns.length; i++) {
            var openDropdown = dropdowns[i];
            if (openDropdown.classList.contains("show")) {
                openDropdown.classList.remove("show");
            }
        }
    }
};