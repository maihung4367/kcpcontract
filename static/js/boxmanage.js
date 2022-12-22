var x = document.getElementById("boxEdit");
var y = document.getElementById("boxAdd");

function btnEdit() {
    x.style.display = "block";
    y.style.display = "none";
}

function btnAdd() {
    if (y.style.display === "none") {
        y.style.display = "block";
        x.style.display = "none";

    } else {
        y.style.display = "none";

    }
}

function btnClose() {
    y.style.display = "none";
    x.style.display = "none";
}