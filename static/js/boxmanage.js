//Box Add

const boxAdd = document.querySelector(".box-add");
const btnAdd = document.querySelector(".box-table .btn-add");

const maincontentManage = document.querySelector(".box-table");

btnAdd.addEventListener("click", () => {
    if (boxAdd.className === "box-add") boxAdd.classList.add("open");
    else boxAdd.classList.remove("open");
    boxEdit.classList.add("open");

    
});
btnAdd.addEventListener("click", () => {
    if (maincontentManage.className === "maincontentmanage")
        maincontentManage.classList.add("open");
    else maincontentManage.classList.remove("open");
});

//Box Edit
const boxEdit = document.querySelector(".box-edit");
const btnEdit = document.querySelector(".box-table .btn-edit");

btnEdit.addEventListener("click", () => {
    if (boxEdit.className === "box-edit") boxEdit.classList.add("open");
    else boxEdit.classList.remove("open");
    boxAdd.classList.add("open");
});
btnEdit.addEventListener("click", () => {
    if (maincontentManage.className === "maincontentmanage")
        maincontentManage.classList.add("open");
    else maincontentManage.classList.remove("open");
});