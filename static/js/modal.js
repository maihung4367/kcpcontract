/*========== Modal CreatePDF============*/

const modalBtns = document.querySelectorAll(".modal-btn-pdf");
const modalOpen = document.querySelector(".modal-open-pdf");
const modalContainer = document.querySelector(".modal-container-pdf");
const modalClose = document.querySelector(".modal-close-pdf");

//Hàm hiển thị tab modal (thêm class open vào modal)
function showModalTab() {
    modalOpen.classList.add("open");
}

//Hàm đóng tab mocal (gỡ bỏ class open của modal)
function hideModalTab() {
    modalOpen.classList.remove("open");
}

//Dùng cho nhiều button mở modal (lặp qua từng thẻ button và nghe hành vi click)
for (const modalBtn of modalBtns) {
    modalBtn.addEventListener("click", showModalTab);
}

// Click vào icon "X" để close Tab Modal (nghe hành vi click vào button close)
modalClose.addEventListener("click", hideModalTab);

// Click vào ngoài màn hình sẽ đóng Tab Modal
modalOpen.addEventListener("click", hideModalTab);
modalContainer.addEventListener("click", function(event) {
    event.stopPropagation();
});


/*========== Modal Info ==========*/
$(document).ready(function() {
    var modal = $(".modal-info");
    var info = $(".info");
    var span = $(".close-info");

    info.click(function() {
        modal.show();
    });

    span.click(function() {
        modal.hide();
    });

    $(window).on("click", function(e) {
        if ($(e.target).is(".modal-info")) {
            modal.hide();
        }
    });
});
