/* When the user clicks on the button, 
toggle between hiding and showing the dropdown content */
function myFunction() {
  document.getElementById("navDropdown").classList.toggle("show");
}

// Close the dropdown if the user clicks outside of it
window.onclick = function(event) {
  if (!event.target.matches('.dropbtn')) {
      var dropdowns = document.getElementsByClassName("dropdown-content");
      var i;
      for (i = 0; i < dropdowns.length; i++) {
          var openDropdown = dropdowns[i];
          if (openDropdown.classList.contains('show')) {
              openDropdown.classList.remove('show');
          }
      }
  }
}

//btnDropdown
function btnDropdown() {
  document.getElementById("btnUpDropdown").classList.toggle("show");
}

// Close the dropdown if the user clicks outside of it
window.onclick = function(event) {
  if (!event.target.matches('.dropbtn')) {
      var dropdowns = document.getElementsByClassName("dropdown-content");
      var i;
      for (i = 0; i < dropdowns.length; i++) {
          var openDropdown = dropdowns[i];
          if (openDropdown.classList.contains('show')) {
              openDropdown.classList.remove('show');
          }
      }
  }
}

// Navlink
var header = document.getElementById("navLink");
var btns = header.getElementsByClassName("nav_item");
for (var i = 0; i < btns.length; i++) {
  btns[i].addEventListener("click", function() {
      var current = document.getElementsByClassName("active");
      if (current.length > 0) {
          current[0].className = current[0].className.replace(" active", "");
      }
      this.className += " active";
  });
}
// Pagination
var header = document.getElementById("pagination");
var btns = header.getElementsByClassName("numb-page");
for (var i = 0; i < btns.length; i++) {
  btns[i].addEventListener("click", function() {
      var current = document.getElementsByClassName("active");
      if (current.length > 0) {
          current[0].className = current[0].className.replace(" active", "");
      }
      this.className += " active";
  });
}

// Tel-input

const phoneInputField = document.querySelector("#numbPhone");
const phoneInput = window.intlTelInput(phoneInputField, {
  preferredCountries: ["vn", "us"],
  utilsScript: "https://cdnjs.cloudflare.com/ajax/libs/intl-tel-input/17.0.8/js/utils.js",
});

// Modal aLarm

const modalBtns = document.querySelectorAll('.js-modal-btn')
const modalOpen = document.querySelector('.js-modal-open')
const modalContainer = document.querySelector('.js-modal-container')
const modalClose = document.querySelector('.js-modal-close')

//Hàm hiển thị tab modal (thêm class open vào modal)
function showModalTab() {
  modalOpen.classList.add('open')
}

//Hàm đóng tab mocal (gỡ bỏ class open của modal)
function hideModalTab() {
  modalOpen.classList.remove('open')
}

//Dùng cho nhiều button mở modal (lặp qua từng thẻ button và nghe hành vi click)
for (const modalBtn of modalBtns) {
  modalBtn.addEventListener('click', showModalTab)
}

// Click vào icon "X" để close Tab Modal (nghe hành vi click vào button close)
modalClose.addEventListener('click', hideModalTab)

// Click vào ngoài màn hình sẽ đóng Tab Modal
modalOpen.addEventListener('click', hideModalTab)
modalContainer.addEventListener('click', function(event) {
  event.stopPropagation();
})

//   Responsive navbar