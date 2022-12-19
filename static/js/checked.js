$(document).ready(function() {
    $(".checkAll").on("click", function() {
        $(this)
            .closest("table")
            .find("tbody :checkbox")
            .prop("checked", this.checked)
            .closest("tr")
            .toggleClass("selected", this.checked);
    });

    $("tbody :checkbox").on("click", function() {
        // toggle selected class to the checkbox in a row
        $(this).closest("tr").toggleClass("selected", this.checked);

        // add selected class on check all
        $(this)
            .closest("table")
            .find(".checkAll")
            .prop(
                "checked",
                $(this).closest("table").find("tbody :checkbox:checked").length ==
                $(this).closest("table").find("tbody :checkbox").length
            );
    });
});