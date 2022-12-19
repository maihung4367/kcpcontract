$(document).ready(function($) {
    // $('table').hide();

    $("#statusSelector").change(function() {
        var selection = $(this).val();
        $("table")[selection ? "show" : "hide"]();

        if (selection) {
            $.each(
                $("#statusTable tbody tr"),
                function(index, item) {
                    $(item)[
                        $(item).is(":contains(" + selection + ")") ? "show" : "hide"
                    ]();
                }
            );
        }
    });
});