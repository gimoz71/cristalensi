$(document).ready(function() {

    //When btn is clicked
    $("#btn-responsive-menu").click(function() {
        $("#responsive-menu").toggleClass("show");

    });
    $(window).load(function() {
        var contentHeight = $('#wrap').outerHeight();
        $("#sidebar-alt").css("height", contentHeight);
    });
    
    $(window).resize(function() {
        contentHeight = $('#main-content').outerHeight();
        $("#sidebar-alt").css("height", contentHeight);
    });
    
});