var contentHeight = 0;
var sideHeight = 0;
var width = 0;

$(document).ready(function() {

    
    //When btn is clicked
    $("#btn-responsive-menu").click(function() {
        $("#responsive-menu").toggleClass("show");

    });
    
    function calculate(){
        contentHeight = $('#content').outerHeight();
        sideHeight = $("#sidebar-alt").outerHeight();
        width = $(window).width();
        
        if (width < 650 || sideHeight > contentHeight) {
            $("#sidebar-alt").css("height", 'auto');
	}
	else {
            $("#sidebar-alt").css("height", contentHeight);
	}
        
    }
    $(window).load(function() { calculate() }).resize(function() { calculate() });
    
});