$(document).ready(function() {

    var contentHeight = 0;
    
    //When btn is clicked
    $("#btn-responsive-menu").click(function() {
        $("#responsive-menu").toggleClass("show");

    });
    $(window).load(function() {
        contentHeight = $('#content').outerHeight();
        $("#sidebar-alt").css("height", contentHeight);
    });
    
    $(window).resize(function() {
        
        var width = $(window).width();
	if (width < 650) {
            $("#sidebar-alt").css("height", 'auto');
	}
	else {
            contentHeight = $('#content').outerHeight();
            $("#sidebar-alt").css("height", contentHeight);
	}
        
        
    });
    
});