var contentHeight = 0;
var sideHeight = 0;
var width = 0;
var height = 0;
var slides = $('.slides');

$(document).ready(function() {
    //When btn is clicked
    $("#btn-responsive-menu").click(function() {
        $("#responsive-menu").slideToggle();
    });
    
    $(".tiptip").tipTip({activation: 'click', maxWidth: "auto", keepAlive: true});
    
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
    calculate(); 
    $('#content').bind('DOMSubtreeModified', function(){
        calculate();
        console.log('dom changed!');
    });
    $(window).load(function() {
        calculate(); 
         $('.blueberry').blueberry({
             pager: false
         });
    }).resize(function() { 
        if (width >= 650) {
            $("#responsive-menu").hide();
            if (slides.height() == 0) {
                console.log('slides height 0');
                $('.blueberry').blueberry({
                    pager: false
                });
            }
        }
        calculate(); 
    });
    
    if ($('#back-to-top').length) {
        var scrollTrigger = 100, // px
            backToTop = function () {
                var scrollTop = $(window).scrollTop();
                if (scrollTop > scrollTrigger) {
                    $('#back-to-top').addClass('show');
                } else {
                    $('#back-to-top').removeClass('show');
                }
            };
        backToTop();
        $(window).on('scroll', function () {
            backToTop();
        });
        $('#back-to-top').on('click', function (e) {
            e.preventDefault();
            $('html,body').animate({
                scrollTop: 0
            }, 700);
        });
    }
});