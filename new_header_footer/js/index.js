$(function() {
	// carousel play pause //

	$('#main_carousel').carousel({
		interval : 4000,
		pause : "false"
	});
	$('#carouselButton').on('click', function() {
		if ($(this).hasClass("pause")) {
			$('#main_carousel').carousel('pause');
			$(this).removeClass("pause");
			$(this).addClass("play");
		} else {
			$('#main_carousel').carousel('cycle');
			$(this).removeClass("play");
			$(this).addClass("pause");
		}
	});

	// campanyText //

	$("[class^=companyImg]").on('mouseenter', function(){
		var $discription = $(this).find("[class^=contentsDescription]")
		if (touchDevice()) {
			setTimeout(function(){
				$discription.show();
			}, 300);
		} else {
			$discription.show();
		}
	});

	$("[class^=companyImg]").on('mouseleave', function(){
		var $discription = $(this).find("[class^=contentsDescription]")
		$discription.hide();
	});

	var touchDevice = function(){
		var agent = navigator.userAgent;
		if(agent.search(/iPhone/) != -1 || agent.search(/iPad/) != -1 || agent.search(/iPod/) != -1 || agent.search(/Android/) != -1){
			return true;
		}
	};


});
