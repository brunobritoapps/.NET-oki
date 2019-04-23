$(function(){
	$('.bxslider').bxSlider({
		buildPager: function(slideIndex){
			/*$.each($(".slideImg").data('images'), function(i, img){
				if (slideIndex === i) {
					//console.log('<img src="' + img + '">')
					//return '<img src="' + img + '">';
					console.log(slideIndex + $(".slideImg").data("images")[i])
				}
			});*/
			switch(slideIndex){
				case 0:
					return '<img src="' + $(".slideImg").data("images")[0] + '">';
				case 1:
					return '<img src="' + $(".slideImg").data("images")[1] + '">';
				case 2:
					return '<img src="' + $(".slideImg").data("images")[2] + '">';
				case 3:
					return '<img src="' + $(".slideImg").data("images")[3] + '">';
				case 4:
					return '<img src="' + $(".slideImg").data("images")[4] + '">';
			}
		}
	});
});