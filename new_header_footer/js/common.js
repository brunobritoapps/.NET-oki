$(function() {
	var lang = $('html').attr('lang');
	if ( lang === 'ja') {
		$('body').css('font-family', '"メイリオ","Meiryo","ＭＳ Ｐゴシック", "arial", "helvetica", "sans-serif"');
	} else if ( lang === 'zh') {
		$('body').css('font-family', '"SimHei", "NSimSun", "sans-serif"');
	}


	// ipad viewport change => ipad 縦のみviewport変更
	var agent = navigator.userAgent;
	if( agent.search(/iPad/) != -1 && window.innerHeight > window.innerWidth ){
		$('meta[name=viewport]').remove();
		$('head').prepend('<meta name="viewport" content="width=960px">');
	}

	var touchDevice = function(){
		if(agent.search(/iPhone/) != -1 || agent.search(/iPad/) != -1 || agent.search(/iPod/) != -1 || agent.search(/Android/) != -1){
			return true;
		}
	};


	// to top icon fade in fade out
	$(window).on("scroll", function() {
		// トップから350px以上スクロールしたら
		if ($(this).scrollTop() > 350) {
			// ページトップのリンクをフェードインする
			$(".pageTop").fadeIn();
		} else { // それ以外は
			// ページトップのリンクをフェードアウトする
			$(".pageTop").fadeOut();
		}
	});

	// smooth scroll
	// ページ内リンクをクリックするとリンク先へスムーズに移動する
	// クリックしたリンク先が#または空のときは移動しない
	$('a[href^=#]').click(function() {
		// スクロールスピード
		var speed = 500;
		// クリックしたリンク先を保存
		var href = $(this).attr("href");
		if (!(href === "#" || href === "")){
			var position = $(href).offset().top;
			$("html, body").animate({
				scrollTop : position
			}, speed, "swing");
		}
		return false;
	});


	// country selector //

	$(".countryName a").on('click', function() {
		var $countryBox = $(".countryWrapper, .countrySelectBox");
		if ( $countryBox.is(':visible') ) {
			$countryBox.hide();
			$(".countryName").removeClass("active");
		} else {
			$countryBox.show();
			$(".countryName").addClass("active");
		}
	});
	
	$(document).on('click', function(e){
		var $countryBox = $(".countryWrapper, .countrySelectBox");
		if ( $countryBox.is(':visible') && !$.contains($('.countryWrapper')[0], e.target)
			 && !$.contains($('.countryBackground2')[0], e.target) ) {
			$countryBox.hide();
			$(".countryName").removeClass("active");
		}
	});


	// mega menu //
	var $megaInner = $(".megaInner");

	$("#headerNav2 .menubar").on('mouseenter', function() {
		$(this).addClass('active');

		var menu = $(this).children(".megaWrapper");
	});


	// アクセシビリティ対応
	$('#headerNav2 > .menubar > a').on('mouseenter', function(e) {
		e.preventDefault();
		var $thisMenu = $(this).closest("div").children(".megaWrapper");
		if ($thisMenu.is(":visible") ) {
			$(this).closest("div").removeClass('active');
		} else {
			$(this).closest("div").addClass('active');
			$thisMenu.show();
		}
	});

	$(".menubar").on('mouseleave', function() {
		$(this).removeClass('active');
	});

	// searchBox //

	$(".searchIconSm a").on('click', function(e){
		e.preventDefault();
		$(".searchBoxSm").show();
		$("#searchSm").focus();
	});
	$(".searchIconXs a").on('click', function(e){
		e.preventDefault();
		$(".searchBoxXs").show();
		$("#searchXs").focus();
	});
	$(".closeSearch").on('click', function(e){
		e.preventDefault();
		$(".searchBox").hide();
	});
	$(document).on('click', function(e){
		if ( e.target !== $('#searchSm')[0] && e.target !== $('#searchXs')[0] ){
			$(".searchBoxSm").hide();
		}
	});
	$(document).on('click', function(e){
		if ( e.target !== $('#searchXs')[0] ){
			$(".searchBoxXs").hide();
		}
	});
	$(window).resize(function() {
		if ( !touchDevice() && $(".searchBox").is(':visible') ) {
			$(".searchBox").hide();
		}
	});


	//IE使用バージョン取得
	var getIeVer = function() {
		var undef, v = 3, div = document.createElement('div');
		while (
			div.innerHTML = '<!--[if gt IE '+(++v)+']><i></i><![endif]-->',
			div.getElementsByTagName('i')[0]
		);
		return v> 4 ? v : undef;
	}

	//IE9以下の場合はplaceholderの値をvalueに変換
	if(getIeVer() <= 9) {
		var searchText = $("#searchSm").attr("placeholder");
		$("#searchSm").val(searchText);
		$("#searchSm").css("color", "#999");
		$("#searchSm").focus(function() {
			if($(this).val() == searchText) {
				$(this).val("");
				$(this).css("color", "#000");
			}
		}).blur(function() {
			if($(this).val() == "") {
				$(this).val(searchText);
				$("#searchSm").css("color", "#999");
			}
		});
	}


	//cookie alert
	var messageDisplay = "block";
	var ca = document.cookie.split(";");
	for(var index=0;index< ca.length;index++){
		var c = ca[index];
		if(c.indexOf("close_cookie_access_message") != -1){
			messageDisplay = "none";
			break;
		}
	}
	$("#cookie_access_message_holder").css("display",messageDisplay);
	
	$("a[href^='#']").on("click", closeCookieAccessMessage);
		$(document).on("click", closeCookieAccessMessage);

	function closeCookieAccessMessage() {
		if ($("#cookie_access_message_holder").is(":visible")) {
			$("#cookie_access_message_holder").css("display","none");
			var lang = document.location.pathname;
			lang = lang.split('/');
			var cookieTxt = "close_cookie_access_message=viewed; path=/" + lang[1] + "/; max-age=" + 15552000;
			document.cookie= cookieTxt;
		}
	}

	// toggle accordion
	$(".specToggle").on('click', function(e){
		e.preventDefault();
		var $this = $(this)
		var $detail = $this.next(".specDetail");
		if ($detail.hasClass("active")){
			$detail.slideUp("normal", function(){
				$detail.removeClass("active");
				$this.children("div").removeClass("dtlOn").addClass("dtlOff");
			});
		} else {
			$detail.slideDown("normal", function(){
				$detail.addClass("active");
				$this.children("div").removeClass("dtlOff").addClass("dtlOn");
			});
		}
	});


	$(".advancedSearch a").on('click', function() {
		if ( $(".searchItemBox").is(':visible') ) {
			$(".searchItemBox").slideUp();
		} else {
			$(".searchItemBox").slideDown();
		}
	});


	//テキストを長すぎる場合後略する関数
	function clipText(selector, maxCount){
		$(selector).each(function(index, element){
			var text = $(element).text().trim();
			text = text.replace(/\s+/g, ' ');
			var len = $(element).text().length;
			if(len > maxCount){
				text = text.substring(0, maxCount-3);
				text = text + "...";
				$(element).html(text);
			}
		});
	}
	clipText(".tabItemContent p",140);
	clipText(".subItemContent p",115);
	clipText(".subItemContent2",120);
	clipText(".selectorText p",95);
	clipText(".selectorTable td span",60);


	// file upload
	// see:http://duckranger.com/2012/06/pretty-file-input-field-in-bootstrap/
	$('.jsSelectFile a').on('click', function() {
		$leFile = $(this).closest(".fileUpForm").find('input[class=jsLefile]');
		$leFile.trigger('click');
	});

	$('input[class=jsLefile]').on('change', function() {
		$selectedFileName = $(this).next(".input-append").children(".jsSelectedFile");
		var filePath = $(this).val();
		$selectedFileName.val(filePath);
		if ( filePath.match(/fakepath/) ) {
			res = filePath.split("\\");
			$selectedFileName.val( res[res.length-1] );
		}
	});
	
	$('.jsSelectedFile').on('focus', function() {
		$leFile = $(this).closest(".fileUpForm").find('input[class=jsLefile]');
		$leFile.trigger('click');
	});
	
	$('.jsSelectedFile').on('blur', function() {
		$selectFileName = $(this).closest(".input-append").prev("input[class=jsLefile]");
		if ($(this).val() === "") {
			$selectFileName.val($(this).val());
		}
	});


	//LinkListの高さを揃える
	$('.linkHeight').matchHeight();


});
