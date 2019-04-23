$(function(){
	$('#language').change(function() {
		var selectLanguage = $('#language').val();
		var currentUrl = window.location.href;
		var langParam = currentUrl.match(/lang=[0-9a-zA-Z]*/);
		if (langParam == null){ 
			if (currentUrl.indexOf('?') > 0) {
				window.location.href = currentUrl + '&lang=' + selectLanguage;
			} else {
				window.location.href = currentUrl + '?lang=' + selectLanguage;
			}
		} else {
			window.location.href = currentUrl.replace(langParam, 'lang=' + selectLanguage);
		}
	});

	$('input[name="supportManualsTab"]').change(function() {
		var downloadManualUrl = $("input[name='supportManualsTab']:checked").val();
		$('#supportManualDownload a').attr('href', downloadManualUrl);
	});

	$('#softwareOs').change(function() {
		softwareDropDownListReload();
	});

	$('#softwareLanguage').change(function() {
		softwareDropDownListReload();
	});

	function softwareDropDownListReload() {
		var selectOs = $('#softwareOs').val();
		var selectLanguage = $('#softwareLanguage').val();
		var currentUrl = window.location.href;
		var langParam = currentUrl.match(/lang=[0-9a-zA-Z]*/);
		var osParam = currentUrl.match(/os=[0-9a-zA-Z]*/);

		if (langParam == null && osParam == null) {
			if (currentUrl.indexOf('?') > 0) {
				if (selectOs != "") {
					window.location.href = currentUrl + '&os=' + selectOs + '&lang=' + selectLanguage;
				} else {
					window.location.href = currentUrl + '&lang=' + selectLanguage;
				}
			} else {
				if (selectOs != "") {
					window.location.href = currentUrl + '?os=' + selectOs + '&lang=' + selectLanguage;
				} else {
					window.location.href = currentUrl + '?lang=' + selectLanguage;
				}
			}
		} else if (langParam != null && osParam == null) {
			if (selectOs != "") {
				window.location.href = currentUrl.replace(langParam, 'os=' + selectOs + '&lang=' + selectLanguage);
			} else {
				window.location.href = currentUrl.replace(langParam, 'lang=' + selectLanguage);
			}
		} else if (langParam == null && osParam != null) {
			if (selectOs != "") {
				window.location.href = currentUrl.replace(osParam, 'os=' + selectOs + '&lang=' + selectLanguage);
			} else {
				window.location.href = currentUrl.replace(osParam, 'lang=' + selectLanguage);
			}
		} else if (langParam != null && osParam != null) {
			if (selectOs != "") {
				window.location.href = currentUrl.replace(osParam, 'os=' + selectOs).replace(langParam, 'lang=' + selectLanguage);
			} else {
				if (currentUrl.match(/&os=[0-9a-zA-Z]*/) != null) {
					window.location.href = currentUrl.replace('&' + osParam, '').replace(langParam, 'lang=' + selectLanguage);
				} else {
					window.location.href = currentUrl.replace(osParam + '&', '').replace(langParam, 'lang=' + selectLanguage);
				}
			}
		}
	}
});
