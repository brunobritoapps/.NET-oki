// JavaScript Document

function updateTransp() {
	var error = 0; 
    //	alert(parseInt(document.frmTransportadoraLc.hiddenIntTransp.value + 1));
	for (var i=0; i < parseInt(document.frmTransportadoraLc.hiddenIntTransp.value); i++) {
	    if (!document.frmTransportadoraLc.transp[i].checked) {
			error++;	
		}
	}
	if (error == parseInt(document.frmTransportadoraLc.hiddenIntTransp.value)) {
		alert("Por favor escolha uma transportadora");
		return
	} else {
	    document.frmTransportadoraLc.submit();
	}
}