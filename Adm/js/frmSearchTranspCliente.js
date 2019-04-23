// JavaScript Document

function updateTransp() {
	var error = 0;
//	alert(parseInt(document.frmSearchTranspSol.hiddenIntTransp.value + 1));
	for (var i=0; i < parseInt(document.frmSearchTranspCliente.hiddenIntTransp.value); i++) {
		if (!document.frmSearchTranspCliente.transp[i].checked) {
			error++;	
		}
	}
	if (error == parseInt(document.frmSearchTranspCliente.hiddenIntTransp.value)) {
		alert("Por favor escolha uma transportadora");
		return
	} else {
		document.frmSearchTranspCliente.submit();		
	}
}