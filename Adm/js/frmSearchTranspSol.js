// JavaScript Document

function updateTransp() {
	var error = 0;
//	alert(parseInt(document.frmSearchTranspSol.hiddenIntTransp.value + 1));
	for (var i=0; i < parseInt(document.frmSearchTranspSol.hiddenIntTransp.value); i++) {
		if (parseInt(document.frmSearchTranspSol.hiddenIntTransp.value) == 1) {
			if (!document.frmSearchTranspSol.transp.checked) {
				error++;	
			}	
		} else {
			if (!document.frmSearchTranspSol.transp[i].checked) {
				error++;	
			}	
		}
	}
	if (error == parseInt(document.frmSearchTranspSol.hiddenIntTransp.value)) {
		alert("Por favor escolha uma transportadora");
		return
	} else {
		document.frmSearchTranspSol.submit();		
	}
}