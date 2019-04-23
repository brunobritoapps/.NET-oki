// JavaScript Document

function updateTransp() {
	var error = 0;
	var valor = null;
//	alert(parseInt(document.frmSearchTranspSol.hiddenIntTransp.value + 1));
	for (var i=0; i < parseInt(document.frmSearchTranspPontoColeta.hiddenIntTransp.value); i++) {
		if (parseInt(document.frmSearchTranspPontoColeta.hiddenIntTransp.value) == 1) {
			if (!document.frmSearchTranspPontoColeta.transp.checked) {
				error++;
			} else {
				valor = document.frmSearchTranspPontoColeta.transp.value;	
			}
		} else {
			if (!document.frmSearchTranspPontoColeta.transp[i].checked) {
				error++;
			}	else {
				valor = document.frmSearchTranspPontoColeta.transp[i].value;	
			}
		}
	}
	if (error == parseInt(document.frmSearchTranspPontoColeta.hiddenIntTransp.value)) {
		alert("Por favor escolha uma transportadora");
		return
	} else {
		window.opener.frmPontoColetaAdm.cbTransp.value = valor;		
		window.close();
	}
}