// JavaScript Document

function updateTransp() {
	var error = 0;
	var valor = null;
//	alert(parseInt(document.frmSearchTranspSol.hiddenIntTransp.value + 1));
	for (var i=0; i < parseInt(document.frmSearchTranspColetaMaster.hiddenIntTransp.value); i++) {
		if (parseInt(document.frmSearchTranspColetaMaster.hiddenIntTransp.value) == 1) {
			if (!document.frmSearchTranspColetaMaster.transp.checked) {
				error++;
			} else {
				valor = document.frmSearchTranspColetaMaster.transp.value;	
			}
		} else {
			if (!document.frmSearchTranspColetaMaster.transp[i].checked) {
				error++;
			}	else {
				valor = document.frmSearchTranspColetaMaster.transp[i].value;	
			}
		}
	}
	if (error == parseInt(document.frmSearchTranspColetaMaster.hiddenIntTransp.value)) {
		alert("Por favor escolha uma transportadora");
		return
	} else {
		window.opener.frmEditSolicitacaoColetaDomiciliarAdm.cbTransp.value = valor;		
		window.close();
	}
}