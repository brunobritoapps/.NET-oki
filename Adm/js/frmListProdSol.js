// JavaScript Document

function selecionar() {
	document.frmListProdSol.action.value = "SELECT";	
	document.frmListProdSol.submit();
}

function adicionar() {
	if (getSoma() > document.frmListProdSol.qtdrec.value || getSoma() < document.frmListProdSol.qtdrec.value ) {
		document.frmListProdSol.action.value = "";
		alert("Por favor preencha corretamente a quantidade de cartuchos entregues para cada produto!");
		return;
	} else {
		document.frmListProdSol.action.value = "ADD";	
		document.frmListProdSol.submit();		
	}
}

function getSoma() {
	var soma = 0;
	for (var i=0; i < parseInt(document.frmListProdSol.produtoslength.value); i++) {
		soma += parseInt(document.getElementById(i).value);	
	}
	return soma;
}
