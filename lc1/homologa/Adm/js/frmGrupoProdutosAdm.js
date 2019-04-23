// JavaScript Document

function validate() {
	var form = document.frmGrupoProdutosAdm;
	if (form.txtDesc.value == "") {
		alert("Preencha o campo Descrição!");
		return;
	}
	form.submit();
}

function showOnClick() {
	var form = document.frmGrupoProdutosAdm;
	var bChecked = false;
	var cont = 0;
	
	for (var i=0; i < parseInt(form.hiddenIntProdutos.value); i++) {
		if (parseInt(form.hiddenIntProdutos.value) == 1) {
			if (form.radioIntProduto.checked) {cont = i;bChecked = true;}	
		} else {
			if (form.radioIntProduto[i].checked) {cont = i;bChecked = true;}
		}
	}
	(bChecked)?document.getElementById("cbgrupos").style.display = "block":document.getElementById("cbgrupos").style.display = "none";
}

function validateChangeListener() {
	document.frmGrupoProdutosAdm.action.value = "updategroup";
	document.frmGrupoProdutosAdm.submit();	
}