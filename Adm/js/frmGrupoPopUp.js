// JavaScript Document

function UpdateGrupo() {
	var form = document.frmGrupoPopUp;

	for (var i=0; i < parseInt(form.hiddenIntGrupoCliente.value); i++) {
		if (form.radioThisGroup[i].checked) {
			window.location.href="frmGrupoPopUp.asp?cnpj="+form.cnpj.value+"&idgrupo="+form.radioThisGroup[i].value+"&action=alterar"
		}	
	}
}
