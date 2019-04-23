// JavaScript Document

function validate() {
	var form = document.frmCadProdutosAdm;
	if (form.cbGrupos.value == -1) {
		alert("Escolha um Grupo!");
		return;
	}
	if (form.txtDesc.value == "") {
		alert("Preencha o campo Descrição!");	
		return;
	}
	form.submit();
}