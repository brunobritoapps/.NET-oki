// JavaScript Document

function validate() {
	var form = document.frmCategoriasAdm;
	
	if (form.txtDescricao.value == "") {
		alert("Preencha o campo Descrição da Categoria!");
		return;
	}
	if (form.txtQtdMinima.value == "") {
		alert("Preencha o campo Quantidade Mínima!");
		return;
	}
	if (isNaN(form.txtQtdMinima.value)) {
		alert("O campo Quantidade Mínima só aceita dados numéricos!");
		return;
	}
	form.submit();
}