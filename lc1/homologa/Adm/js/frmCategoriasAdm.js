// JavaScript Document

function validate() {
	var form = document.frmCategoriasAdm;
	
	if (form.txtDescricao.value == "") {
		alert("Preencha o campo Descri��o da Categoria!");
		return;
	}
	if (form.txtQtdMinima.value == "") {
		alert("Preencha o campo Quantidade M�nima!");
		return;
	}
	if (isNaN(form.txtQtdMinima.value)) {
		alert("O campo Quantidade M�nima s� aceita dados num�ricos!");
		return;
	}
	form.submit();
}