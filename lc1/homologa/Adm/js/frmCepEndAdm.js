// JavaScript Document

function windowLocationFind(url) {
	var find = true;
	var busca = document.frmCepEndAdm.search.value;
	var por = document.frmCepEndAdm.cbBusca.value;
	window.location.href = url + "/adm/frmCepEndAdm.asp?find=true&search="+busca+"&changetype="+por;
}
