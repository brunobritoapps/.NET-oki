
//Peterson 17-5-2014
//Função para excluir a solicitação de resgate que ainda não foi aprovada;
//
function lcDelSR(element, element2) {
	var sid = element;
	var ssr = element2;
	
	if (confirm('Confirma a exclusão definitiva desta solicitação de resgate?')) { 
		window.location.href="frmsolicitacoesresgatecliente.asp?query=DELETE&id="+sid+"&sr=" + ssr
		return true;
	}
	else { 
		window.location.href="frmsolicitacoesresgatecliente.asp"
		return false;
	}

}

	
	