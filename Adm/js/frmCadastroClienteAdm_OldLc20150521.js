// JavaScript Document

function windowLocationFind(url) {
	var find = true;
	var busca = document.frmCadastroClienteAdm.txtFindCliente.value;
	var por = document.frmCadastroClienteAdm.typeFindCliente.value;
	var status = document.frmCadastroClienteAdm.cbStatusFindCliente.value;
	var categoria = document.frmCadastroClienteAdm.cbCategoriasFindCliente.value;
	var grupos = document.frmCadastroClienteAdm.grupos.value;
	window.location.href = url + "/adm/frmCadastroClienteAdm.asp?find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria+"&grupos="+grupos;
}


function aprovar() {
	if (document.frmCadastroClienteAdm.hiddenActionIsColetaDomiciliar.value == 1) 
	{
		if (document.frmCadastroClienteAdm.cbTransp.value == "-1") 
		{
			alert("Selecione uma Transportadora!");
			return;
		}
	}
	document.frmCadastroClienteAdm.hiddenActionManagerProve.value = 'true';
	document.frmCadastroClienteAdm.hiddenActionForm.value = "APROVAR";
	document.frmCadastroClienteAdm.submit();
}

function reprovar() {
	document.frmCadastroClienteAdm.hiddenActionManagerProve.value = 'false';		
	document.frmCadastroClienteAdm.hiddenActionForm.value = "REPROVAR";
	document.frmCadastroClienteAdm.submit();
}

function validar() {
	if (document.frmCadastroClienteAdm.radioPessoa[0].checked) {
		if (document.frmCadastroClienteAdm.txtNomeCliente.value == "") {
			alert("Preencha o campo Nome!");	
			return;
		}
		validaCPF();	
	} else {
		if (document.frmCadastroClienteAdm.radioPessoa[1].checked) {
			if (document.frmCadastroClienteAdm.txtNomeFantasiaCliente.value == "") {
				alert("Preencha o campo Nome Fantasia!");
				return;
			}
			if (document.frmCadastroClienteAdm.txtCNPJCliente.value == "") {
				alert("Preencha o campo CNPJ!");
				return;
			}
			validaCnpj();
		}
	}
	if (document.frmCadastroClienteAdm.txtDDDCliente.value == "") {
		alert("Preencha o campo DDD!");
		return;
	}
	if (isNaN(document.frmCadastroClienteAdm.txtDDDCliente.value)) {
		alert("Preencha o campo DDD somente com números!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtDDDCliente.value.length < 2) {
		alert("DDD inválido!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtTelefoneCliente.value == "") {
		alert("Preencha o campo Telefone!");
		return;
	}
	if (isNaN(document.frmCadastroClienteAdm.txtTelefoneCliente.value)) {
		alert("Preencha o campo Telefone somente com números!");
		return;
	}
	if(document.frmCadastroClienteAdm.txtTelefoneCliente.value.length < 8) {
		alert("Telefone inválido!");
		return;
	}
	if (document.frmCadastroClienteAdm.cbCategorias.value == -1) {
		alert("Selecione uma Categoria!");
		return;
	}
	if (document.frmCadastroClienteAdm.cbTipoColeta.value == -1) {
		alert("Selecione um Tipo de Coleta!");
		return;
	}
	if (document.frmCadastroClienteAdm.cbGrupos.value == -1) {
		alert("Selecione um Grupo!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtCEPEnderecoComumCliente.value == "") {
		alert("Preencha o campo CEP do Endereço do Cliente!");
		return;
	}
	if (isNaN(document.frmCadastroClienteAdm.txtCEPEnderecoComumCliente.value)) {
		alert("Preencha o campo CEP somente com números!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtCEPEnderecoComumCliente.value.length < 8) {
		alert("CEP do Endereço do Cliente é inválido!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtLogradouro.value == "") {
		alert("Preencha o campo Logradouro do Endereço do Cliente!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtNumero.value == "") {
		alert("Preencha o campo Numero do Endereço do Cliente!");
		return;
	}
	if (isNaN(document.frmCadastroClienteAdm.txtNumero.value)) {
		alert("Preencha o campo Numero do Endereço do Cliente somente com números!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtBairro.value == "") {
		alert("Preencha o campo Bairro do Endereço do Cliente");
		return;
	}
	if (document.frmCadastroClienteAdm.txtMunicipio.value == "") {
		alert("Preencha o campo Município do Endereço do Cliente");
		return;
	}
	if (document.frmCadastroClienteAdm.txtEstado.value == "") {
		alert("Preencha o campo Estado do Endereço do Cliente!");
		return;
	}
	if (document.frmCadastroClienteAdm.hiddenActionIsColetaDomiciliar.value == 1) {
		if (document.frmCadastroClienteAdm.cbTransp.value == "-1") {
			alert("Selecione uma Transportadora!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtCEPEnderecoComumClienteColeta.value == "") {
			alert("Preencha o campo CEP do Endereço de Coleta!");
			return;
		}
		if (isNaN(document.frmCadastroClienteAdm.txtCEPEnderecoComumClienteColeta.value)) {
			alert("Preencha o campo CEP do Endereço de Coleta somente com números!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtLogradouroColeta.value == "") {
			alert("Preencha o campo Logradouro do Endereço de Coleta!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtNumeroColeta.value == "") {
			alert("Preencha o campo Número do Endereço de Coleta!");
			return;
		}
		if (isNaN(document.frmCadastroClienteAdm.txtNumeroColeta.value)) {
			alert("Preencha o campo Número do Endereço de Coleta somente com números!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtBairroColeta.value == "") {
			alert("Preencha o campo Bairro do Endereço de Coleta!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtMunicipioColeta.value == "") {
			alert("Preencha o campo Município do Endereço de Coleta!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtEstadoColeta.value == "") {
			alert("Preencha o campo Estado do Endereço de Coleta!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtRespColeta.value == "") {
			alert("Preencha o campo Resp. Coleta!");		
			return;
		}
		if (document.frmCadastroClienteAdm.txtDDDRespColeta.value == "") {
			alert("Preencha o campo DDD Resp. Coleta!");
			return;
		}
		if (isNaN(document.frmCadastroClienteAdm.txtDDDRespColeta.value)) {
			alert("Preencha o campo DDD Resp. Coleta somente com números!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtDDDRespColeta.value.length < 2) {
			alert("DDD Resp. Coleta inválido!");
			return;
		}
		if (document.frmCadastroClienteAdm.txtTelefoneRespColeta.value == "") {
			alert("Preencha o campo Telefone Resp. Coleta!");
			return;
		}
		if (isNaN(document.frmCadastroClienteAdm.txtTelefoneRespColeta.value)) {
			alert("Preencha o campo Telefone Resp. Coleta somente com números!");	
			return;
		}
		if (document.frmCadastroClienteAdm.txtTelefoneRespColeta.value.length < 8) {
			alert("Telefone Resp. Coleta inválido!");
			return;
		}
	}
	document.frmCadastroClienteAdm.hiddenActionForm.value = "SALVAR";
	document.frmCadastroClienteAdm.submit();
}

function validaCnpj() {
	var numeros1Dig = new Array(5,4,3,2,9,8,7,6,5,4,3,2);
	var soma1Dig = 0;
	var resto1Dig = 0;
	var digVer1 = 0;
	var numeros2Dig = new Array(6,5,4,3,2,9,8,7,6,5,4,3,2);
	var soma2Dig = 0;
	var resto2Dig = 0;
	var digVer2 = 0;
	var i = 0;
	var j = 0;
	var cnpj = "";

	cnpj = document.frmCadastroClienteAdm.txtCNPJCliente.value;

	digVer2 = cnpj.charAt(cnpj.length - 1);
	digVer1 = cnpj.charAt(cnpj.length - 2);
	
	if (document.frmCadastroClienteAdm.txtCNPJCliente.value.indexOf('/') == -1) {
		alert("Preencha corretamente o campo CNPJ");
		return;
	}
	if (document.frmCadastroClienteAdm.txtCNPJCliente.value.length < 18) {
		alert("Preencha corretamente o campo CNPJ");
		return;
	}
	cnpj = cnpj.replace('/','');
	cnpj = cnpj.replace('-','');
	cnpj = cnpj.replace('.','');
	cnpj = cnpj.replace('.','');
	for(i = 0; i < cnpj.length - 2; i++) {
		if (!isNaN(cnpj.charAt(i)) && !isNaN(numeros1Dig[i])) {
			soma1Dig += cnpj.charAt(i) * numeros1Dig[i];
		}
	}
	resto1Dig = soma1Dig % 11;
	if (resto1Dig < 2) {
		if (!(digVer1 == 0)) {
			alert("Preencha corretamente o campo CNPJ");
			return;
		}	
	} else {
		resto1Dig = 11 - resto1Dig;
		if (!(resto1Dig == digVer1)) {
			alert("Preencha corretamente o campo CNPJ");
			return;
		}	
	}
	for(j = 0; j < cnpj.length - 1; j++) {
		soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
	}
	resto2Dig = soma2Dig % 11;
	if (resto2Dig < 2) {
		if (!(digVer2 == 0)) {
			alert("Preencha corretamente o campo CNPJ");
			return;
		}	
	} else {
		resto2Dig = 11 - resto2Dig;
		if (!(resto2Dig == digVer2)) {
			alert("Preencha corretamente o campo CNPJ");
			return;
		}	
	}
}

function Ajax() {
	var ajax = null;
	if (window.ActiveXObject) {
		try {
			ajax = new ActiveXObject("Msxml2.XMLHTTP");	
		} catch (ex) {
			try {
				ajax = new ActiveXObject("Microsoft.XMLHTTP");
			} catch(ex2) {
				alert("Seu browser não suporta Ajax.");
			}			
		}
	} else {
		if (window.XMLHttpRequest) {
			try {
				ajax = new XMLHttpRequest();	
			} catch(ex3) {
				alert("Seu browser não suporta Ajax.");
			}	
		}	
	}
	
	return ajax;
}

function getEndereco() {
	var oAjax = Ajax();
	var form = document.frmCadastroClienteAdm;
	var strRet = "";

	form.txtLogradouro.value = "Carregando...";
	form.txtBairro.value = "Carregando...";
	form.txtMunicipio.value = "Carregando...";
	form.txtEstado.value = "Carregando...";
	form.txtCompLogradouro.value = "";
	form.txtNumero.value = "";
	document.body.style.cursor = 'wait';
	
	oAjax.onreadystatechange = function() {
		if (oAjax.readyState == 4 && oAjax.status == 200) {
			strRet = oAjax.responseText.split(";");
			strRet[6] = strRet[6].replace("        ",'');
			form.txtLogradouro.value = strRet[6] + ". " + strRet[2];
			form.txtBairro.value = strRet[3];
			form.txtMunicipio.value = strRet[4];
			form.txtEstado.value = strRet[5];
			document.body.style.cursor = 'default';
		}
	}
	
	oAjax.open("GET", "../../ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCEPEnderecoComumCliente.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
}

function getEnderecoColeta() {
	var oAjax = Ajax();
	var form = document.frmCadastroClienteAdm;
	var strRet = "";

	form.txtLogradouroColeta.value = "Carregando...";
	form.txtBairroColeta.value = "Carregando...";
	form.txtMunicipioColeta.value = "Carregando...";
	form.txtEstadoColeta.value = "Carregando...";
	form.txtCompLogradouroColeta.value = "";
	form.txtNumeroColeta.value = "";
	document.body.style.cursor = 'wait';

	oAjax.onreadystatechange = function() {
		if (oAjax.readyState == 4 && oAjax.status == 200) {
			strRet = oAjax.responseText.split(";");
			strRet[6] = strRet[6].replace("        ",'');
			form.txtLogradouroColeta.value = strRet[6] + ". " + strRet[2];
			form.txtBairroColeta.value = strRet[3];
			form.txtMunicipioColeta.value = strRet[4];
			form.txtEstadoColeta.value = strRet[5];
			document.body.style.cursor = 'default';
		}
	}
	
	oAjax.open("GET", "../../ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCEPEnderecoComumClienteColeta.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
}

function getPontoColeta() {
	var oAjax = Ajax();
	var form = document.frmCadastroClienteAdm;
	var strRet = "";

	oAjax.onreadystatechange = function() {
		if (oAjax.readyState == 4 && oAjax.status == 200) {
			if (oAjax.responseText != "") {
				strRet = oAjax.responseText.split(";");
				form.txtIDPontoDeColeta.value = strRet[0];
				form.txtPontoDeColeta.value = strRet[1];
				form.txtCNPJPontoDeColeta.value = strRet[2];
			} else {
				alert("Não existe Ponto de Coleta com esse ID!");
				form.txtIDPontoDeColeta.value = "";
				form.txtPontoDeColeta.value = "";
				form.txtCNPJPontoDeColeta.value = "";
			}
		}	
	}

	oAjax.open("GET", "ajax/ajax.asp?sub=getpontocoleta&value="+form.txtIDPontoDeColeta.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
}

function tipopessoa() {
	var form = document.frmCadastroClienteAdm;
	if (form.radioPessoa[0].checked) {
		document.getElementById("razaosocial").style.display = 'none';
		document.getElementById("nomefantasia").style.display = 'none';
		document.getElementById("cnpj").style.display = 'none';
		document.getElementById("inscestadual").style.display = 'none';
		document.getElementById("nomecliente").style.display = 'block';	
		document.getElementById("cpf").style.display = 'block';	
	} else {
		if (form.radioPessoa[1].checked) {
			document.getElementById("nomecliente").style.display = 'none';	
			document.getElementById("cpf").style.display = 'none';	
			document.getElementById("razaosocial").style.display = 'block';
			document.getElementById("nomefantasia").style.display = 'block';
			document.getElementById("cnpj").style.display = 'block';
			document.getElementById("inscestadual").style.display = 'block';
		}	
	}
}

function validaCPF() {
	var numeros1Dig = new Array(10,9,8,7,6,5,4,3,2);
	var soma1Dig = 0;
	var resto1Dig = 0;
	var digVer1 = 0;
	var numeros2Dig = new Array(11,10,9,8,7,6,5,4,3,2);
	var soma2Dig = 0;
	var resto2Dig = 0;
	var digVer2 = 0;
	var i = 0;
	var j = 0;
	var cnpj = "";

	cnpj = document.frmCadastroClienteAdm.txtCPFCliente.value;

	digVer2 = cnpj.charAt(cnpj.length - 1);
	digVer1 = cnpj.charAt(cnpj.length - 2);
	
	if (document.frmCadastroClienteAdm.txtCPFCliente.value == "") {
		alert("Preencha o campo CPF!");
		return;
	}
	if (document.frmCadastroClienteAdm.txtCPFCliente.value.length < 14) {
		alert("Preencha o campo CPF corretamente!");
		return;
	}
	cnpj = cnpj.replace('-','');
	cnpj = cnpj.replace('.','');
	cnpj = cnpj.replace('.','');
	for(i = 0; i < cnpj.length - 2; i++) {
		if (!isNaN(cnpj.charAt(i)) && !isNaN(numeros1Dig[i])) {
			soma1Dig += cnpj.charAt(i) * numeros1Dig[i];
		}
	}
	resto1Dig = soma1Dig % 11;
	if (resto1Dig < 2) {
		if (!(digVer1 == 0)) {
			alert("Preencha o campo CPF corretamente!");
			return;
		}	
	} else {
		resto1Dig = 11 - resto1Dig;
		if (!(resto1Dig == digVer1)) {
			alert("Preencha o campo CPF corretamente!");
			return;
		}	
	}
	for(j = 0; j < cnpj.length - 1; j++) {
		soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
	}
	resto2Dig = soma2Dig % 11;
	if (resto2Dig < 2) {
		if (!(digVer2 == 0)) {
			alert("Preencha o campo CPF corretamente!");
			return;
		}	
	} else {
		resto2Dig = 11 - resto2Dig;
		if (!(resto2Dig == digVer2)) {
			alert("Preencha o campo CPF corretamente!");
			return;
		}	
	}
}

function changeListenerTipoColeta() {
		var form = document.frmCadastroClienteAdm;
		if (form.cbTipoColeta.value == 0) {
			form.txtQtdCartuchos.value = 1;	
		} else {
			if (form.hiddenIntQtdCartuchos.value <= 1) {
				form.txtQtdCartuchos.value = 0;
			} else {
				form.txtQtdCartuchos.value = form.hiddenIntQtdCartuchos.value;
			}
		}
}



