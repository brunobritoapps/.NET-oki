function tipopessoa() {

    return;
}
function windowLocationFind(url) {
	var find = true;
	var busca = document.getElementsByName("txtFindCliente")[0].value;
	var por = document.getElementsByName("typeFindCliente")[0].value;
	var status = document.getElementsByName("cbStatusFindCliente")[0].value;
	var categoria = document.getElementsByName("cbCategoriasFindCliente")[0].value;
	//var grupos = document.frmEditCadastroClienteLc.grupos.value;
	//window.location.href = url + "frmEditCadastroClienteLc.asp?find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria;
	window.location.href = url + "../lc/homologa/adm/frmCadastroClienteAdm.asp?find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria;
}


function aprovar() {
	if (document.frmEditCadastroClienteLc.hiddenActionIsColetaDomiciliar.value == 1) 
	{
		//if (document.frmEditCadastroClienteLc.cbTransp.value == "-1") 
		//{
		//	alert("Selecione uma Transportadora!");
		//	return;
		//}
	}
	document.frmEditCadastroClienteLc.hiddenActionManagerProve.value = 'true';
	document.frmEditCadastroClienteLc.hiddenActionForm.value = "APROVAR";
	document.frmEditCadastroClienteLc.submit();
}

function reprovar() {
	document.frmEditCadastroClienteLc.hiddenActionManagerProve.value = 'false';		
	document.frmEditCadastroClienteLc.hiddenActionForm.value = "REPROVAR";
	document.frmEditCadastroClienteLc.submit();
}

function validar() {
	document.getElementsByName("hiddenActionForm")[0].value = "SALVAR";
	document.frmEditCadastroClienteLc.submit();
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

	cnpj = document.frmEditCadastroClienteLc.txtCNPJCliente.value;

	digVer2 = cnpj.charAt(cnpj.length - 1);
	digVer1 = cnpj.charAt(cnpj.length - 2);
	
	if (document.frmEditCadastroClienteLc.txtCNPJCliente.value.indexOf('/') == -1) {
		alert("Preencha corretamente o campo CNPJ erro:215");
		return;
	}
	if (document.frmEditCadastroClienteLc.txtCNPJCliente.value.length < 18) {
		alert("Preencha corretamente o campo CNPJ erro:219");
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
			alert("Preencha corretamente o campo CNPJ erro:234");
			return;
		}	
	} else {
		resto1Dig = 11 - resto1Dig;
		if (!(resto1Dig == digVer1)) {
			alert("Preencha corretamente o campo CNPJ erro:240");
			return;
		}	
	}
	for(j = 0; j < cnpj.length - 1; j++) {
		soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
	}
	resto2Dig = soma2Dig % 11;
	if (resto2Dig < 2) {
		if (!(digVer2 == 0)) {
			alert("Preencha corretamente o campo CNPJ erro:250");
			return;
		}	
	} else {
		resto2Dig = 11 - resto2Dig;
		if (!(resto2Dig == digVer2)) {
			alert("Preencha corretamente o campo CNPJ eroo:256");
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
	var form = document.frmEditCadastroClienteLc;
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
	        strRet[6] = strRet[6].replace("        ", '');
			form.txtLogradouro.value = strRet[6] + "" + strRet[2];
			form.txtBairro.value = strRet[3];
			form.txtMunicipio.value = strRet[4];
			form.txtEstado.value = strRet[5];
			form.txtLogradouro.value = strRet[2];
			document.body.style.cursor = 'default';
		}
	}
	
	oAjax.open("GET", "../ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCEPEnderecoComumCliente.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
	return;
}

function getEnderecoColeta() {
	var oAjax = Ajax();
	var form = document.frmEditCadastroClienteLc;
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
	oAjax.open("GET", "../ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCEPEnderecoComumClienteColeta.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
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

	cnpj = document.frmEditCadastroClienteLc.txtCPFCliente.value;

	digVer2 = cnpj.charAt(cnpj.length - 1);
	digVer1 = cnpj.charAt(cnpj.length - 2);
	
	if (document.frmEditCadastroClienteLc.txtCPFCliente.value == "") {
		alert("Preencha o campo CPF!");
		return;
	}
	if (document.frmEditCadastroClienteLc.txtCPFCliente.value.length < 14) {
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