// JavaScript Document

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

function validate() {
	var form = document.frmPontoColetaAdm;
	if (form.txtRazaoSocial.value == "") {
		alert("Preencha o campo Razão Social!");
		return;
	}
	if (form.txtNomeFantasia.value == "") {
		alert("Preencha o campo Nome Fantasia!");
		return;
	}
	if (!cnpj()) {
		return;
	}
	if (form.txtUsuario.value == "") {
		alert("Preencha o campo Usuário!");
		return;
	}
	if (form.txtSenha.value == "") {
		alert("Preencha o campo Senha!");
		return;
	}
	if (form.txtCEP.value == "") {
		alert("Preencha o campo CEP!");
		return;
	}
	if (form.txtLogradouro.value == "") {
		alert("Preencha o campo Logradouro!");
		return;
	}
	if (form.txtNumero.value == "") {
		alert("Preencha o campo Número!");
		return;
	}
	if (form.txtBairro.value == "") {
		alert("Preencha o campo Bairro!");
		return;
	}
	if (form.txtMunicipio.value == "") {
		alert("Preencha o campo Município!");
		return;
	}
	if (form.txtEstado.value == "") {
		alert("Preencha o campo Estado!");
		return;
	}
	if (form.txtDDD.value == "") {
		alert("Preencha o campo DDD!");
		return;
	}
	if (form.txtTelefone.value == "") {
		alert("Preencha o campo Telefone!");
		return;
	}
	if (form.txtQtdCartuchos.value == "") {
		form.txtQtdCartuchos.value = 0;
	}
	form.submit();
}

function endereco() {
	var oAjax = Ajax();
	var form = document.frmPontoColetaAdm;
	var strRet = "";

	form.txtLogradouro.value = "Carregando...";
	form.txtBairro.value = "Carregando...";
	form.txtMunicipio.value = "Carregando...";
	form.txtEstado.value = "Carregando...";
	
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
	
	oAjax.open("GET", "../../ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCEP.value, true);
	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
	oAjax.send(null);
}

function cnpj() {
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

	cnpj = document.frmPontoColetaAdm.txtCNPJ.value;

	digVer2 = cnpj.charAt(cnpj.length - 1);
	digVer1 = cnpj.charAt(cnpj.length - 2);
	
	if (document.frmPontoColetaAdm.txtCNPJ.value.indexOf('/') == -1) {
		alert("Preencha corretamente o campo CNPJ");
		return false;
	}
	if (document.frmPontoColetaAdm.txtCNPJ.value.length < 18) {
		alert("Preencha corretamente o campo CNPJ");
		return false;
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
			return false;
		}	
	} else {
		resto1Dig = 11 - resto1Dig;
		if (!(resto1Dig == digVer1)) {
			alert("Preencha corretamente o campo CNPJ");
			return false;
		}	
	}
	for(j = 0; j < cnpj.length - 1; j++) {
		soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
	}
	resto2Dig = soma2Dig % 11;
	if (resto2Dig < 2) {
		if (!(digVer2 == 0)) {
			alert("Preencha corretamente o campo CNPJ");
			return false;
		}	
	} else {
		resto2Dig = 11 - resto2Dig;
		if (!(resto2Dig == digVer2)) {
			alert("Preencha corretamente o campo CNPJ");
			return false;
		}	
	}
	return true;
}

function cnpj_format(cnpj) {
	var form = document.frmPontoColetaAdm;
	if (cnpj.value.length == 2 || cnpj.value.length == 6) {
		form.txtCNPJ.value += ".";	
	}
	if (cnpj.value.length == 10) {
		form.txtCNPJ.value += "/";		
	}
	if (cnpj.value.length == 15) {
		form.txtCNPJ.value += "-";		
	}
}

function cnpjExists() {
	var oAjax2 = Ajax();
	if (document.frmPontoColetaAdm.txtCNPJ.value != "" && (document.frmPontoColetaAdm.verifycnpj.value != document.frmPontoColetaAdm.txtCNPJ.value)) {
		oAjax2.onreadystatechange = function() {
			if (oAjax2.readyState == 4 && oAjax2.status == 200) {
				if (oAjax2.responseText == "true") {
					alert("CNPJ já cadastrado, favor insira outro CNPJ!");
					document.frmPontoColetaAdm.txtCNPJ.focus();
				} 		
			}	
		}
		
		oAjax2.open("GET", "ajax/Ajax.asp?sub=cnpjexists&id="+document.frmPontoColetaAdm.txtCNPJ.value, true);
		oAjax2.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax2.send(null);
	}
}