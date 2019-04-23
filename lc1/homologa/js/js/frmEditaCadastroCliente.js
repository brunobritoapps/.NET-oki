// JavaScript Document
/*
'|--------------------------------------------------------------------
'| Arquivo: frmEditaCadastroCliente.js																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação: 15/04/2007																		 
'| Descrição: Arquivo de Formulário para UPDATE do Cliente (Javascript)
'|--------------------------------------------------------------------
*/
	var errForm = false;
	var	msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";
	
	function showClienteEndereco() {
		var form = document.frmEditaCadastroCliente;

		if (form.txtNomeFantasia.value == "") {
			msgErrForm += "Campo: Nome Fantasia\n";
			errForm = true;
		}
		if (form.txtCep.value == "") {
			msgErrForm += "Campo: Cep\n";
			errForm = true;
		}
		if (isNaN(form.txtCep.value)) {
			msgErrForm += "Campo: Somente números no Cep\n";
			errForm = true;
		}
		if (form.txtCep.value.length < 8) {
			msgErrForm += "Campo: Cep\n";
			errForm = true;
		}
		if (parseInt(form.tipopessoa.value) == 1) {
			validaCnpj();	
		} else {
			validaDDD();	
		}
		validaTelefone();
		if (!errForm) {
			preencheEnderecoToCep();
		} else {
			alert(msgErrForm);	
			msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";
			errForm = false;
			return false;
		}
		errForm = false;
	}
	
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Validação do CNPJ da Empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaCnpj() {
		var form = document.frmEditaCadastroCliente;
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
	
		cnpj = form.txtCNPJ.value;

		digVer2 = cnpj.charAt(cnpj.length - 1);
		digVer1 = cnpj.charAt(cnpj.length - 2);
		
		if (form.txtCNPJ.value == "") {
			msgErrForm += "Campo: CNPJ\n";
			errForm = true;
			return false;
		}
		if (form.txtCNPJ.value.indexOf('/') == -1) {
			msgErrForm += "Campo: CNPJ\n";
			errForm = true;
			return false;
		}
		if (form.txtCNPJ.value.length < 18) {
			msgErrForm += "Campo: CNPJ\n";
			errForm = true;
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
				msgErrForm += "Campo: CNPJ\n";
				errForm = true;
				return false;
			}	
		} else {
			resto1Dig = 11 - resto1Dig;
			if (!(resto1Dig == digVer1)) {
				msgErrForm += "Campo: CNPJ\n";
				errForm = true;
				return false;
			}	
		}
		for(j = 0; j < cnpj.length - 1; j++) {
			soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
		}
		resto2Dig = soma2Dig % 11;
		if (resto2Dig < 2) {
			if (!(digVer2 == 0)) {
				msgErrForm += "Campo: CNPJ\n";
				errForm = true;
				return false;
			}	
		} else {
			resto2Dig = 11 - resto2Dig;
			if (!(resto2Dig == digVer2)) {
				msgErrForm += "Campo: CNPJ\n";
				errForm = true;
				return false;
			}	
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Validação DDD e auto preenchimento
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaDDD() {
		var form = document.frmEditaCadastroCliente;

		if (form.txtDDD.value == "") {
			msgErrForm += "Campo: DDD\n";
			errForm = true;
			return false;
		}		
		if (isNaN(form.txtDDD.value)) {
			msgErrForm += "Campo: DDD\n";
			errForm = true;
			return false;
		}
		if (form.txtDDD.value.length == 1) {
			msgErrForm += "Campo: DDD\n";
			errForm = true;
			return false;
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Validação do Telefone da Empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaTelefone() {
		var form = document.frmEditaCadastroCliente;

		if (form.txtTelefone.value == "") {
			msgErrForm += "Campo: Telefone\n";
			errForm = true;
			return false;
		}		
		if (isNaN(form.txtTelefone.value)) {
			msgErrForm += "Campo: Telefone\n";
			errForm = true;
			return false;
		}
		if (form.txtTelefone.value.length < 8) {
			msgErrForm += "Campo: Telefone\n";
			errForm = true;
			return false;
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheEnderecoToCep() {
		var oAjax = Ajax();
		var form = document.frmEditaCadastroCliente;
		var strRet = "";

		document.getElementById("loadingdisplay").style.display = 'block';
		form.txtLogradouro.value = "Loading...";
		form.txtBairro.value = "Loading...";
		form.txtMunicipio.value = "Loading...";
		form.txtEstado.value = "Loading...";
		form.txtCompLogradouro.value = "";
		form.txtNumero.value = "";

		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				document.getElementById("loadingdisplay").style.display = 'none';
				strRet = oAjax.responseText.split(";");
				strRet[6] = strRet[6].replace("        ",'');
				form.txtLogradouro.value = strRet[6] + ". " + strRet[2];
				form.txtBairro.value = strRet[3];
				form.txtMunicipio.value = strRet[4];
				form.txtEstado.value = strRet[5];
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtCep.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Criação do Objeto Ajax
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Verifica se o o CNPJ se já está cadastrado
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkCNPJEmpresa() {
		var oAjax = Ajax();
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				if (oAjax.responseText == "true") {
					alert("CNPJ já cadastrado. Favor cadastrar outro CNPJ!");
					document.frmEditaCadastroCliente.txtCNPJ.focus();
					return false;
				} else {
					return true;
				}
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcheckcnpjempresa&id="+document.frmEditaCadastroCliente.txtCNPJ.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca De Pontos de Coleta do cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showClientePostoColeta() {
		var form = document.frmEditaCadastroCliente;
		var oAjax = Ajax();
		var strRet = "";
		if (form.txtCepConsultaPonto.value.length < 8) {
			alert("Preencha corretamente o Cep para Busca!");
			form.txtCepConsultaPonto.focus();
			return false;
		}
		if (isNaN(form.txtCepConsultaPonto.value)) {
			alert("Preencha somente números no Cep para Busca!");
			form.txtCepConsultaPonto.focus();
			return false;
		}
		if (form.txtCepConsultaPonto.value == "") {
			alert("Preencha o campo de Cep para busca dos Pontos de Coleta!");
			form.txtCepConsultaPonto.focus();
			return false;
		} else {
			oAjax.onreadystatechange = function() {
				if (oAjax.readyState == 4 && oAjax.status == 200) {
					strRet = oAjax.responseText.split(";");
					document.getElementById("titTableListPontoColeta").style.display = 'block';
					document.getElementById("tableListPontoColeta").innerHTML = strRet[0];
					document.frmEditaCadastroCliente.hiddenIntPontoColeta.value = strRet[1];
				} 	
			}
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getlistpontocoleta&id="+document.frmEditaCadastroCliente.txtCepConsultaPonto.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		}
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function loadCepColeta() {
		var oAjax = Ajax();
		var strRet = "";

		if (document.frmEditaCadastroCliente.txtCepColeta.value == "") {
			alert("Preencha o campo Cep para Busca do Endereço de Coleta!");
			return false;
		}
		if (isNaN(document.frmEditaCadastroCliente.txtCepColeta.value)) {
			alert("Preencha o Cep somente com números!");
			return false;
		}
		if (document.frmEditaCadastroCliente.txtCepColeta.value.length < 8) {
			alert("Preencha corretamente o Campo Cep!");
			return false;
		}
		//====================================================================
		// Bloqueio dos campos de consulta
		//====================================================================
		document.frmEditaCadastroCliente.chkMesmoEndereco.checked = false;
		document.frmEditaCadastroCliente.txtLogradouroColeta.disabled = true;
		document.frmEditaCadastroCliente.txtLogradouroColeta.value = "Loading...";
		document.frmEditaCadastroCliente.txtCompLogradouroColeta.disabled = true;
		document.frmEditaCadastroCliente.txtNumeroColeta.disabled = true;
		document.frmEditaCadastroCliente.txtBairroColeta.disabled = true;
		document.frmEditaCadastroCliente.txtBairroColeta.value = "Loading...";
		document.frmEditaCadastroCliente.txtMunicipioColeta.disabled = true;
		document.frmEditaCadastroCliente.txtMunicipioColeta.value = "Loading...";
		document.frmEditaCadastroCliente.txtEstadoColeta.disabled = true;
		document.frmEditaCadastroCliente.txtEstadoColeta.value = "Loading...";
		document.getElementById("btnBuscarCepColeta").style.cursor = 'wait';
		document.body.style.cursor = 'wait';
		//====================================================================
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				strRet = oAjax.responseText.split(";");
				document.frmEditaCadastroCliente.hiddenIntEnderecoCepColeta.value = strRet[0];
				document.frmEditaCadastroCliente.txtLogradouroColeta.value = strRet[2];
				document.frmEditaCadastroCliente.txtBairroColeta.value = strRet[3];
				document.frmEditaCadastroCliente.txtMunicipioColeta.value = strRet[4];
				document.frmEditaCadastroCliente.txtEstadoColeta.value = strRet[5];
				document.frmEditaCadastroCliente.txtLogradouroColeta.disabled = false;
				document.frmEditaCadastroCliente.txtCompLogradouroColeta.disabled = false;
				document.frmEditaCadastroCliente.txtNumeroColeta.disabled = false;
				document.frmEditaCadastroCliente.txtBairroColeta.disabled = false;
				document.frmEditaCadastroCliente.txtMunicipioColeta.disabled = false;
				document.frmEditaCadastroCliente.txtEstadoColeta.disabled = false;
				document.getElementById("btnBuscarCepColeta").style.cursor = 'pointer';
				document.body.style.cursor = 'default';
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+document.frmEditaCadastroCliente.txtCepColeta.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Preenchimento Automático do mesmo Endereço
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheMesmoEndereco() {
		var form = document.frmEditaCadastroCliente;
		if (form.chkMesmoEndereco.checked) {
			form.txtCepColeta.value = form.txtCep.value;
			form.txtLogradouroColeta.value = form.txtLogradouro.value;
			form.txtCompLogradouroColeta.value = form.txtCompLogradouro.value;
			form.txtNumeroColeta.value = form.txtNumero.value;
			form.txtBairroColeta.value = form.txtBairro.value;
			form.txtMunicipioColeta.value = form.txtMunicipio.value;
			form.txtEstadoColeta.value = form.txtEstado.value;
		} else {
			form.txtCepColeta.value = "";
			form.txtLogradouroColeta.value = "";
			form.txtCompLogradouroColeta.value = "";
			form.txtNumeroColeta.value = "";
			form.txtBairroColeta.value = "";
			form.txtMunicipioColeta.value = "";
			form.txtEstadoColeta.value = "";
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Check se foi selecionado algum ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkChangePontoColeta() {
		var bErr = false;
		var bSelected = false;
		for (var i=0;i <= parseInt(document.frmEditaCadastroCliente.hiddenIntPontoColeta.value);i++) {
			if (!document.getElementById("radioCheckPonto"+i).checked) {
				bErr = true;
			} else {
				bSelected = true;
				document.frmEditaCadastroCliente.hiddenIntChangePontoColeta.value = document.getElementById("radioCheckPonto"+i).value;
			}
		}
		
		if (bErr && !bSelected) {
			return false
		} else {
			return true;	
		}
 	}
	//=========================================================================================================
	
	function validate() {
		var form = document.frmEditaCadastroCliente;
		if (form.txtNomeFantasia.value == "") {
			msgErrForm += "Campo: Nome Fantasia\n";
			errForm = true;
		}
		if (form.txtCep.value == "") {
			msgErrForm += "Campo: Cep\n";
			errForm = true;
		}
		if (isNaN(form.txtCep.value)) {
			msgErrForm += "Campo: Somente números no Cep\n";
			errForm = true;
		}
		if (form.txtCep.value.length < 8) {
			msgErrForm += "Campo: Cep\n";
			errForm = true;
		}
		if (parseInt(form.tipopessoa.value) == 1) {
			validaCnpj();
		} else {
			validateCPF();	
		}
		validaDDD();
		validaTelefone();
		if (form.txtNumero.value == "") {
			msgErrForm += "Campo: Numero\n";	
			errForm = true;
		}

		if (!errForm) {
			form.submit();
		} else {
			alert(msgErrForm);	
			msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";
			errForm = false;
			return false;
		}
		errForm = false;
		
	}
	
	
	function validateCPF() {
		var form = document.frmEditaCadastroCliente;
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
	
		cnpj = form.txtCPF.value;

		digVer2 = cnpj.charAt(cnpj.length - 1);
		digVer1 = cnpj.charAt(cnpj.length - 2);
		
		if (form.txtCPF.value == "") {
			msgErrForm += "Campo: CPF\n";
			errForm = true;
			return false;
		}
		if (form.txtCPF.value.length < 14) {
			msgErrForm += "Campo: CPF\n";
			errForm = true;
			return false;
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
				msgErrForm += "Campo: CPF\n";
				errForm = true;
				return false;
			}	
		} else {
			resto1Dig = 11 - resto1Dig;
			if (!(resto1Dig == digVer1)) {
				msgErrForm += "Campo: CPF\n";
				errForm = true;
				return false;
			}	
		}
		for(j = 0; j < cnpj.length - 1; j++) {
			soma2Dig += cnpj.charAt(j) * numeros2Dig[j];
		}
		resto2Dig = soma2Dig % 11;
		if (resto2Dig < 2) {
			if (!(digVer2 == 0)) {
				msgErrForm += "Campo: CPF\n";
				errForm = true;
				return false;
			}	
		} else {
			resto2Dig = 11 - resto2Dig;
			if (!(resto2Dig == digVer2)) {
				msgErrForm += "Campo: CPF\n";
				errForm = true;
				return false;
			}	
		}
	}

