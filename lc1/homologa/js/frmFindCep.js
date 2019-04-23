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
	
	function showClienteEndColeta() {
		var form = document.frmEditaCadastroCliente;

		if (form.txtColetaCep.value == "") {
			msgErrForm += "CEP Coleta inválido\n";
			errForm = true;
		}
		if (isNaN(form.txtColetaCep.value)) {
			msgErrForm += "CEP de Coleta deve ter somente números\n";
			errForm = true;
		}
		if (form.txtColetaCep.value.length < 8) {
			msgErrForm += "CEP de Coleta deve ter 8 digitos\n";
			errForm = true;
		}
		if (!errForm) {
			preencheEndColetaToCep();
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
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheEndColetaToCep() {
		var oAjax = Ajax();
		var form = document.frmEditaCadastroCliente;
		var strRet = "";

		document.getElementById("loadingdisplay").style.display = 'block';
		
		form.txtLogradouroColeta.value = "Loading...";
		form.txtBairroColeta.value = "Loading...";
		form.txtMunicipioColeta.value = "Loading...";
		form.txtEstadoColeta.value = "Loading...";
		form.txtCompLogradouroColeta.value = "";
		form.txtNumeroColeta.value = "";

		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				document.getElementById("loadingdisplay").style.display = 'none';
				strRet = oAjax.responseText.split(";");
				strRet[6] = strRet[6].replace("        ",'');
				form.txtLogradouroColeta.value = strRet[6] + ". " + strRet[2];
				form.txtBairroColeta.value = strRet[3];
				form.txtMunicipioColeta.value = strRet[4];
				form.txtEstadoColeta.value = strRet[5];
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+form.txtColetaCep.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}	
	
	function usaMesmoEndereco() {
	
		var form = document.frmEditaCadastroCliente;
		var bReadOnly = true;
		
		if (document.frmEditaCadastroCliente.chkMesmoEndereco.checked == 1){
			bReadOnly = true;
		}
		else {
			bReadOnly = false;
		}
		
		document.frmEditaCadastroCliente.txtColetaCep.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtLogradouroColeta.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtCompLogradouroColeta.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtNumeroColeta.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtBairroColeta.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtMunicipioColeta.readOnly = bReadOnly;
		document.frmEditaCadastroCliente.txtEstadoColeta.readOnly = bReadOnly;
		
		if (bReadOnly == true) {
			document.frmEditaCadastroCliente.txtColetaCep.value = document.frmEditaCadastroCliente.txtCep.value;
			document.frmEditaCadastroCliente.txtLogradouroColeta.value = document.frmEditaCadastroCliente.txtLogradouro.value;
			document.frmEditaCadastroCliente.txtCompLogradouroColeta.value = document.frmEditaCadastroCliente.txtCompLogradouro.value;
			document.frmEditaCadastroCliente.txtNumeroColeta.value = document.frmEditaCadastroCliente.txtNumero.value;
			document.frmEditaCadastroCliente.txtBairroColeta.value = document.frmEditaCadastroCliente.txtBairro.value;
			document.frmEditaCadastroCliente.txtMunicipioColeta.value = document.frmEditaCadastroCliente.txtMunicipio.value;
			document.frmEditaCadastroCliente.txtEstadoColeta.value = document.frmEditaCadastroCliente.txtEstado.value;
		}
		else {
			document.frmEditaCadastroCliente.txtColetaCep.value = "";
			document.frmEditaCadastroCliente.txtLogradouroColeta.value = "";
			document.frmEditaCadastroCliente.txtCompLogradouroColeta.value = "";
			document.frmEditaCadastroCliente.txtNumeroColeta.value = "";
			document.frmEditaCadastroCliente.txtBairroColeta.value = "";
			document.frmEditaCadastroCliente.txtMunicipioColeta.value = "";
			document.frmEditaCadastroCliente.txtEstadoColeta.value = "";
		}
		
		
	}