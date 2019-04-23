// JavaScript Document
/*
'|--------------------------------------------------------------------
'| Arquivo: frmAddSolicitacao.js																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação: 15/04/2007																		 
'| Descrição: Arquivo de Formulário para Nova Solicitacao (Javascript)
'|--------------------------------------------------------------------
*/

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
	
	//libera digitação do endereço novo
	function checkEndereco(clicked_id) {
	
		var elThisElement = clicked_id;
		var elToggled = document.getElementById("tagendnovo");
		var elToggledmesmo = document.getElementById("tagendmesmo");
		var oPreenche = preencheMesmoEndereco();
		
		if (elThisElement == "radioendmesmo") {
			elToggled.style.display = "none";
			elToggledmesmo.style.display = "block";
		}
		
		if (elThisElement == "radioendnovo") {
			elToggled.style.display = "block";
			elToggledmesmo.style.display = "none";
		}
		
	}
	

	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Check se foi selecionado algum ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkChangePontoColeta() {
		try {
			var bErr = false;
			var bSelected = false;
			for (var i=0;i <= parseInt(document.frmAddSolicitacao.hiddenIntPontoColeta.value);i++) {
				if (!document.getElementById("radioCheckPonto"+i).checked) {
					bErr = true;
				} else {
					bSelected = true;
					document.frmAddSolicitacao.hiddenIntChangePontoColeta.value = document.getElementById("radioCheckPonto"+i).value;
				}
			}
			
			if (bErr && !bSelected) {
				return false
			} else {
				return true;	
			}
		} catch (ex) {
			return true;
		}
 	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Valida Formulário de Nova Solicitação
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaFormulario() {
		if (document.frmAddSolicitacao.txtQtdCartuchos.value == "") {
			alert("Preencha o campo Quantidade de Cartuchos");
			return false;
		}
		if (isNaN(document.frmAddSolicitacao.txtQtdCartuchos.value)) {
			alert("O Campo Quantidade de Cartuchos só aceita dados numéricos");
			return false;
		}
		//if (parseInt(document.frmAddSolicitacao.txtQtdCartuchos.value) < parseInt(document.frmAddSolicitacao.hiddenMinCartuchos.value)) {
		//	alert("O mínimo de cartuchos para esta categoria é de " + document.frmAddSolicitacao.hiddenMinCartuchos.value + " cartuchos. Por favor preencha com um valor maior de cartuchos a serem entregues!");
		//	return false;
		//}
		//if (!checkChangePontoColeta()) {
		//	if (!confirm("Deseja usar o mesmo Ponto de Coleta da Solicitação anterior?")) {
		//		alert("Selecione um Ponto de Coleta");
		//		return false;
		//	}
		//}
		if (document.frmAddSolicitacao.hiddenSessionisColetaDomiciliar.value == 1) {
		    if (isNaN(document.frmAddSolicitacao.txtCepColeta.value)) {
		        alert("Preencha o campo CEP apenas com números");
		        document.frmAddSolicitacao.hiddenActionForm.value = "3";
		        return false;
		    }
		    if (document.frmAddSolicitacao.txtCepColeta.value.length < 8) {
		        alert("Preencha o campo CEP com 8 números!");
		        document.frmAddSolicitacao.hiddenActionForm.value = "3";
		        return false;
		    }

			if (document.frmAddSolicitacao.txtCepColeta.value == "") {
				alert("Preencha o campo CEP de coleta");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}	
			if (document.frmAddSolicitacao.txtLogradouroColeta.value == "") {
				alert("Preencha o campo Endereço");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}
			if (document.frmAddSolicitacao.txtNumeroColeta.value == "") {
				alert("Preencha o campo Numero do Endereço");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}
			if (document.frmAddSolicitacao.txtBairroColeta.value == "") {
				alert("Preencha o campo Bairro");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}
			if (document.frmAddSolicitacao.txtMunicipioColeta.value == "") {
				alert("Preencha o campo Município");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}
			if (document.frmAddSolicitacao.txtEstadoColeta.value == "") {
				alert("Preencha o campo Estado");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				return false;
			}
			if (document.frmAddSolicitacao.txtRespColContato.value == "") {
				alert("Preencha o campo do Contato responsável pela Coleta!");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				form.txtRespColContato.focus();
				return false;
			}
			if (document.frmAddSolicitacao.txtDDDContatoRespColeta.value == "") {
					alert("Preencha o campo do DDD do responsável pela Coleta!");
					document.frmAddSolicitacao.hiddenActionForm.value = "3";
					form.txtDDDContatoRespColeta.focus();
					return false;
			}
			if (isNaN(document.frmAddSolicitacao.txtDDDContatoRespColeta.value)) {
				alert("Preencha o campo DDD do Contato somente com dados numéricos!");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				form.txtDDDContatoRespColeta.focus();
				return false;
			}
			if (document.frmAddSolicitacao.txtDDDContatoRespColeta.value.length < 2) {
				alert("Preencha o campo do DDD do Contato com no mínimo 2 caracteres válidos!");
				form.txtDDDContatoRespColeta.focus();
				return false;
			}
			if (document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value == "") {
					alert("Preencha o campo do Telefone do responsável pela Coleta!");
					document.frmAddSolicitacao.hiddenActionForm.value = "3";
					form.txtTelefoneContatoRespColeta.focus();
					return false;
			}
			if (isNaN(document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value)) {
				alert("Preencha o campo Telefone do Contato somente com dados numéricos!");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				form.txtTelefoneContatoRespColeta.focus();
				return false;
			}
			if (document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value.length < 8) {
				alert("Preencha o campo do Telefone do Contato com no mínimo 8 caracteres válidos!");
				document.frmAddSolicitacao.hiddenActionForm.value = "3";
				form.txtTelefoneContatoRespColeta.focus();
				return false;
			}
			if (document.frmAddSolicitacao.txtDepContatoRespColeta.value == "") {
			    alert("Preencha o campo Departamento do contato!");
			    document.frmAddSolicitacao.hiddenActionForm.value = "3";
			    form.txtDepContatoRespColeta.focus();
			    return false;

			}
		}
		document.frmAddSolicitacao.submit();
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca De Pontos de Coleta do cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showClientePostoColeta() {
		var form = document.frmAddSolicitacao;
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
					document.frmAddSolicitacao.hiddenIntPontoColeta.value = strRet[1];
				} 	
			}
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getlistpontocoleta&id="+form.txtCepConsultaPonto.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		}
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function loadCepColeta() {
		var oAjax = Ajax();
		var strRet = "";

		if (document.frmAddSolicitacao.txtCepColeta.value == "") {
			alert("Preencha o campo Cep para Busca do Endereço de Coleta!");
			return false;
		}
		if (isNaN(document.frmAddSolicitacao.txtCepColeta.value)) {
			alert("Preencha o Cep somente com números!");
			return false;
		}
		if (document.frmAddSolicitacao.txtCepColeta.value.length < 8) {
			alert("Preencha corretamente o Campo Cep!");
			return false;
		}
		//====================================================================
		// Bloqueio dos campos de consulta
		//====================================================================
		document.frmAddSolicitacao.chkMesmoEndereco.checked = false;
		document.frmAddSolicitacao.chkNovoEndereco.checked = true;
		document.frmAddSolicitacao.txtLogradouroColeta.disabled = true;
		document.frmAddSolicitacao.txtLogradouroColeta.value = "Loading...";
		document.frmAddSolicitacao.txtCompLogradouroColeta.value = "";
		document.frmAddSolicitacao.txtNumeroColeta.value = "";
		document.frmAddSolicitacao.txtBairroColeta.disabled = true;
		document.frmAddSolicitacao.txtBairroColeta.value = "Loading...";
		document.frmAddSolicitacao.txtMunicipioColeta.disabled = true;
		document.frmAddSolicitacao.txtMunicipioColeta.value = "Loading...";
		document.frmAddSolicitacao.txtEstadoColeta.disabled = true;
		document.frmAddSolicitacao.txtEstadoColeta.value = "Loading...";
		document.getElementById("btnBuscarCepColeta").style.cursor = 'wait';
		document.body.style.cursor = 'wait';
		//====================================================================
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				strRet = oAjax.responseText.split(";");
				document.frmAddSolicitacao.hiddenIntEnderecoCepColeta.value = strRet[0];
				document.frmAddSolicitacao.txtLogradouroColeta.value = strRet[6] + " "+ strRet[2];
				document.frmAddSolicitacao.txtBairroColeta.value = strRet[3];
				document.frmAddSolicitacao.txtMunicipioColeta.value = strRet[4];
				document.frmAddSolicitacao.txtEstadoColeta.value = strRet[5];
				document.frmAddSolicitacao.txtLogradouroColeta.disabled = false;
				document.frmAddSolicitacao.txtCompLogradouroColeta.disabled = false;
				document.frmAddSolicitacao.txtNumeroColeta.disabled = false;
				document.frmAddSolicitacao.txtBairroColeta.disabled = false;
				document.frmAddSolicitacao.txtMunicipioColeta.disabled = false;
				document.frmAddSolicitacao.txtEstadoColeta.disabled = false;
				document.getElementById("btnBuscarCepColeta").style.cursor = 'pointer';
				document.body.style.cursor = 'default';
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+document.frmAddSolicitacao.txtCepColeta.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	
	function preencheNovoEndereco(){
	
		var form = document.frmAddSolicitacao;
		var oAjax = Ajax();
		var strRet;
		
		//if (form.chkNovoEndereco.checked) {
			form.txtCepColeta.value = "";
			form.txtLogradouroColeta.value = "";
			form.txtCompLogradouroColeta.value = "";
			form.txtNumeroColeta.value = "";
			form.txtBairroColeta.value = "";
			form.txtMunicipioColeta.value = "";
			form.txtEstadoColeta.value = "";
			
			//lock texts
		//lock texts
		document.frmAddSolicitacao.txtLogradouroColeta.readOnly = false;
		document.frmAddSolicitacao.txtRespColContato.readOnly = false;
		document.frmAddSolicitacao.txtLogradouroColeta.readOnly = false;
		document.frmAddSolicitacao.txtCepColeta.readOnly = false;
		document.frmAddSolicitacao.txtCompLogradouroColeta.readOnly = false;
		document.frmAddSolicitacao.txtNumeroColeta.readOnly = false;
		document.frmAddSolicitacao.txtBairroColeta.readOnly = false;
		document.frmAddSolicitacao.txtMunicipioColeta.readOnly = false;
		document.frmAddSolicitacao.txtEstadoColeta.readOnly = false;
		document.frmAddSolicitacao.txtDDDContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtTelefoneContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtRamalContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtDepContatoRespColeta.readOnly = false;
			//}
	}

	//
	//Peterson de Aquino - 4/5/2014
	//Preenche com os dados do endereço padrão de coleta
	//
	function preencheEndColeta() {
	
		var form = document.frmAddSolicitacao;
		var oAjax = Ajax();
		var strRet;
		
		//if (form.chkMesmoEndereco.checked) {
			document.frmAddSolicitacao.txtLogradouroColeta.disabled = true;
			document.frmAddSolicitacao.txtLogradouroColeta.value = "Loading...";
			document.frmAddSolicitacao.txtCompLogradouroColeta.disabled = true;
			document.frmAddSolicitacao.txtNumeroColeta.disabled = true;
			document.frmAddSolicitacao.txtNumeroColeta.value = "Loading...";
			document.frmAddSolicitacao.txtBairroColeta.disabled = true;
			document.frmAddSolicitacao.txtBairroColeta.value = "Loading...";
			document.frmAddSolicitacao.txtMunicipioColeta.disabled = true;
			document.frmAddSolicitacao.txtMunicipioColeta.value = "Loading...";
			document.frmAddSolicitacao.txtEstadoColeta.disabled = true;
			document.frmAddSolicitacao.txtEstadoColeta.value = "Loading...";
			document.getElementById("btnBuscarCepColeta").style.cursor = 'wait';
			document.body.style.cursor = 'wait';

			oAjax.onreadystatechange = function() {
				if (oAjax.readyState == 4 && oAjax.status == 200) {
					strRet = oAjax.responseText.split(";");
					document.frmAddSolicitacao.hiddenIntEnderecoCepColeta.value = strRet[0];
					document.frmAddSolicitacao.txtCepColeta.value = strRet[1];
					document.frmAddSolicitacao.txtLogradouroColeta.value = strRet[2];
					document.frmAddSolicitacao.txtCompLogradouroColeta.value = strRet[7]
					document.frmAddSolicitacao.txtNumeroColeta.value = strRet[6];
					document.frmAddSolicitacao.txtBairroColeta.value = strRet[3];
					document.frmAddSolicitacao.txtMunicipioColeta.value = strRet[4];
					document.frmAddSolicitacao.txtEstadoColeta.value = strRet[5];
					document.frmAddSolicitacao.txtLogradouroColeta.disabled = false;
					document.frmAddSolicitacao.txtCompLogradouroColeta.disabled = false;
					document.frmAddSolicitacao.txtNumeroColeta.disabled = false;
					document.frmAddSolicitacao.txtBairroColeta.disabled = false;
					document.frmAddSolicitacao.txtMunicipioColeta.disabled = false;
					document.frmAddSolicitacao.txtEstadoColeta.disabled = false;
					document.getElementById("btnBuscarCepColeta").style.cursor = 'pointer';
					document.body.style.cursor = 'default';
				}	
			}
			
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getEndColeta&id="+document.frmAddSolicitacao.hiddenClienteId.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		/*
		} else {
			form.txtCepColeta.value = "";
			form.txtLogradouroColeta.value = "";
			form.txtCompLogradouroColeta.value = "";
			form.txtNumeroColeta.value = "";
			form.txtBairroColeta.value = "";
			form.txtMunicipioColeta.value = "";
			form.txtEstadoColeta.value = "";
		}*/
		
		//lock texts
		document.frmAddSolicitacao.txtLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtCepColeta.readOnly = true;
		document.frmAddSolicitacao.txtCompLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtNumeroColeta.readOnly = true;
		document.frmAddSolicitacao.txtBairroColeta.readOnly = true;
		document.frmAddSolicitacao.txtMunicipioColeta.readOnly = true;
		document.frmAddSolicitacao.txtEstadoColeta.readOnly = true;
		document.frmAddSolicitacao.txtRespColContato.readOnly = false;
		document.frmAddSolicitacao.txtDDDContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtTelefoneContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtRamalContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtDepContatoRespColeta.readOnly = false;
		
	}
	
	
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Preenchimento Automático do mesmo Endereço
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheMesmoEndereco() {
	
		var form = document.frmAddSolicitacao;
		var oAjax = Ajax();
		var strRet;
		
		//if (form.chkMesmoEndereco.checked) {
			document.frmAddSolicitacao.txtLogradouroColeta.disabled = true;
			document.frmAddSolicitacao.txtLogradouroColeta.value = "Loading...";
			document.frmAddSolicitacao.txtCompLogradouroColeta.disabled = true;
			document.frmAddSolicitacao.txtNumeroColeta.disabled = true;
			document.frmAddSolicitacao.txtBairroColeta.disabled = true;
			document.frmAddSolicitacao.txtBairroColeta.value = "Loading...";
			document.frmAddSolicitacao.txtMunicipioColeta.disabled = true;
			document.frmAddSolicitacao.txtMunicipioColeta.value = "Loading...";
			document.frmAddSolicitacao.txtEstadoColeta.disabled = true;
			document.frmAddSolicitacao.txtEstadoColeta.value = "Loading...";
			document.getElementById("btnBuscarCepColeta").style.cursor = 'wait';
			document.body.style.cursor = 'wait';

			oAjax.onreadystatechange = function() {
				if (oAjax.readyState == 4 && oAjax.status == 200) {
					strRet = oAjax.responseText.split(";");
					document.frmAddSolicitacao.hiddenIntEnderecoCepColeta.value = strRet[0];
					document.frmAddSolicitacao.txtCepColeta.value = document.frmAddSolicitacao.hiddenGetCepEnderecoComum.value;
					document.frmAddSolicitacao.txtLogradouroColeta.value = strRet[6] + " " +strRet[2];
					document.frmAddSolicitacao.txtCompLogradouroColeta.value = document.frmAddSolicitacao.hiddenGetCompLogradouroEnderecoCliente.value;
					document.frmAddSolicitacao.txtNumeroColeta.value = document.frmAddSolicitacao.hiddenGetNumeroEnderecoCliente.value;
					document.frmAddSolicitacao.txtBairroColeta.value = strRet[3];
					document.frmAddSolicitacao.txtMunicipioColeta.value = strRet[4];
					document.frmAddSolicitacao.txtEstadoColeta.value = strRet[5];
					document.frmAddSolicitacao.txtLogradouroColeta.disabled = false;
					document.frmAddSolicitacao.txtCompLogradouroColeta.disabled = false;
					document.frmAddSolicitacao.txtNumeroColeta.disabled = false;
					document.frmAddSolicitacao.txtBairroColeta.disabled = false;
					document.frmAddSolicitacao.txtMunicipioColeta.disabled = false;
					document.frmAddSolicitacao.txtEstadoColeta.disabled = false;
					document.getElementById("btnBuscarCepColeta").style.cursor = 'pointer';
					document.body.style.cursor = 'default';
				}	
			}
			
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+document.frmAddSolicitacao.hiddenGetCepEnderecoComum.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		/*
		} else {
			form.txtCepColeta.value = "";
			form.txtLogradouroColeta.value = "";
			form.txtCompLogradouroColeta.value = "";
			form.txtNumeroColeta.value = "";
			form.txtBairroColeta.value = "";
			form.txtMunicipioColeta.value = "";
			form.txtEstadoColeta.value = "";
		}*/
		
		//lock texts
		document.frmAddSolicitacao.txtLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtCepColeta.readOnly = true;
		document.frmAddSolicitacao.txtCompLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtNumeroColeta.readOnly = true;
		document.frmAddSolicitacao.txtBairroColeta.readOnly = true;
		document.frmAddSolicitacao.txtMunicipioColeta.readOnly = true;
		document.frmAddSolicitacao.txtEstadoColeta.readOnly = true;
		document.frmAddSolicitacao.txtRespColContato.readOnly = false;
		document.frmAddSolicitacao.txtDDDContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtTelefoneContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtRamalContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtDepContatoRespColeta.readOnly = false;
	}
	
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Autentica a edição do endereço de Coleta do Cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function authenticateUpdateAdress() {
		if (document.frmAddSolicitacao.txtCepColeta.value == "") {
			alert("Preencha o campo CEP de Coleta");
			return false;
		}	
		if (document.frmAddSolicitacao.txtLogradouroColeta.value == "") {
			alert("Preencha o campo Logradouro");
			return false;
		}
		if (document.frmAddSolicitacao.txtNumeroColeta.value == "") {
			alert("Preencha o campo Numero do Endereço");
			return false;
		}
		if (document.frmAddSolicitacao.txtBairroColeta.value == "") {
			alert("Preencha o campo Bairro");
			return false;
		}
		if (document.frmAddSolicitacao.txtMunicipioColeta.value == "") {
			alert("Preencha o campo Município");
			return false;
		}
		if (document.frmAddSolicitacao.txtEstadoColeta.value == "") {
			alert("Preencha o campo Estado");
			return false;
		}
		if (document.frmAddSolicitacao.txtRespColContato.value == "") {
			alert("Preencha o campo do Contato responsável pela Coleta!");
			form.txtRespColContato.focus();
			return false;
		}
		if (document.frmAddSolicitacao.txtDDDContatoRespColeta.value == "") {
				alert("Preencha o campo do DDD do responsável pela Coleta!");
				form.txtDDDContatoRespColeta.focus();
				return false;
		}
		if (isNaN(document.frmAddSolicitacao.txtDDDContatoRespColeta.value)) {
			alert("Preencha o campo DDD do Contato somente com dados numéricos!");
			form.txtDDDContatoRespColeta.focus();
			return false;
		}
		if (document.frmAddSolicitacao.txtDDDContatoRespColeta.value.length < 2) {
			alert("Preencha o campo do DDD do Contato com no mínimo 2 caracteres válidos!");
			form.txtDDDContatoRespColeta.focus();
			return false;
		}
		if (document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value == "") {
				alert("Preencha o campo do Telefone do responsável pela Coleta!");
				form.txtTelefoneContatoRespColeta.focus();
				return false;
		}
		if (isNaN(document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value)) {
			alert("Preencha o campo Telefone do Contato somente com dados numéricos!");
			form.txtTelefoneContatoRespColeta.focus();
			return false;
		}
		if (document.frmAddSolicitacao.txtTelefoneContatoRespColeta.value.length < 8) {
			alert("Preencha o campo do Telefone do Contato com 8 caracteres válidos!");
			form.txtTelefoneContatoRespColeta.focus();
			return false;
		}
		document.frmAddSolicitacao.hiddenActionForm.value = "1";
		document.frmAddSolicitacao.submit();
	}
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Carrega as informações se for o mesmo endereço para coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function loadInfoSameAdress() {
		if (document.frmAddSolicitacao.txtCepColeta.value == document.frmAddSolicitacao.hiddenGetCepEnderecoComum.value) {
			document.frmAddSolicitacao.chkMesmoEndereco.checked = true;
			document.frmAddSolicitacao.txtCompLogradouroColeta.value = document.frmAddSolicitacao.hiddenGetCompLogradouroEnderecoCliente.value;
		}		
	}
	
	function loadClear() {
		var form = document.frmAddSolicitacao;
		var oAjax = Ajax();
		var strRet;
	
		form.txtCepColeta.value = "";
		form.txtLogradouroColeta.value = "";
		form.txtCompLogradouroColeta.value = "";
		form.txtNumeroColeta.value = "";
		form.txtBairroColeta.value = "";
		form.txtMunicipioColeta.value = "";
		form.txtEstadoColeta.value = "";	
		form.txtRespColContato.value = "";
		form.txtDDDContatoRespColeta.value = "";
		form.txtTelefoneContatoRespColeta.value = "";
		
		document.getElementsByName('txtRespColContato').readOnly = true;
		document.getElementsByName('txtLogradouroColeta').readOnly = true;
		document.getElementsByName('txtCepColeta').readOnly = true;
		document.getElementsByName('txtCompLogradouroColeta').readOnly = true;
		document.getElementsByName('txtNumeroColeta').readOnly = true;
		document.getElementsByName('txtBairroColeta').readOnly = true;
		document.getElementsByName('txtMunicipioColeta').readOnly = true;
		document.getElementsByName('txtEstadoColeta').readOnly = true;
		document.getElementsByName('txtDDDContatoRespColeta').readOnly = true;
		document.getElementsByName('txtTelefoneContatoRespColeta').readOnly = true;
		document.getElementsByName('txtRamalContatoRespColeta').readOnly = true;
		document.getElementsByName('txtDepContatoRespColeta').readOnly = true;
		
	}
	
	

