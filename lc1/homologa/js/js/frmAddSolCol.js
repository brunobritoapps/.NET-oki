	// JavaScript Document
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
					document.frmAddSolicitacao.txtCepColeta.value = document.frmAddSolicitacao.hiddenCEPCol.value;
					document.frmAddSolicitacao.txtLogradouroColeta.value = document.frmAddSolicitacao.hiddenLogrCol.value;;
					document.frmAddSolicitacao.txtCompLogradouroColeta.value = document.frmAddSolicitacao.hiddenComplCol.value;
					document.frmAddSolicitacao.txtNumeroColeta.value = document.frmAddSolicitacao.hiddenNumCol.value;
					document.frmAddSolicitacao.txtBairroColeta.value = document.frmAddSolicitacao.hiddenBaiCol.value;
					document.frmAddSolicitacao.txtMunicipioColeta.value = document.frmAddSolicitacao.hiddenMunCol.value;
					document.frmAddSolicitacao.txtEstadoColeta.value = document.frmAddSolicitacao.hiddenEstCol.value;
					
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
		document.frmAddSolicitacao.txtContatoRespColeta.readOnly = true;
		document.frmAddSolicitacao.txtLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtCepColeta.readOnly = true;
		document.frmAddSolicitacao.txtCompLogradouroColeta.readOnly = true;
		document.frmAddSolicitacao.txtNumeroColeta.readOnly = true;
		document.frmAddSolicitacao.txtBairroColeta.readOnly = true;
		document.frmAddSolicitacao.txtMunicipioColeta.readOnly = true;
		document.frmAddSolicitacao.txtEstadoColeta.readOnly = true;
		document.frmAddSolicitacao.txtContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtDDDContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtTelefoneContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtRamalContatoRespColeta.readOnly = false;
		document.frmAddSolicitacao.txtDepContatoRespColeta.readOnly = false;		
		
	}