// JavaScript Document
/*
'|--------------------------------------------------------------------
'| Arquivo: frmCadCliente.js																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação : 15/04/2007																		 
'| Descrição: Arquivo de Formulário para cadastro de Cliente (Javascript)
'|--------------------------------------------------------------------
*/
	
	var errForm = false;
	var	msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";

	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Verifica se a coleta tem que ser domiciliar ou de ponto de coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function ckeckColeta() {
		var form = document.frmCadCliente;
		if (form.cbCategorias.value == -1) {
			alert("Escolha uma Categoria!");
			return false;
		} else {
			if (isNaN(form.txtQtdCartuchos.value)) {
				alert("Digite somente números na quantidade de cartuchos!");
				return false;
			}
			if (form.txtQtdCartuchos.value == 0 || form.txtQtdCartuchos.value == "") {
				alert("Digite a quantidade de cartuchos a ser coletada!");
				return false;
			} else {
				if (form.hiddenControleColeta.value == 0) {
					form.hiddenTypeColeta.value = 0;															
				} else {
					if (parseInt(form.txtQtdCartuchos.value) < parseInt(form.hiddenMinCartuchos.value)) {
						form.hiddenTypeColeta.value = 0;			
					} else {
						form.hiddenTypeColeta.value = 1;
					}
				}
				if (form.hiddenTypeColeta.value == 0) {
					document.getElementById("tableCadClienteCategoria").style.display = 'none';
					document.getElementById("tableCadClientePontoColeta").style.display = 'block';
					document.frmCadCliente.btnNextToCadEmpresaPontoColeta.disabled = true;
				} else {
					document.getElementById("tableCadClienteCategoria").style.display = 'none';
					document.getElementById("tableCadCliente").style.display = 'block';
				}        
			}
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Conforme a escolha do usuário
	// Verifica se o usuario é coleta Domiciliar ou coleta em Ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function setTypeColeta() {
		var oAjax = Ajax();
		var form = document.frmCadCliente;
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				form.hiddenControleColeta.value = oAjax.responseText;
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=gettypecoleta&id="+form.cbCategorias.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Atualiza o valor do campo hidden para seja guardado o valor do mínimo de cartuchos
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function updateMinimo() {
		var oAjax = Ajax();
		var form = document.frmCadCliente;
		
		if (form.cbCategorias.value != -1) {
			oAjax.onreadystatechange = function() {
				if (oAjax.readyState == 4 && oAjax.status == 200) {
					form.hiddenMinCartuchos.value = oAjax.responseText;
					setTypeColeta();
				}
			}
			
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getminimo&id="+form.cbCategorias.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Retorno do Usuario quando o mesmo é Coleta Domiciliar
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function returnCadCategoriaDomiciliar() {
		document.getElementById("tableCadCliente").style.display = 'none';
		if (!checkTypeColeta()) {
			document.getElementById("tableCadClientePontoColeta").style.display = 'block';
		} else {
			document.getElementById("tableCadClienteCategoria").style.display = 'block';
		}				
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Retorno do Usuario quando o mesmo é Coleta Ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function returnCadCategoriaPontoColeta() {
		document.getElementById("tableCadClientePontoColeta").style.display = 'none';
		document.getElementById("tableCadClienteCategoria").style.display = 'block';
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Preenchimento do Endereço da Empresa
	// É validado os campos referente ao Cadastro da Empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showClienteEndereco() {
		var form = document.frmCadCliente;
		if (form.radioPessoa[0].checked) {
			if (form.txtNome.value == "") {
				msgErrForm += "Campo: Nome\n";					
				errForm = true;
			}
			validateCPF();
			validaDDD();
			validaTelefone();
			if (!errForm) {
				document.getElementById('tableCadClienteEndereco').style.display = 'block';
				document.getElementById('tableCadCliente').style.display = 'none';
				document.getElementById('tableCadClienteCategoria').style.display = 'none';
				form.txtNumero.value = "";
				form.chkMesmoEndereco.checked = false;
				form.txtCepColeta.value = "";
				form.txtLogradouroColeta.value = "";
				form.txtCompLogradouroColeta.value = "";
				form.txtNumeroColeta.value = "";
				form.txtBairroColeta.value = "";
				form.txtMunicipioColeta.value = "";
				form.txtEstadoColeta.value = "";
			} else {
				alert(msgErrForm);	
				msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";
			}
			errForm = false;
		} else {
			if (form.radioPessoa[1].checked) {
				if (form.txtRazaoSocial.value == "") {
					msgErrForm += "Campo: Razão Social\n";
					errForm = true;
				}
				if (form.txtNomeFantasia.value == "") {
					msgErrForm += "Campo: Nome Fantasia\n";
					errForm = true;
				}
				validaCnpj();
				if (updateDisplayIE()) {
					if (form.txtIE.value == "") {
						msgErrForm += "Campo: Inscrição Estadual\n";
						errForm = true;
					}
					if (form.txtIE.value.length < 15) {
						msgErrForm += "Campo: Inscrição Estadual\n";
						errForm = true;
					}
				}
				validaDDD();
				validaTelefone();
				if (!errForm) {
					document.getElementById('tableCadClienteEndereco').style.display = 'block';
					document.getElementById('tableCadCliente').style.display = 'none';
					document.getElementById('tableCadClienteCategoria').style.display = 'none';
					form.txtNumero.value = "";
					form.chkMesmoEndereco.checked = false;
					form.txtCepColeta.value = "";
					form.txtLogradouroColeta.value = "";
					form.txtCompLogradouroColeta.value = "";
					form.txtNumeroColeta.value = "";
					form.txtBairroColeta.value = "";
					form.txtMunicipioColeta.value = "";
					form.txtEstadoColeta.value = "";
				} else {
					alert(msgErrForm);	
					msgErrForm = "Os seguintes campos foram preenchidos incorretamente!\n";
				}
				errForm = false;
			}
		}
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Direciona para o cadastro da empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showCadCliente() {
		if (document.frmCadCliente.txtCepConsultaPonto.value == "") {
			alert("Por favor preencha o campo de Cep para Busca dos Pontos de Coleta!");
			return false;
		}
		if (checkChangePontoColeta()) {
			document.getElementById('tableCadClientePontoColeta').style.display = 'none';
			document.getElementById('tableCadCliente').style.display = 'block';
		} else {
			alert("Escolha um Ponto de Coleta para que seja feita a Solicitação!");
		}
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Check se foi selecionado algum ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkChangePontoColeta() {
		var bErr = false;
		var bSelected = false;
		for (var i=0;i <= parseInt(document.frmCadCliente.hiddenIntPontoColeta.value);i++) {
			if (!document.getElementById("radioCheckPonto"+i).checked) {
				bErr = true;
			} else {
				bSelected = true;
				document.frmCadCliente.hiddenIntChangePontoColeta.value = document.getElementById("radioCheckPonto"+i).value;
			}
		}
		
		if (bErr && !bSelected) {
			return false
		} else {
			return true;	
		}
 	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Check se foi selecionado algum ponto de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkTypeContinue() {
		if (document.frmCadCliente.txtNumero.value == "") {
			alert("Preencha o número da Localização da Empresa");
			document.frmCadCliente.txtNumero.focus();
			return false;
		}
		if (document.frmCadCliente.txtCep.value == "") {
			alert("Preencha o campo Cep para Busca do Endereço de Coleta!");
			return false;
		}
		if (isNaN(document.frmCadCliente.txtCep.value)) {
			alert("Preencha o Cep somente com números!");
			return false;
		}
		if (document.frmCadCliente.txtCep.value.length < 8) {
			alert("Preencha corretamente o Campo Cep!");
			return false;
		} else {
			document.getElementById("tableCadClienteEndereco").style.display = 'none';
			if (!checkTypeColeta()) {
				document.getElementById("tableCadClienteContato").style.display = 'block';
			} else {
				document.getElementById("tableCadClienteEnderecoColeta").style.display = 'block';
			}
		}
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Check qual tipo selecionado de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkTypeColeta() {
		if (document.frmCadCliente.hiddenTypeColeta.value == 0) {
			return false;
		} else {
			return true;
		}
	}
	//=========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Validação do CNPJ da Empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaCnpj() {

		var form = document.frmCadCliente;
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
	
		cnpj = form.txtNCNPJ.value;

		digVer2 = cnpj.charAt(cnpj.length - 1);
		digVer1 = cnpj.charAt(cnpj.length - 2);
		
		if (form.txtNCNPJ.value == "") {
			msgErrForm += "Campo: CNPJ\n";
			errForm = true;
			return false;
		}
		if (form.txtNCNPJ.value.indexOf('/') == -1) {
			msgErrForm += "Campo: CNPJ\n";
			errForm = true;
			return false;
		}
		if (form.txtNCNPJ.value.length < 18) {
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
		var form = document.frmCadCliente;

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
	// Busca De Pontos de Coleta do cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showClientePostoColeta() {
		var form = document.frmCadCliente;
		var oAjax = Ajax();
		var strRet = "";

		document.body.style.cursor='wait';
		//form.btnBuscarCepUser.disabled = true;

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
			document.frmCadCliente.btnBuscarCepUser.disabled = 'true';
			document.frmCadCliente.btnNextToCadEmpresaPontoColeta.disabled = '';
			document.frmCadCliente.txtCepConsultaPonto.value;
			oAjax.onreadystatechange = function() {
				if (oAjax.readyState == 4 && oAjax.status == 200) {
//					alert(oAjax.responseText);
					document.frmCadCliente.btnBuscarCepUser.disabled = false;
					strRet = oAjax.responseText.split(";");
					document.getElementById("titTableListPontoColeta").style.display = 'block';
					document.getElementById("tableListPontoColeta").innerHTML = strRet[0];
					document.frmCadCliente.hiddenIntPontoColeta.value = strRet[1];
					if (document.frmCadCliente.hiddenIntPontoColeta.value == -1) {
						alert("Não foi encontrado nenhum ponto de coleta no CEP digitado, favor tente outro CEP!");
						document.frmCadCliente.btnNextToCadEmpresaPontoColeta.disabled = 'true';
						document.frmCadCliente.txtCepConsultaPonto.value = "";
						document.frmCadCliente.txtCepConsultaPonto.focus();
						return false;
					} else {
						document.frmCadCliente.btnNextToCadEmpresaPontoColeta.disabled = '';
						if (!confirm("Este é o endereço do(s) ponto(s) de coleta mais próximo(s), onde você deverá entregar os cartuchos vazios.\nDeseja prosseguir com a solicitação?")) {
							alert("Grato pelo contato!");
							window.location.href="index.asp";
						} else {
						        document.body.style.cursor='default';
						}
					}
				} 	
			}
			oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getlistpontocoleta&id="+document.frmCadCliente.txtCepConsultaPonto.value, true);
			oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
			oAjax.send(null);
		}
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function loadCepColeta() {
		var oAjax = Ajax();
		var strRet = "";

		if (document.frmCadCliente.txtCepColeta.value == "") {
			alert("Preencha o campo Cep para Busca do Endereço de Coleta!");
			return false;
		}
		if (isNaN(document.frmCadCliente.txtCepColeta.value)) {
			alert("Preencha o Cep somente com números!");
			return false;
		}
		if (document.frmCadCliente.txtCepColeta.value.length < 8) {
			alert("Preencha corretamente o Campo Cep!");
			return false;
		}
		//====================================================================
		// Bloqueio dos campos de consulta
		//====================================================================
		document.frmCadCliente.chkMesmoEndereco.checked = false;
		document.frmCadCliente.btnBuscarCepComum.disabled = true;
		document.frmCadCliente.txtLogradouroColeta.disabled = true;
		document.frmCadCliente.txtLogradouroColeta.value = "Carregando...";
		document.frmCadCliente.txtCompLogradouroColeta.disabled = true;
		document.frmCadCliente.txtNumeroColeta.disabled = true;
		document.frmCadCliente.txtBairroColeta.disabled = true;
		document.frmCadCliente.txtBairroColeta.value = "Carregando...";
		document.frmCadCliente.txtMunicipioColeta.disabled = true;
		document.frmCadCliente.txtMunicipioColeta.value = "Carregando...";
		document.frmCadCliente.txtEstadoColeta.disabled = true;
		document.frmCadCliente.txtEstadoColeta.value = "Carregando...";
		document.frmCadCliente.txtCompLogradouroColeta.value = "";
		document.frmCadCliente.txtNumeroColeta.value = "";
		document.getElementById("btnBuscarCepComum").style.cursor = 'wait';
		document.frmCadCliente.btnNextToContatoMaster.style.cursor = 'wait';
		document.frmCadCliente.btnBackCadClienteColeta.style.cursor = 'wait';
		document.body.style.cursor = 'wait';
		//====================================================================
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				strRet = oAjax.responseText.split(";");
				document.frmCadCliente.btnBuscarCepComum.disabled = false;
				document.frmCadCliente.hiddenIntEnderecoCepColeta.value = strRet[0];
				document.frmCadCliente.txtLogradouroColeta.value = strRet[2];
				document.frmCadCliente.txtBairroColeta.value = strRet[3];
				document.frmCadCliente.txtMunicipioColeta.value = strRet[4];
				document.frmCadCliente.txtEstadoColeta.value = strRet[5];
				document.frmCadCliente.txtLogradouroColeta.disabled = false;
				document.frmCadCliente.txtCompLogradouroColeta.disabled = false;
				document.frmCadCliente.txtNumeroColeta.disabled = false;
				document.frmCadCliente.txtBairroColeta.disabled = false;
				document.frmCadCliente.txtMunicipioColeta.disabled = false;
				document.frmCadCliente.txtEstadoColeta.disabled = false;
				document.getElementById("btnBuscarCepComum").style.cursor = 'pointer';
				document.frmCadCliente.btnNextToContatoMaster.style.cursor = 'default';
				document.frmCadCliente.btnBackCadClienteColeta.style.cursor = 'default';
				document.body.style.cursor = 'default';
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+document.frmCadCliente.txtCepColeta.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function loadCepComum() {
		var oAjax = Ajax();
		var strRet = "";

		if (document.frmCadCliente.txtCep.value == "") {
			alert("Preencha o campo Cep para Busca do Endereço de Coleta!");
			return false;
		}
		if (isNaN(document.frmCadCliente.txtCep.value)) {
			alert("Preencha o Cep somente com números!");
			return false;
		}
		if (document.frmCadCliente.txtCep.value.length < 8) {
			alert("Preencha corretamente o Campo Cep!");
			return false;
		}
		//====================================================================
		// Bloqueio dos campos de consulta
		//====================================================================
//		document.frmCadCliente.chkMesmoEndereco.checked = false;
		document.frmCadCliente.btnBuscarCepColeta.disabled = true;
		document.frmCadCliente.txtLogradouro.disabled = true;
		document.frmCadCliente.txtLogradouro.value = "Carregando...";
		document.frmCadCliente.txtCompLogradouro.disabled = true;
		document.frmCadCliente.txtNumero.disabled = true;
		document.frmCadCliente.txtBairro.disabled = true;
		document.frmCadCliente.txtBairro.value = "Carregando...";
		document.frmCadCliente.txtMunicipio.disabled = true;
		document.frmCadCliente.txtMunicipio.value = "Carregando...";
		document.frmCadCliente.txtEstado.disabled = true;
		document.frmCadCliente.txtEstado.value = "Carregando...";
		document.frmCadCliente.txtCompLogradouro.value = "";
		document.frmCadCliente.txtNumero.value = "";
//		document.getElementById("btnBuscarCepColeta").style.cursor = 'wait';
//		document.frmCadCliente.btnNextToEnderecoColeta.style.cursor = 'wait';
//		document.frmCadCliente.btnBackCadCliente.style.cursor = 'wait';
		document.body.style.cursor = 'wait';
		//====================================================================
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				strRet = oAjax.responseText.split(";");
				document.frmCadCliente.btnBuscarCepColeta.disabled = false;				
				document.frmCadCliente.hiddenIntEnderecoCep.value = strRet[0];				
				document.frmCadCliente.txtLogradouro.value = strRet[2];
				document.frmCadCliente.txtBairro.value = strRet[3];
				document.frmCadCliente.txtMunicipio.value = strRet[4];
				document.frmCadCliente.txtEstado.value = strRet[5];
				document.frmCadCliente.txtLogradouro.disabled = false;
				document.frmCadCliente.txtCompLogradouro.disabled = false;
				document.frmCadCliente.txtNumero.disabled = false;
				document.frmCadCliente.txtBairro.disabled = false;
				document.frmCadCliente.txtMunicipio.disabled = false;
				document.frmCadCliente.txtEstado.disabled = false;
//				document.getElementById("btnBuscarCepColeta").style.cursor = 'pointer';
				//document.frmCadCliente.btnNextToEnderecoColeta.style.cursor = 'default';
				//document.frmCadCliente.btnBackCadCliente.style.cursor = 'default';
				document.body.style.cursor = 'default';
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcependereco&id="+document.frmCadCliente.txtCep.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Busca do Endereço para Preenchimento automático do Endereço de Coleta
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheEnderecoToCep() {
		var oAjax = Ajax();
		var form = document.frmCadCliente;
		var strRet = "";
		
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				strRet = oAjax.responseText.split(";");
				strRet[6] = strRet[6].replace("        ",'');
				form.hiddenIntEnderecoCep.value = strRet[0];
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
	// Preenchimento Automático do mesmo Endereço
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function preencheMesmoEndereco() {
		var form = document.frmCadCliente;
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
	// Validação do Ambiente de Contato com o cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function validaCadClienteContato() {

	    var form = document.frmCadCliente;

        //peterson alteração: 11-5-2014
	    //verifica se valida o nome e o CPF para pessoa física
	    if (form.radioPessoa[0].checked) {
	        if (form.txtNome.value=="") {
	            alert('Preencha o campo Nome!');
	            return false;
	        }
	        if (form.txtCPF.value = "") {
	            alert('Preencha corretamente o campo CPF');
	            return false;
	        }
	    }
	    if (form.radioPessoa[1].checked) {
	        if (form.txtRazaoSocial.value == "") {
	            alert('Preencher o campo Razão Social!');
	            return false;
	        }
	        if (form.txtNomeFantasia.value="") {
	            alert('Preencger o campo Nome fantasia');
	            return false;
	        }
	        if (form.txtNCNPJ.value = "") {
	            alert('Preencher o campo CNPJ!');
	            return false;
	        }
	        if (form.txtIE.value="") {
	            alert('Preencher o campo Inscrição Estadual. Caso não tenha, preencha com o texto: ISENTO');
	            return false;
	        }

	    }
		
		if (form.txtContatoColeta.value == "") {
			alert("Preencha o campo Contato!");
			return false;
		}
		if (form.txtUsuario.value == "") {
			alert("Preencha o campo Usuario!");
			return false;
		}
		if (form.txtUsuario.value.length < 6) {
			alert("Preencha o campo Usuário com no mínimo 6 caracteres!");
			return false;
		}
		if (form.txtSenha.value == "") {
			alert("Preencha o campo Senha!");
			return false;
		}
		if (form.txtSenhaconfirma.value == "") {
		    alert("Preencha o campo Confirmação de Senha!");
		    return false;
		}
		if (form.txtSenha.value != form.txtSenhaconfirma.value) {
		    alert('As senhas não conferem, favor digitar novamente!');
		    return false;
		}
		if (form.txtSenha.value.length < 6) {
			alert("Preencha o campo Senha com no mínimo 6 caracteres!");
			return false;
		}

  		if (!is_email(form.txtEmail.value))
		{
		    alert("Email inválido!");
		    return false;
		}

		
		/*
		if (form.txtEmail.value == "") {
			alert("Preencha o campo Email!");
			return false;
		} 
		if (form.txtEmail.value.indexOf('@') == -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('@.') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('.@') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('.') == -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('com') == -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('[') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf(']') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('(') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf(')') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('/') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('\\') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('..') != -1) {
			alert("Email inválido!");
			return false;
		}
		if (form.txtEmail.value.indexOf('com.com') != -1) {
			alert("Email inválido!");
			return false;
		}
		*/

		checkUserContato();
	}

	function is_email(email)
	{
	  er = /^[a-zA-Z0-9][a-zA-Z0-9\._-]+@([a-zA-Z0-9\._-]+\.)[a-zA-Z-0-9]{2}/;
	  
	  if(er.exec(email))
		{
		  return true;
		} else {
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
		var form = document.frmCadCliente;

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
	// Retorno do Contato corretamente de acordo com o tipo de Cliente
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkReturnEndCorrespondente() {
		var form = document.frmCadCliente;
		if (form.hiddenTypeColeta.value == 0) {
			document.getElementById("tableCadClienteContato").style.display = 'none';
			document.getElementById("tableCadClienteEndereco").style.display = 'block';
		} else {
			document.getElementById("tableCadClienteContato").style.display = 'none';
			document.getElementById("tableCadClienteEnderecoColeta").style.display = 'block';
		}        
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Retorno do Endereço do Cliente para o Cadastro da Empresa
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function returnCadCliente() {
		document.getElementById("tableCadClienteEndereco").style.display = 'none';
		document.getElementById("tableCadCliente").style.display = 'block';
	}

	function returnCadClienteColeta() {
		document.getElementById("tableCadClienteEnderecoColeta").style.display = 'none';
		document.getElementById("tableCadClienteEndereco").style.display = 'block';
	}


	function returnClienteEnderecoPonto() {
		document.getElementById("tableCadClientePontoColeta").style.display = 'none';
		document.getElementById("tableCadClienteEndereco").style.display = 'block';
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Prossegue para o cadastro do Contato no Tipo Coleta Domiciliar
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function showCadClienteContato() {
		var form = document.frmCadCliente;
		if (form.txtCepColeta.value == "") {
			alert("Preencha o campo de Cep do Endereço de Coleta!");
			form.txtCepColeta.focus();
			return false;
		}
		if (form.txtNumeroColeta.value == "") {
			alert("Preecha o campo número do Endereço de Coleta!");
			form.txtNumeroColeta.focus();
			return false;
		}
		if (form.txtContatoRespColeta.value == "") {
			alert("Preencha o campo do Contato responsável pela Coleta!");
			form.txtContatoRespColeta.focus();
			return false;
		}
		if (form.txtDDDContatoRespColeta.value == "") {
				alert("Preencha o campo do DDD do responsável pela Coleta!");
				form.txtDDDContatoRespColeta.focus();
				return false;
		}
		if (isNaN(form.txtDDDContatoRespColeta.value)) {
			alert("Preencha o campo DDD do Contato somente com dados numéricos!");
			form.txtDDDContatoRespColeta.focus();
			return false;
		}
		if (form.txtDDDContatoRespColeta.value.length < 2) {
			alert("Preencha o campo do DDD do Contato com no mínimo 2 caracteres válidos!");
			form.txtDDDContatoRespColeta.focus();
			return false;
		}
		if (form.txtTelefoneContatoRespColeta.value == "") {
				alert("Preencha o campo do Telefone do responsável pela Coleta!");
				form.txtTelefoneContatoRespColeta.focus();
				return false;
		}
		if (isNaN(form.txtTelefoneContatoRespColeta.value)) {
			alert("Preencha o campo Telefone do Contato somente com dados numéricos!");
			form.txtTelefoneContatoRespColeta.focus();
			return false;
		}
		if (form.txtTelefoneContatoRespColeta.value.length < 8) {
			alert("Preencha o campo do Telefone do Contato com 8 caracteres válidos!");
			form.txtTelefoneContatoRespColeta.focus();
			return false;
		}
		document.getElementById("tableCadClienteEnderecoColeta").style.display = 'none';
		document.getElementById("tableCadClienteContato").style.display = 'block';
	}
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	//==========================================================================================================
	// Verifica se o usuário já está cadastrado
	//==========================================================================================================
	////////////////////////////////////////////////////////////////////////////////////////////////////////////
	function checkUserContato() {
		var oAjax = Ajax();
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				if (oAjax.responseText == "true") {
					alert("Usuário ou Senha já cadastrado!");
				} else {
					document.frmCadCliente.submit();
				}
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcheckusercontato&id=0&user="+document.frmCadCliente.txtUsuario.value+"&senha="+document.frmCadCliente.txtSenha.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	
	function checkUsuario() {
		var oAjaxUsuario = Ajax();
		oAjaxUsuario.onreadystatechange = function() {
			if (oAjaxUsuario.readyState == 4 && oAjaxUsuario.status == 200) {
				if (oAjaxUsuario.responseText == "true") {
					alert("Usuário já cadastrado. Favor cadastre outro usuário");
					document.frmCadCliente.txtUsuario.focus();
					return;
				}
			}
		}
		
		oAjaxUsuario.open("GET", "ajax/frmCadCliente.asp?sub=getcheckusuario&id=0&user="+document.frmCadCliente.txtUsuario.value, true);
		oAjaxUsuario.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjaxUsuario.send(null);
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
					document.frmCadCliente.txtNCNPJ.focus();
					return false;
				} else {
					return true;
				}
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcheckcnpjempresa&id="+document.frmCadCliente.txtNCNPJ.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}

	function checkCPF() {
		var oAjax = Ajax();
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				if (oAjax.responseText == "true") {
					alert("CPF já cadastrado. Favor cadastrar outro CPF!");
					document.frmCadCliente.txtCPF.focus();
					return false;
				} else {
					return true;
				}
			}
		}
		
		oAjax.open("GET", "ajax/frmCadCliente.asp?sub=getcheckcpf&id="+document.frmCadCliente.txtCPF.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
	
	function checkPessoa() {

	    var form = document.frmCadCliente;

		if (form.radioPessoa[0].checked) {
			document.getElementById("razaosocial").style.display = 'none';
			document.getElementById("nomefantasia").style.display = 'none';
			document.getElementById("cnpj").style.display = 'none';
			document.getElementById("inscestadual").style.display = 'none';
			document.getElementById("nome").style.display = 'block';
			document.getElementById("cpf").style.display = 'block';
			//document.getElementById("possuiinscestadual").style.display = 'none';
		} else {
			if (form.radioPessoa[1].checked) {
				document.getElementById("razaosocial").style.display = 'block';
				document.getElementById("nomefantasia").style.display = 'block';
				document.getElementById("cnpj").style.display = 'block';
				document.getElementById("inscestadual").style.display = 'block';
				document.getElementById("nome").style.display = 'none';
				document.getElementById("cpf").style.display = 'none';
				//document.getElementById("possuiinscestadual").style.display = 'block';
			}	
		}
	}
	
	function validateCPF() {
		var form = document.frmCadCliente;
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
	
	function updateDisplayIE() {
		var form = document.frmCadCliente;
			if (form.hasEstadual[0].checked) {
				document.getElementById("inscestadual").style.display = 'block';
				return true;
			}
			if (form.hasEstadual[1].checked) {
				document.getElementById("inscestadual").style.display = 'none';	
				return false;
			}
	}
	
	function cnpj_format(cnpj) {
		var form = document.frmCadCliente;
		if (cnpj.value.length == 2 || cnpj.value.length == 6) {
			form.txtNCNPJ.value += ".";	
		}
		if (cnpj.value.length == 10) {
			form.txtNCNPJ.value += "/";		
		}
		if (cnpj.value.length == 15) {
			form.txtNCNPJ.value += "-";		
		}
	}
	
	function cpf_format(cpf) {
		var form = document.frmCadCliente;
		if (cpf.value.length == 3 || cpf.value.length == 7) {
			form.txtCPF.value += ".";	
		}
		if (cpf.value.length == 11) {
			form.txtCPF.value += "-";		
		}
	}
	
	function keypressIE(value) {
		var form = document.frmCadCliente;
		if (value.length == 3 || value.length == 7 || value.length == 11) {
			form.txtIE.value += ".";	
		}
	}

