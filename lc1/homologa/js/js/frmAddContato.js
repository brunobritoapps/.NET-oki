// JavaScript Document
/*
'|--------------------------------------------------------------------
'| Arquivo: frmAddContato.js																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação : 15/04/2007																		 
'| Descrição: Arquivo de Formulário para cadastro de Contato (Javascript)
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
//==========================================================================================================
////////////////////////////////////////////////////////////////////////////////////////////////////////////
//==========================================================================================================
// Validação do Ambiente de Contato com o cliente
//==========================================================================================================
////////////////////////////////////////////////////////////////////////////////////////////////////////////
function validaCadClienteContato() {
	var form = document.frmAddContato;
	
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
	if (!form.radioIsMaster[0].checked && !form.radioIsMaster[1].checked) {
		alert("Escolha o Tipo de Usuário (Master)!");
		return false;
	}
	document.frmAddContato.submit();
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
// Verifica se o usuário já está cadastrado
//==========================================================================================================
////////////////////////////////////////////////////////////////////////////////////////////////////////////
//function checkUserContato() {
//	var oAjax = Ajax();
//	oAjax.onreadystatechange = function() {
//		if (oAjax.readyState == 4 && oAjax.status == 200) {
//			if (oAjax.responseText == "true") {
//				alert("Usuário já cadastrado. Favor cadastre outro Usuário!");
//			} else {
//				document.frmAddContato.submit();
//			}
//		}
//	}
//	
//	oAjax.open("GET", "ajax/frmAddContato.asp?sub=getcheckusercontato&id=0&user="+document.frmAddContato.txtUsuario.value+"&senha="+document.frmAddContato.txtSenha.value, true);
//	oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
//	oAjax.send(null);
//}
//==========================================================================================================

function redirActionContato() {
	var contCheckFalse = 0;
	var elemento;
	for (var cont=0; cont < document.frmAddContato.hiddenIntContatos.value; cont++) {
		if (!document.getElementById("checkContato"+cont).checked) {
			contCheckFalse++;
		} else {
			elemento = document.getElementById("checkContato"+cont);
		}	
	}
	if (contCheckFalse == document.frmAddContato.hiddenIntContatos.value && !document.getElementById("checkContato"+document.frmAddContato.hiddenIntContatos.value).checked) {
		alert("Escolha um Contato para executar a Ação escolhida!");
		return false;
	} else {
		if (document.getElementById("checkContato"+document.frmAddContato.hiddenIntContatos.value).checked) {
			elemento = document.getElementById("checkContato"+document.frmAddContato.hiddenIntContatos.value);
		}
		
		if (document.frmAddContato.cbActionContatos.value == 1) {
			window.location.href="frmAddContato.asp?Query=UPDATE&ID="+elemento.value
			
		} else if(document.frmAddContato.cbActionContatos.value == 2) {
			window.location.href="frmAddContato.asp?Query=DELETE&ID="+elemento.value
		}
		
	}
}

