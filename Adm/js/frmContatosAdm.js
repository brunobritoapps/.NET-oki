// JavaScript Document

// Retorna o objeto ajax
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


function windowLocationFind(url) {
	var find			 = true;
	var busca 			 = document.frmContatosAdm.txtFindContato.value;
	var por 			 = document.frmContatosAdm.typeFindContato.value;
	var status 			 = document.frmContatosAdm.cbStatusFindContato.value;
	var categoria		 = document.frmContatosAdm.cbCategoriasFindContato.value;
	var grupo 			 = document.frmContatosAdm.cbGruposCliente.value;
	//window.location.href = url + "frmContatosAdm.asp?find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria+"&grupo="+grupo;
	window.location.href = "frmContatosAdm.asp?find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria+"&grupo="+grupo;
}

//function checkUserContato() {
//	if ((document.frmContatosAdm.txtUsuarioContato.value != document.frmContatosAdm.hidden_usuario_info.value) && 
//		 (document.frmContatosAdm.txtSenhaContato.value != document.frmContatosAdm.hidden_senha_info.value)) {
//		var oAjaxUserContato = Ajax();
//	//	var berror = false;
//		oAjaxUserContato.onreadystatechange = function() {
//			if (oAjaxUserContato.readyState == 4 && oAjaxUserContato.status == 200) {
//				alert(oAjaxUserContato.responseText);
//				if (oAjaxUserContato.responseText == "true") {
//					alert("Usuário e Senha já cadastrado, favor cadastre outro usuário");
//					document.frmContatosAdm.txtUsuarioContato.focus();
//				}
//			}
//		}
//		
//		oAjaxUserContato.open("GET", "../../ajax/frmCadCliente.asp?sub=getcheckusercontato&id=0&user="+document.frmContatosAdm.txtUsuarioContato.value+"&senha="+document.frmContatosAdm.txtSenhaContato.value, true);
//		oAjaxUserContato.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
//		oAjaxUserContato.send(null);
//	}
//}
//
//function checkUsuario() {
//	if (document.frmContatosAdm.txtUsuarioContato.value != document.frmContatosAdm.hidden_usuario_info.value) {
//		var oAjaxUsuario = Ajax();
//	//	var berror = false;
//		oAjaxUsuario.onreadystatechange = function() {
//			if (oAjaxUsuario.readyState == 4 && oAjaxUsuario.status == 200) {
//				alert(oAjaxUsuario.responseText);
//				if (oAjaxUsuario.responseText == "true") {
//					alert("Usuário já cadastrado, favor cadastre outro usuário");
//					document.frmContatosAdm.txtUsuarioContato.focus();
//				}
//			}
//		}
//		
//		oAjaxUsuario.open("GET", "../../ajax/frmCadCliente.asp?sub=getcheckusuario&id=0&user="+document.frmContatosAdm.txtUsuarioContato.value, true);
//		oAjaxUsuario.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
//		oAjaxUsuario.send(null);
//	}
//}

//nova função devido a Tab1 e Tab2
function setAct(action)
{
    var form = document.frmEditCadastroClienteLc;
    document.getElementsByName("hiddenIDCliente").value = action;
	//form.hiddenTypeAction.value = action;

	if (document.getElementsByName("txtNomeContato")[0].value == "") {
		alert("Preencha o campo Nome");
		return;
	}
	if (document.getElementsByName("txtUsuarioContato")[0].value == "") {
		alert("Preencha o campo Usuário")
		return;
	}
	if (document.getElementsByName("txtSenhaContato")[0].value == "") {
		alert("Preencha o campo Senha");
		return;
	}

	if (!is_email(document.getElementsByName("txtEmailContato")[0].value))
	{
		alert("Email inválido!");
		return false;
	}
	if (document.getElementsByName("cbClienteContato")[0].value == -1) {
		alert("Escolha um Cliente");
		return;
	}
	form.submit(action);
}


// Seta a acao do ADMIN no envio do form
function setActionForm(action) {
	var form = document.frmContatosAdm;
    
	form.hiddenTypeAction.value = action;

	if (form.txtNomeContato.value == "") {
		alert("Preencha o campo Nome");
		return;
	}
	if (form.txtUsuarioContato.value == "") {
		alert("Preencha o campo Usuário")
		return;
	}
	if (form.txtSenhaContato.value == "") {
		alert("Preencha o campo Senha");
		return;
	}

	if (!is_email(form.txtEmailContato.value))
	{
		alert("Email inválido!");
		return false;
	}
	/*
	if (form.txtEmailContato.value == "") {
		alert("Preencha o campo Email");
		return;
	}
	if (form.txtEmailContato.value.indexOf('@') == -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('@.') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('.@') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('.') == -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('com') == -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('[') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf(']') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('(') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf(')') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('/') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('\\') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('..') != -1) {
		alert("Email inválido!");
		return;
	}
	if (form.txtEmailContato.value.indexOf('com.com') != -1) {
		alert("Email inválido!");
		return;
	}
	*/
	if (form.cbClienteContato.value == -1) {
		alert("Escolha um Cliente");
		return;
	}
	form.submit();	
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



