// JavaScript Document

function validate() {
	form = document.frmTransportadorasAdm;
	if (form.txtRazaoSocial.value == "") {
		alert("Preencha o campo Razão Social!");	
		return;
	}
	if (form.txtNomeFantasia.value == "") {
		alert("Preencha o campo Nome Fantasia!");
		return;
	}
	if (form.txtCNPJ.value == "") {
		alert("Preencha o campo CNPJ!");
		return;
	}
	if (form.txtContato.value == "") {
		alert("Preencha o campo Contato!");
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
	if (form.txtFax.value == "") {
		alert("Preencha o campo Fax!");
		return;
	}
	if (form.radioColetaEmail[1].checked) {
		document.getElementById("obrig_email").innerHTML = "*";	

  		if (!is_email(form.txtEmail.value))
		{
		    alert("Email inválido!");
		    return false;
		}
		/*
		if (form.txtEmail.value == "") {
			alert("Preencha o campo Email");
			return;
		}
		if (form.txtEmail.value.indexOf('@') == -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('@.') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('.@') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('.') == -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('com') == -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('[') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf(']') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('(') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf(')') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('/') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('\\') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('..') != -1) {
			alert("Email inválido!");
			return;
		}
		if (form.txtEmail.value.indexOf('com.com') != -1) {
			alert("Email inválido!");
			return;
		}
		*/
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

function cnpj_format(cnpj) {
	var form = document.frmTransportadorasAdm;
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

function cnpjExists() {
	var oAjax = Ajax();
	
	if (document.frmTransportadorasAdm.txtCNPJ.value != "" && (document.frmTransportadorasAdm.verifycnpj.value != document.frmTransportadorasAdm.txtCNPJ.value)) {
		oAjax.onreadystatechange = function() {
			if (oAjax.readyState == 4 && oAjax.status == 200) {
				if (oAjax.responseText == "true") {
					alert("CNPJ já cadastrado, favor insira outro CNPJ!");
					document.frmTransportadorasAdm.txtCNPJ.focus();
				} 		
			}	
		}
		
		oAjax.open("GET", "ajax/Ajax.asp?sub=cnpjexists2&id="+document.frmTransportadorasAdm.txtCNPJ.value, true);
		oAjax.setRequestHeader("Content-Type","application/x-www-form-urlencoded; charset=iso-8859-1");
		oAjax.send(null);
	}
}

function checkObrigatorioEmail() {
	var form = document.frmTransportadorasAdm;
	if (form.radioColetaEmail[1].checked) {
		document.getElementById("obrig_email").innerHTML = " *";	
	} else {
		document.getElementById("obrig_email").innerHTML = "";	
	}
}
