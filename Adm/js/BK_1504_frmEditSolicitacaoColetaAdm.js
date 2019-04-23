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

function date(campo) {
		var string = campo.value;
		var _char = "/";
		for (var i=0;i<string.length;i++) {
			if (i == 2 || i == 5) {
				continue;	
			}
		}
		switch (string.length) {
			case 2:
				string += _char;
				break;
			case 5:
				string += _char;
				break;						
		}
		campo.value = string;
}

function aprovar(ID) {
	_data();
	if (document.frmEditSolicitacaoColetaAdm.cbStatusSolColeta.value == 1) {
		if (confirm("Deseja realmente aprovar essa solicitação?")) {
			try {
				var oAjax = Ajax();
				oAjax.onreadystatechange = function() {
					if (oAjax.readyState == 4 && oAjax.status == 200) {
						document.getElementById("msgret").innerHTML = oAjax.responseText;
						document.getElementById("btnprove").style.display = 'none';
						document.getElementById("btnatualizar").style.display = 'block';
					}	
				}
				
				oAjax.open("GET","ajax/Ajax.asp?sub=aprovarsolicitacao&id="+ID,true);
				oAjax.setRequestHeader("Content-Type","application/x-form-www-urlencoded; charset=iso-8859-1");
				oAjax.send(null);
			} catch (exception) {
				alert(exception);	
			}
			window.opener.location.reload();
			window.close();
		}
	}
}

function reprovar(ID) {
	var change = document.frmEditSolicitacaoColetaAdm.cbStatusSolColeta.value; 
	document.frmEditSolicitacaoColetaAdm.cbStatusSolColeta.value = 3;
	if (confirm("Deseja realmente rejeitar essa Solicitação?")) {
		document.frmEditSolicitacaoColetaAdm.submit();	
	} else {
		document.frmEditSolicitacaoColetaAdm.cbStatusSolColeta.value = change;
		return;		
	}
}

function _data() {
	var data = new Date();
	var	_dia = data.getDate();
	var	_mes = data.getMonth();
	var	_ano = data.getFullYear();
	if (_dia < 10) {
		_dia = "0" + _dia;	
	} 
}

function validateStandByColect() {
	var form = document.frmEditSolicitacaoColetaAdm;
	
	if (form.cbStatusSolColeta.value == 5) {
		if (form.hiddenReqColetaDomiciliar.value == 1) {
			if (parseInt(form.hiddenIsColetaEmail.value) == 1 && form.txtDataEnvioTransportadora.value == "") {
				form.txtDataEnvioTransportadora.value =  _data();		
			} else {
				if (form.txtDataEnvioTransportadora.value == "") {
					alert("O status escolhido necessita que o campo Data Envio para Transportadora esteja preenchido!");
					return false;
				}
			}
			if (form.txtDataProgramada.value == "") {
				alert("O status escolhido necessita que o campo Data Programada esteja preenchido!");
				return false;
			}
			if (form.txtNumConhTransportadora.value == "") {
				alert("O status escolhido necessita que o campo Número de conhecimento da Transportadora esteja preenchido!");	
				return false;
			}
			if (form.cbTransp.value == -1) {
				alert("O status escolhido necessita que o campo Número de conhecimento da Transportadora esteja preenchido!");	
				return false;
			}
		}
	}
	return true;
}


function validateInTransit() {
	var form = document.frmEditSolicitacaoColetaAdm;
	
	if (form.txtNumConhTransportadora.value != "") {
		form.cbStatusSolColeta.value = 7;
		if (form.txtDataProgramada.value == "") {
			alert("O status escolhido necessita que o campo Data Programada esteja preenchido!");
			return false;
		}
	}
	if (form.cbStatusSolColeta.value == 7) {
		if (form.hiddenReqColetaDomiciliar.value == 1) {
			if (form.txtNumConhTransportadora.value == "") {
				alert("O status escolhido necessita que o campo Número de conhecimento da Transportadora esteja preenchido!");
				return false;
			}
		} 
	}
	return true;
}

function validateFinish() {
	var form = document.frmEditSolicitacaoColetaAdm;
	
	if (form.txtDataRecebimento.value != "") {
		form.cbStatusSolColeta.value = 6;
	}
	if (form.cbStatusSolColeta.value == 6) {
		if (form.txtDataRecebimento.value == "") {
			alert("O status escolhido necessita que o campo Data recebimento pelo Operador Logístico esteja preenchido!");
			return false;
		}
	}
	return true;
}

function validateForm() {
	var error = 0;
	changeStatusKeyPress();
	if (!validateStandByColect()) {
		error++;		
	} else {
		if (!validaDataProgramada()) {
			error++;	
		}
		if (!validaDataEnvioTransp()) {
			error++;	
		}
	}
	if (!validateInTransit()) {
		error++;
	}
	if (!validateFinish()) {
		error++;	
	}
	if (error == 0) {
		document.frmEditSolicitacaoColetaAdm.submit();		
	}
}

function changeStatusKeyPress() {
	var form = document.frmEditSolicitacaoColetaAdm;
	if (form.txtDataRecebimento.value == "" && form.txtNumConhTransportadora.value == "") {
		if (form.txtDataEnvioTransportadora.value != "" || form.txtDataProgramada.value != "") {
			form.cbStatusSolColeta.value = 5;
		} else {
			form.cbStatusSolColeta.value = 2;
		}	
	}
}

function validaDataProgramada() {
	var form = document.frmEditSolicitacaoColetaAdm;
	if (form.cbStatusSolColeta.value == 5 || form.cbStatusSolColeta.value == 7) {
		if (form.txtDataEnvioTransportadora.value == "") {
			alert("Data Envio para Transportadora tem que estar preenchido!");
			return false;
		} else {
			if (form.txtDataProgramada.value == "") {
				alert("Preencha o campo Data Programada para coleta!");
				return false;
			}	 else {
	//			var testedata = getData(form.txtDataProgramada.value);
	//			alert(testedata.getDate() + "/" + testedata.getMonth() + "/" + testedata.getFullYear());
				if (!validateGetDate(getData(form.txtDataEnvioTransportadora.value), getData(form.txtDataProgramada.value))) {
					alert("Preencha o campo Data Programada para coleta corretamente!");
					return false;
				} else {
					form.cbStatusSolColeta.value = 5;
					return true;	
				}
			}
		}
	} else {
		return true;	
	}
}

function validaDataEnvioTransp() {
	var form = document.frmEditSolicitacaoColetaAdm;	
	if (form.cbStatusSolColeta.value == 5 || form.cbStatusSolColeta.value == 7) {
		if (form.txtDataAprovacao.value == "") {
			alert("Data de Aprovação tem que estar preenchido!");	
			return false;
		} else {
			if (form.txtDataEnvioTransportadora.value == "") {
				alert("Preencha o campo Data Envio Transportadora!");
				return false;
			} else {
				if (!validateGetDate(getData(form.txtDataAprovacao.value), getData(form.txtDataEnvioTransportadora.value)))  {
					alert("Preencha o campo Data de Envio para Transportadora corretamente!");
					return false;	
				} else {
					return true;	
				}
			}
		}
	} else {
		return true;	
	}
}

function getData(value) {
	var _date = new Date();
	var _arrData = value.split("/");

	var _dia = _arrData[0];
	var _mes = parseInt(_arrData[1]) - 1;
	var _ano = _arrData[2];
	
//	alert(_dia + "/" + _mes + "/" + _ano);
	
	_date.setFullYear(_ano, _mes, _dia);

	return _date;
}

function validateGetDate(dataDefault, date) {
//	alert(dataDefault.getDate() + "/" + dataDefault.getMonth() + "/" + dataDefault.getFullYear());
//	alert(date.getDate() + "/" + date.getMonth() + "/" + date.getFullYear());
	if (parseInt(date.getFullYear()) < parseInt(dataDefault.getFullYear())) {
		return false;		
	} else {
		if (parseInt(date.getMonth()) < parseInt(dataDefault.getMonth())) {
			return false;	
		} else {
			switch(parseInt(date.getMonth() + 1)) {
				case 1: // Janeiro
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false
					}
					break;
				case 2: // Fevereiro
					if (parseInt(date.getDate()) > 28 || parseInt(date.getDate()) < 1) {
						return false;		
					}
					break;
				case 3: // Março
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 4: // Abril
					if (parseInt(date.getDate()) > 30 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 5: // Maio
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 6: // Junho
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 7: // Julho
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 8: // Agosto
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 9: // Setembro
					if (parseInt(date.getDate()) > 30 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 10: // Outubro
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 11: // Novembro
					if (parseInt(date.getDate()) > 30 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				case 12: // Dezembro
					if (parseInt(date.getDate()) > 31 || parseInt(date.getDate()) < 1) {
						return false;	
					}
					break;
				default:
					return false;
					break;
			}
			if (parseInt(date.getDate()) < parseInt(dataDefault.getDate())) {
				return false;	
			}
			return true;
		}
	}
}



