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

function sleep(milliseconds) {
	  var start = new Date().getTime();
	  for (var i = 0; i < 1e7; i++) {
		if ((new Date().getTime() - start) > milliseconds){
		  break;
		}
	  }
  }

function aprovar(ID, TIPO) {

    
		if (confirm("Deseja realmente aprovar essa solicitação?")) {
			try {

			    var oAjax = Ajax();
				oAjax.onreadystatechange = function () {
				}

				oAjax.open("GET", "ajax/ajax.asp?sub=aprovarsolicitacao&id=" + ID + "&Tipo=" + TIPO, true);
				oAjax.setRequestHeader("Content-Type","application/x-form-www-urlencoded; charset=iso-8859-1");
				oAjax.send(null);

			} catch (exception) {
				alert(exception);	
			}
		    //window.opener.location.reload();
			alert('Solicitação aprovada. Será enviado um e-mail para os responsáveis envolvidos.');
		    //document.frmEditSolicitacaoColetaDomiciliarAdm.submit();
			//sleep(500);
			window.close();
			window.opener.location.reload();
			
			
		}
    }

function reprovar(ID) {
	var change = document.frmEditSolicitacaoColetaDomiciliarAdm.cbStatusSolColeta.value; 
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	
	document.frmEditSolicitacaoColetaDomiciliarAdm.cbStatusSolColeta.value = 3;
	form.cbStatusSolColeta.value = 3;
	
	if (confirm("Deseja realmente rejeitar essa Solicitação?")) {
		document.frmEditSolicitacaoColetaDomiciliarAdm.submit();	
	} else {
		document.frmEditSolicitacaoColetaDomiciliarAdm.cbStatusSolColeta.value = change;
		return;		
	}
}

function cancelar(ID) {
	document.frmEditSolicitacaoColetaDomiciliarAdm.cbStatusSolColeta.value = 4;
	if (confirm("Deseja realmente cancelar essa Solicitação?")) {
		document.frmEditSolicitacaoColetaDomiciliarAdm.submit();	
	}
}

function _data() {
	var data = new Date();
	var	_dia = data.getDate();
	var	_mes = parseInt(data.getMonth()) + 1;
	var	_ano = data.getFullYear();
	if (_dia < 10) {
		_dia = "0" + _dia;	
	} 
	if (_mes < 10) {
		_mes = "0" + _mes;	
	}
	return _dia + "/" + _mes + "/" + _ano;
}

function validateStandByColect() {
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	
	if (form.cbStatusSolColeta.value == 5) {
		if (form.hiddenReqColetaDomiciliar.value == 1) {
			if (parseInt(form.hiddenIsColetaEmail.value) == 1 && form.txtDataEnvioTransportadora.value == "") {
				form.txtDataEnvioTransportadora.value =  _data();		
			}
		}
	}
	return true;
}

function validateInTransit() {
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	
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
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
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
		document.frmEditSolicitacaoColetaDomiciliarAdm.submit();		
	}
}

function changeStatusKeyPress() {
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	if (form.txtDataRecebimento.value == "" && form.txtNumConhTransportadora.value == "") {
		if (form.txtDataEnvioTransportadora.value != "" || form.txtDataProgramada.value != "") {
			form.cbStatusSolColeta.value = 5;
		} else {
			form.cbStatusSolColeta.value = 2;
		}	
	}
}

function validaDataProgramada() {
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	if (form.cbStatusSolColeta.value == 7 || form.txtDataEnvioTransportadora.value != "") {
		if (form.txtDataEnvioTransportadora.value == "") {
			alert("Data Envio para Transportadora tem que estar preenchido!");
			return false;
		} else {
			if (form.txtDataProgramada.value == "") {
				alert("Preencha o campo Data Programada para coleta!");
				return false;
			}	 else {
				//var testedata = getData(form.txtDataProgramada.value);
				//alert(testedata.getDate() + "/" + testedata.getMonth() + "/" + testedata.getFullYear());
				if (!validateGetDate(form.txtDataEnvioTransportadora.value, form.txtDataProgramada.value)) {
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
	var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
	if (form.cbStatusSolColeta.value == 5 || form.cbStatusSolColeta.value == 7) {
		if (form.txtDataAprovacao.value == "") {
			alert("Data de Aprovação tem que estar preenchido!");	
			return false;
		} else {
			if (form.txtDataEnvioTransportadora.value == "") {
				alert("Preencha o campo Data Envio Transportadora!");
				return false;
			} else {
				if (!validateGetDate(form.txtDataAprovacao.value, form.txtDataEnvioTransportadora.value))  {
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
	//alert(value+' - data entrada');
	var _date = new Date();
	var _arrData = value.split("/");

	var _dia = _arrData[0];
	var _mes = _arrData[1] - 1;
	var _ano = _arrData[2];
	
	//alert(_dia + "/" + _mes + "/" + _ano);
	
	_date.setFullYear(_ano, _mes, _dia);

	//alert(_date+' - data atual');

	return _date;
}

function validateGetDate(dataDefault, date) 
{
	//alert(date);
	//alert(dataDefault);

	var arrData1 = dataDefault.split("/");
	var arrData2 = date.split("/");

	var dia1 = arrData1[0];
	var mes1 = arrData1[1];
	var ano1 = arrData1[2];
	
	var dia2 = arrData2[0];
	var mes2 = arrData2[1];
	var ano2 = arrData2[2];
	
	var data1 = ano1+mes1+dia1
	var data2 = ano2+mes2+dia2

	//alert(data2);
	//alert(data1);

	if (parseInt(data2) >= parseInt(data1))
	{
	  //alert("maior ou igual");
	  return true;
	}
	else
	{
	  //alert("menor");
	  return false;
	}
}

function validateGetDate1(dataDefault, date) {
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

// JavaScript Document