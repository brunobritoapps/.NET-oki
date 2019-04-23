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
	document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value = 2;
	if (document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value == 2) {
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
				
				oAjax.open('GET','ajax/Ajax.asp?sub=aprovarsolicitacao&id='+ID,true);
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
	var change = document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value; 
	document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value = 3;
	if (confirm("Deseja realmente rejeitar essa Solicitação?")) {
		document.frmEditSolicitacaoColetaPontoColetaAdm.submit();	
	} else {
		document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value = change;
		return;		
	}
}

function cancelar(ID) {
	document.frmEditSolicitacaoColetaPontoColetaAdm.cbStatusSolColeta.value = 4;
	if (confirm("Deseja realmente cancelar essa Solicitação?")) {
		document.frmEditSolicitacaoColetaPontoColetaAdm.submit();	
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
	var form = document.frmEditSolicitacaoColetaPontoColetaAdm;
	
	if (form.cbStatusSolColeta.value == 5) {
		if (form.hiddenReqColetaDomiciliar.value == 1) {
			if (form.txtDataEnvioTransportadora.value == "") {
				alert("O status escolhido necessita que o campo Data Envio para Transportadora esteja preenchido!");
				return false;
			} 
		}
	}
	return true;
}

function validateInTransit() {
	var form = document.frmEditSolicitacaoColetaPontoColetaAdm;
	
	if (form.cbStatusSolColeta.value == 7) {
		if (form.txtDataEntregaPontoColeta.value == "") {
			alert("O status escolhido necessita que o campo Data Entrega no Ponto de Coleta esteja preenchido!");
			return false;
		}	
	}
	return true;
}

function validateFinish() {
	var form = document.frmEditSolicitacaoColetaPontoColetaAdm;
	
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
	if (!validateStandByColect()) {
		error++;		
	}
	if (!validateInTransit()) {
		error++;
	} else {
		if (!validaDataEnvPontoColeta()) {
			error++;	
		}	
	}
	if (!validateFinish()) {
		error++;	
	}
	if (error == 0) {
		document.frmEditSolicitacaoColetaPontoColetaAdm.submit();		
	}
}

function getData(value) {
	var _date = new Date();
	var _arrData = value.split("/");

	var _dia = _arrData[0];
	var _mes = _arrData[1];
	var _ano = _arrData[2];
	
	//alert(_dia + "/" + _mes + "/" + _ano);
	//alert(_dia);
	//alert(_mes);
	
	_date.setFullYear(_ano, _mes, _dia);
	//_date = _dia + "/" + _mes + "/" + _ano

	return _date;
}

function validaDataEnvPontoColeta() {
	var form = document.frmEditSolicitacaoColetaPontoColetaAdm;
	if (form.txtDataAprovacao.value == "") {
		alert("Data de Aprovação tem que estar preenchido!");	
		return false;
	} else {
		if (form.txtDataEntregaPontoColeta.value == "") {
			alert("Preencha o campo Data Entrega no Ponto Coleta corretamente!");			
			return false;
		} else {
			if (!validateGetDate(form.txtDataAprovacao.value, form.txtDataEntregaPontoColeta.value))  {
				alert("Preencha o campo Data Entrega no Ponto Coleta corretamente!");
				return false;	
			} else {
				form.cbStatusSolColeta.value = 8;
				return true;	
			}
		}
	}
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