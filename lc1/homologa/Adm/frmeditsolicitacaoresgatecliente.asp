<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim NumSolColeta
	Dim QtdCartuchos
	Dim QtdCartuchosRecebidos
	Dim DataSolicitacao
	Dim DataAprovacao
	Dim DataProgramada
	Dim DataEnvioTransp
	Dim DataReceb
	Dim StatusSol
	Dim MotivoStatus
	Dim Master
	Dim RazaoSocial
	Dim NomeFantasia
	Dim CEP
	Dim LogradouroColeta
	Dim NumEndColeta
	Dim CompEndColeta
	Dim MunEndColeta
	Dim UFEndColeta
	Dim DDDEndColeta
	Dim TelEndColeta
	Dim ContatoColeta
	Dim ReqColetaDomiciliar
	Dim NumRecTransportadora
	Dim IDCliente
	Dim IDTransp
	Dim StatusAprovar
	Dim StatusAtualizar
	Dim DocBaixa
	Dim DataBaixa

	ReqColetaDomiciliar = Request.QueryString("iscoletadomiciliar")

	Sub GetSolicitacao()
		Dim sSql, arrSolicitacao, intSolicitacao, i
		sSql = "SELECT " & _
				"[Status_coleta_idStatus_coleta], " & _
				"[numero_solicitacao_coleta], " & _
				"[qtd_cartuchos], " & _
				"[qtd_cartuchos_recebidos], " & _
				"[data_solicitacao], " & _
				"[data_aprovacao], " & _
				"[data_programada], " & _
				"[data_envio_transportadora], " & _
				"[data_entrega_pontocoleta], " & _
				"[data_recebimento], " & _
				"[motivo_status], " & _
				"[isMaster] " & _
				"FROM [marketingoki2].[dbo].[Solicitacao_coleta] " & _
				"WHERE [idSolicitacao_coleta] = " & Request.QueryString("idsolic")

'		Response.Write sSql
'		Response.End()
		Call search(sSql, arrSolicitacao, intSolicitacao)
		dim arrSolicitacao2, intSolicitacao2
		sSql = "SELECT [idSolicitacoes_resgate] " & _
				  ",[cod_bonus] " & _
				  ",[idsolicitacao] " & _
				  ",[documento_baixa] " & _
				  ",[data_baixa] " & _
				  ",[data_solicitacao_resgate] " & _
				  ",[numero_solicitacao_geracao] " & _
				  ",[idproduto] " & _
				  ",[quantidade] " & _
				  ",[idcliente] " & _
			  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] WHERE [idsolicitacao] = " & Request.QueryString("idsolic")
'		response.write sSql & "<br />"
		Call search(sSql, arrSolicitacao2, intSolicitacao2)
		If intSolicitacao > -1 and intSolicitacao2 > -1 Then
			For i=0 To intSolicitacao
				StatusSol 							= arrSolicitacao(0,i)
				NumSolColeta 						= arrSolicitacao(1,i)
				QtdCartuchos 						= arrSolicitacao(2,i)
				QtdCartuchosRecebidos 				= arrSolicitacao(3,i)
				DataSolicitacao 					= arrSolicitacao(4,i)
				DataAprovacao 		 				= arrSolicitacao(5,i)
				DataProgramada 						= arrSolicitacao(6,i)
				DataEnvioTransp 					= arrSolicitacao(7,i)
				DataReceb 							= arrSolicitacao(9,i)
				MotivoStatus 						= arrSolicitacao(10,0)
				DocBaixa							= arrSolicitacao2(3,0)
				if DocBaixa = "NULL" then DocBaixa = ""
				DataBaixa							= arrSolicitacao2(4,i)
			Next

			If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
				If Not isNull(DataSolicitacao)				Then DataSolicitacao	= FormatDateTime(DataSolicitacao, 2)
				If Not isNull(DataAprovacao)				Then DataAprovacao		= FormatDateTime(DataAprovacao, 2)
				If Not isNull(DataProgramada)				Then DataProgramada		= FormatDateTime(DataProgramada, 2)
				If Not isNull(DataEnvioTransp)				Then DataEnvioTransp	= FormatDateTime(DataEnvioTransp, 2)
				If Not isNull(DataReceb)					Then DataReceb			= FormatDateTime(DataReceb, 2)
				If Not isNull(DataBaixa)					Then DataBaixa			= FormatDateTime(DataBaixa, 2)
			Else
				'If Not isNull(DataSolicitacao)				Then DataSolicitacao	= DateRight(FormatDateTime(DataSolicitacao, 2))
				'If Not isNull(DataAprovacao)				Then DataAprovacao		= DateRight(FormatDateTime(DataAprovacao, 2))
				If Not isNull(DataSolicitacao)				Then DataSolicitacao	= FormatDateTime(DataSolicitacao, 2)
				If Not isNull(DataAprovacao)				Then DataAprovacao		= FormatDateTime(DataAprovacao, 2)
				If Not isNull(DataProgramada)				Then DataProgramada		= DateRight(FormatDateTime(DataProgramada, 2))
				If Not isNull(DataEnvioTransp)				Then DataEnvioTransp	= DateRight(FormatDateTime(DataEnvioTransp, 2))
				If Not isNull(DataReceb)					Then DataReceb			= DateRight(FormatDateTime(DataReceb, 2))
				If Not isNull(DataBaixa)					Then DataBaixa			= DateRight(FormatDateTime(DataBaixa, 2))
			End If

				'response.write StatusSol & "<br />"
				'response.write NumSolColeta & "<br />"
				'response.write QtdCartuchos & "<br />"
				'response.write QtdCartuchosRecebidos & "<br />"
				'response.write DataSolicitacao & "<br />"
				'response.write DataAprovacao & "<br />"
				'response.write DataProgramada & "<br />"
				'response.write DataEnvioTransp & "<br />"
				'response.write DataReceb & "<br />"
				'response.write MotivoStatus & "<br />"
				'response.write DocBaixa & "<br />"
				'response.write DataBaixa & "<br />"

		End If
		Call GetCliente()
	End Sub

	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano

		Dia = Left(sData, 2)
		Dia = Replace(Dia, "/", "")
		If Len(Dia) = 1 Then
			Dia = "0" & Dia
		End If
		If Len(Replace(Left(sData, 2), "/", "")) = 1 Then
			Mes = Mid(sData, 3, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		Else
			Mes = Mid(sData, 4, 2)
			Mes = Replace(Mes, "/", "")
			If Len(Mes) = 1 Then
				Mes = "0" & Mes
			End If
		End If
		Ano = Right(sData, 4)
		Ano = Replace(Ano, "/", "")
		If Len(Ano) = 1 Then
			Ano = "0" & Ano
		End If
		DateRight = Mes & "/" & Dia & "/" & Ano
	End Function

	Sub GetCliente()
		Dim sSql, arrCliente, intCliente, i
		sSql = "SELECT " & _
						"B.[razao_social], " & _
						"B.[nome_fantasia], " & _
						"B.[cnpj], " & _
						"B.[compl_endereco_coleta], " & _
						"B.[numero_endereco_coleta], " & _
						"B.[contato_respcoleta], " & _
						"B.[ddd_respcoleta], " & _
						"B.[telefone_respcoleta], " & _
						"B.[idClientes] " & _
						"FROM [marketingoki2].[dbo].[Solicitacoes_Resgate_Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
						"ON A.[idcliente] = B.[idClientes] " & _
						"WHERE A.[idsolicitacao] = " & Request.QueryString("idsolic")

		Call search(sSql, arrCliente, intCliente)
		If intCliente > -1 Then
			For i=0 To inCliente
				RazaoSocial   = arrCliente(0,i)
				NomeFantasia  = arrCliente(1,i)
				NumEndColeta  = arrCliente(4,i)
				CompEndColeta = arrCliente(3,i)
				IDCliente	  = arrCliente(8,i)
				DDDEndColeta  = arrCliente(6,i)
				TelEndColeta  = arrCliente(7,i)
				ContatoColeta = arrCliente(5,i)
			Next
		End If
	End Sub

	function getMoeda(id)
		dim sql, arr, intarr, i
		sql = "SELECT [moeda] FROM [marketingoki2].[dbo].[Bonus_Gerado_Clientes] WHERE [Clientes_idClientes] = " & id
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getMoeda = arr(0,i)
			next
		else
			getMoeda = ""
		end if
	end function

	Sub GetStatusColeta()
		Dim sSql, arrStatus, intStatus, i
		Dim sSelected
		sSelected = ""

		sSql = "SELECT " & _
						"[idStatus_coleta], " & _
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta]"
		Call search(sSql, arrStatus, intStatus)
		If intStatus > -1 Then
			For i=0 To intStatus
				If StatusSol = arrStatus(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value='"&arrStatus(0,i)&"' "&sSelected&">"&arrStatus(1,i)&"</option>"
			Next
		End If
	End Sub

	Sub GetDescStatusColeta(lId)
		Dim sSql, arrStatus, intStatus, i
		
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _ 
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta] " & _
						"WHERE idStatus_coleta = " & lId
		Call search(sSql, arrStatus, intStatus)						
		If intStatus > -1 Then
			Response.Write arrStatus(1,0)
		End If
	End Sub

	Sub RequestForm()
		QtdCartuchosRecebidos 		= Request.Form("txtQtdCatuchosRecebidos")
		DataAprovacao 				= Request.Form("txtDataAprovacao")
		DataProgramada 				= Request.Form("txtDataProgramada")
		DataReceb 					= Request.Form("txtDataRecebimento")
		StatusSol 					= Request.Form("cbStatusSolColeta")
		MotivoStatus 				= Request.Form("txtMotivoStatus")
		DataEnvioTransp				= Request.Form("txtDataEnvioTransportadora")
		NumRecTransportadora  		= Request.Form("txtNumConhTransportadora")
		IDTransp 					= Request.Form("cbTransp")
		DocBaixa					= request.form("txtDocBaixa")
		DataBaixa					= request.form("txtDataBaixa")


		If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
			If QtdCartuchosRecebidos 	= "" Then QtdCartuchosRecebidos				= "NULL"
			If DataReceb 				= "" Then DataReceb 						= "NULL" Else DataReceb = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If
			If DataAprovacao 			= "" Then DataAprovacao 					= "NULL" Else DataAprovacao = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If
			If DataEnvioTransp 			= "" Then DataEnvioTransp 					= "NULL" Else DataEnvioTransp = "CONVERT(DATETIME, '"&FormatDate(DataEnvioTransp)&"')" End If
			If NumRecTransportadora  	= "" Then NumRecTransportadora				= "NULL"
			If DataEntPontoColeta 	 	= "" Then DataEntPontoColeta 				= "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
			If IDTransp 				= "" Then IDTransp 							= "NULL"
			If MotivoStatus 			= "" Then MotivoStatus 						= "NULL"
			If DocBaixa 				= "" Then DocBaixa 							= "NULL"
			If DataBaixa 				= "" Then DataBaixa 						= "NULL" Else DataBaixa = "CONVERT(DATETIME, '"&FormatDate(DataBaixa)&"')" End If
		else
			If QtdCartuchosRecebidos 	= "" Then QtdCartuchosRecebidos				= "NULL"
			If DataReceb 				= "" Then DataReceb 						= "NULL" Else DataReceb = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If
			If DataAprovacao 			= "" Then DataAprovacao 					= "NULL" Else DataAprovacao = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If
			If DataEnvioTransp 			= "" Then DataEnvioTransp 					= "NULL" Else DataEnvioTransp = "CONVERT(DATETIME, '"&FormatDate(DataEnvioTransp)&"')" End If
			If NumRecTransportadora  	= "" Then NumRecTransportadora				= "NULL"
			If DataEntPontoColeta 	 	= "" Then DataEntPontoColeta 				= "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
			If IDTransp 				= "" Then IDTransp 							= "NULL"
			If MotivoStatus 			= "" Then MotivoStatus 						= "NULL"
			If DocBaixa 				= "" Then DocBaixa 							= "NULL"
			If DataBaixa 				= "" Then DataBaixa 						= "NULL" Else DataBaixa = "CONVERT(DATETIME, '"&FormatDate(DataBaixa)&"')" End If
		end if

'		response.write DocBaixa
'		response.write DataBaixa
'		response.end

'		response.write DataReceb & "<br />"
'		response.write DataAprovacao & "<br />"
'		response.write DataEnvioTransp & "<br />"
'		response.write DataEntPontoColeta & "<br />"
'		response.write DataBaixa & "<br />"
'		response.end
	End Sub

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			Call UpdateSol()
		Else
			Call GetSolicitacao()
		End If
	End Sub

	Sub UpdateSol()
		Dim sSql, arrSol, intSol, i

		select case request.Form("btnAprovar")
			case "Aprovar"
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"SET [Status_coleta_idStatus_coleta]	= 2 " & _
						  ",[data_aprovacao]					= CONVERT(DATETIME, '"&month(now())&"/"&day(now())&"/"&year(now())&"') " & _
						  ",[data_programada]					= NULL " & _
						  ",[data_envio_transportadora]			= NULL " & _
						  ",[data_entrega_pontocoleta]			= NULL " & _
						  ",[data_recebimento]					= NULL " & _
						"WHERE [idSolicitacao_coleta]			= " & Request.Form("id")
		'			response.write sSql & "<br />"
		'			response.end
				Call exec(sSql)
		'		response.write DataBaixa
		'		response.end
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
						   "SET [documento_baixa] = '"&DocBaixa&"' " & _
						   ",[data_baixa] = "&DataBaixa&" " & _
						   "WHERE [idsolicitacao] = " & Request.Form("id")
		'		response.write sSql & "<br />"
		'		response.end
				Call exec(sSql)
			case "Salvar"
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"SET [Status_coleta_idStatus_coleta]	= 6 " & _
						  ",[data_aprovacao]					= CONVERT(DATETIME, '"&month(now())&"/"&day(now())&"/"&year(now())&"') " & _
						  ",[data_programada]					= NULL " & _
						  ",[data_envio_transportadora]			= NULL " & _
						  ",[data_entrega_pontocoleta]			= NULL " & _
						  ",[data_recebimento]					= NULL " & _
						"WHERE [idSolicitacao_coleta]			= " & Request.Form("id")
		'			response.write sSql & "<br />"
		'			response.end
				Call exec(sSql)
		'		response.write DataBaixa
		'		response.end
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
						   "SET [documento_baixa] = '"&DocBaixa&"' " & _
						   ",[data_baixa] = "&DataBaixa&" " & _
						   "WHERE [idsolicitacao] = " & Request.Form("id")
		'		response.write sSql & "<br />"
		'		response.end
				Call exec(sSql)
			case "Rejeitar"
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"SET [Status_coleta_idStatus_coleta]	= 3 " & _
						  ",[data_aprovacao]					= NULL" & _
						  ",[data_programada]					= NULL " & _
						  ",[data_envio_transportadora]			= NULL " & _
						  ",[data_entrega_pontocoleta]			= NULL " & _
						  ",[data_recebimento]					= NULL " & _
						"WHERE [idSolicitacao_coleta]			= " & Request.Form("id")
				call exec(sSql)
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
						   "SET [documento_baixa] = NULL " & _
						   ",[data_baixa] = NULL " & _
						   "WHERE [idsolicitacao] = " & Request.Form("id")
				Call exec(sSql)
                Call DevSaldo()
			case "Cancelar"
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"SET [Status_coleta_idStatus_coleta]	= 4 " & _
						  ",[data_aprovacao]					= NULL" & _
						  ",[data_programada]					= NULL " & _
						  ",[data_envio_transportadora]			= NULL " & _
						  ",[data_entrega_pontocoleta]			= NULL " & _
						  ",[data_recebimento]					= NULL " & _
						"WHERE [idSolicitacao_coleta]			= " & Request.Form("id")
				call exec(sSql)
				sSql = "UPDATE [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] " & _
						   "SET [documento_baixa] = NULL " & _
						   ",[data_baixa] = NULL " & _
						   "WHERE [idsolicitacao] = " & Request.Form("id")
				Call exec(sSql)
                Call DevSaldo()
		end select
		response.write "<script>window.opener.location.reload();</script>"
	    response.Write "<script>window.parent.close();</script>"
	End Sub

    '04/12/2014 - Thiago de Menezes (Loop Consultoria) - o código abaixo foi recuperado de um backup antigo, após a solicitação
    'do André Aggio e da Elaine Nascimento. Essa função é chamada no evento dos botões Cancelar e Rejeitar
    '#25/08/2014
	'devolve o saldo de bônus que foi rejeitado.
	Sub DevSaldo()
		'
		'localiza a solicitação de resgate (inicial com R)
		Dim sSql, arrSR, intSR, i, NumBon
		
		sSql = "select numero_solicitacao_coleta from Solicitacao_coleta where idSolicitacao_coleta = " & Request.Form("id") & ""
		call search(sSql, arrSR, intSR)
	
		If intSR > -1 Then
			NumBon = arrSR(0,i)
		End If
		'
		'
		'recalcula o saldos e salva saldo recalculado.
		dim sql, arr, intarr
		dim nVlrPontuacao
		dim nBonus

		sql = " select numero_solicitacao from Solicitacao_Resgate_has_Solicitacao_Composicao "
		sql = sql & " where numero_resgate = '" & NumBon & "' "
		sql = sql & " and numero_solicitacao <> '' "
		
		call search(sql, arr, intarr)

		If intarr > -1 Then
			For i = 0 To intarr
				dim iJ
				'
				'
				'Faz a devolução do saldo para a solicitação de coleta.
				'Localiza o que foi baixado e retorna o saldo.
				sql = " select pontuacao, pontuacao_atingir, saldo, numero_solicitacao from bonus_gerado_clientes where numero_solicitacao  = '" & arr(0,i) & "' "
				
				call search(sql, arrJ, intArrJ)

				If intArrJ > -1 Then
				
					For iJ = 0 To intArrJ
						nVlrPontuacao   = nVlrPontuacao + arrJ(0,iJ)
						nBonus          = nBonus + arrJ(0,iJ)
					Next

				End If

				'
				'depois de ler todo o bonus gerado. salva novamente o bonus.
				sql = "update bonus_gerado_clientes set saldo = " & nVlrPontuacao & " where numero_solicitacao = '"& arr(0,i) & "'"
				call exec(sql)
				'response.write "<td><tr>" & sql & "</tr></td>"
				
				'
				'zera novamente as variáveis.
				nVlrPontuacao = 0

			Next

		End If		

		'response.write "<script>alert('Resgate rejeitado com sucesso! Foi retornado um Bônus de: '"& nBonus &"');</script>"
		
	End Sub

	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia
		Dia = Left(sDate, 2)
		Mes = Mid(sDate, 4, 2)
		Mes = Replace(Mes, "/" ,"")
		If Len(Mes) = 1 Then
			Mes = "0" & Mes
		End If
		Ano = Right(sDate, 4)

		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function

	Call SubmitForm()

	If CInt(StatusSol) = 3 Or CInt(StatusSol) = 1 Then
		StatusAprovar = True
	Else
		StatusAprovar = False
	End if

	If Not CInt(StatusSol) = 1 And Not CInt(StatusSol) = 3 and not cint(StatusSol) = 6 Then
		StatusAtualizar = True
	Else
		StatusAtualizar = False
	End If



%>
<html><head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<style>
	label {
		font-weight:bold;
	}
</style>
<script language="javascript" src="js/frmEditSolicitacaoColetaDomiciliarAdm.js"></script>
<SCRIPT LANGUAGE="JavaScript" SRC="js/CalendarPopup.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var cal = new CalendarPopup();
</SCRIPT>
<script>
	function validaFormResgate() 
	{
		var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
		<%If StatusSol <> 1 Then%>
			if (form.txtDocBaixa.value == "") {
				alert("Preencha o campo Documento Baixa");
				return false;
			}
			
			if (form.txtDataBaixa.value == "") {
				alert("Preencha o campo Data Baixa");
				return false;
			}else{				
				if (!validateGetDate(form.txtDataAprovacao.value, form.txtDataBaixa.value)) 
				{
					alert("Preencha o campo Data de Baixa não pode\n ser menor que o campo Data aprovação!");
					return false;
				}	
			}
		<%end if%>
		//return true;
	}
	

	function rejeitarResgate() {
		var form = document.frmEditSolicitacaoColetaDomiciliarAdm;
		form.cbStatusSolColeta.value = 2;
		form.submit();
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
</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmeditsolicitacaoresgatecliente.asp" name="frmEditSolicitacaoColetaDomiciliarAdm" method="POST">
		<input type="hidden" name="hiddenReqColetaDomiciliar" value="<%=ReqColetaDomiciliar%>" />
		<input type="hidden" name="id" value="<%=Request.QueryString("idsolic")%>" />
		<input type="hidden" name="hiddenIdCliente" value="<%=IDCliente%>" />
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Administrar Solicitação de Resgate</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de coleta: </label></td>
				<td><%=NumSolColeta%>&nbsp;<img src="img/buscar.gif" class="imgexpandeinfo" align="absmiddle" alt="Buscar Solicitações que compuseram a solicitação Master" onClick="javascript:window.open('frmviewsolicitacaocompoeresgate.asp?idsolic=<%=NumSolColeta%>','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"/></td>
			</tr>
			<tr id="tridcliente">
				<td width="35%" align="right"><label id="idcliente">ID. Cliente: </label></td>
				<td><%=IDCliente%></td>
			</tr>
			<tr id="trrazaosocial">
				<td width="35%" align="right"><label id="razaosocial">Razão Social: </label></td>
				<td><%=RazaoSocial%></td>
			</tr>
			<tr id="trnomefantasia">
				<td width="35%" align="right"><label id="nomefantasia">Nome Fantasia: </label></td>
				<td><%=NomeFantasia%></td>
			</tr>
			<tr id="trqtdcartuchos">
				<td width="35%" align="right"><label id="qtdcartuchos"><%if getMoeda(IDCliente) = "P" then%>Qtd. cartuchos:<%Else%>Valor:<%End If%></label></td>
				<td><%=QtdCartuchos%>&nbsp;<img src="img/produtos.gif" align="absmiddle" class="imgexpandeinfo" width="25" height="22" name="listprodutos" alt="Produtos" onClick="javascript:window.open('frmlistaprodutosresgate.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /></td>
			</tr>
			<tr id="trdatasolicitacao">
				<td width="35%" align="right"><label id="datasolicitacao">Data solicitação: </label></td>
				<td><%=DataSolicitacao%></td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="status">Status: </label></td>
				<td>
					<!--
					<select name="cbStatusSolColeta" class="select">
						<option value="-1"> --- Selecione --- </option>
						<%Call GetStatusColeta()%>
					</select>
					-->
					<%call GetDescStatusColeta(StatusSol)%>
					<INPUT type="hidden" id="cbStatusSolColeta" name="cbStatusSolColeta" value="<%=StatusSol%>">
				</td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="motivostatus">Motivo status: </label></td>
				<td><textarea name="txtMotivoStatus" style="width:250px;height:100px;"><%If Not MotivoStatus = "NULL" Then Response.Write MotivoStatus End If%></textarea></td>
			</tr>
			<%If StatusSol <> 1 Then%>
			<tr id="trdataaprovacao">
				<td width="35%" align="right"><label id="dataaprovacao">Data aprovação: </label></td>
				<td>
					<input type="text" <%If isNull(DataAprovacao) Then%>class="textreadonly" value="<%=DataAprovacao%>"<%ELse%>class="text" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" readonly="readonly" size="13" maxlength="10" onKeyPress="date(this)" />
					<%If StatusSol = 2 or StatusSol = 6 or StatusSol = 4 or StatusSol = 3 Then%>
					<%else%>
					<A HREF="#" onClick="cal.select(document.forms['frmEditSolicitacaoColetaDomiciliarAdm'].txtDataAprovacao,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img align="absmiddle" src="img/btn_calendario.gif" border="0"></A> 
					<%end if%>
				</td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Documento Baixa: </label></td>
				<td><input type="text" <%If isNull(DocBaixa) Or DocBaixa = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDocBaixa" value="<%=DocBaixa%>" size="13" maxlength="10" /></td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Data Baixa: </label></td>
				<td><input type="text" <%If isNull(DataBaixa) Or DataBaixa = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataBaixa" value="<%=DataBaixa%>" size="13" maxlength="10" onKeyPress="date(this)" readonly />
 					<%If StatusSol = 6 or StatusSol = 4 or StatusSol = 3 Then%>
					<%else%>
				  <A HREF="#" onClick="cal.select(document.forms['frmEditSolicitacaoColetaDomiciliarAdm'].txtDataBaixa,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img align="absmiddle" src="img/btn_calendario.gif" border="0"></A> 
				  <%end if%>
				</td>
			</tr>
			<%end if%>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<% If StatusSol = "1" or StatusSol = "2" Then %>
				<td align="right">
					<input type="submit" class="btnform" name="btnAprovar" value="<%if StatusSol = "2" Then Response.Write "Salvar" else Response.Write "Aprovar" end if%>" onClick="return validaFormResgate()" />
					<input type="submit" class="btnform" name="btnAprovar" value="<%if StatusSol = "2" Then Response.Write "Cancelar" else Response.Write "Rejeitar" end if%>" onClick="rejeitarResgate()" />
				</td>
			<% End If %>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
