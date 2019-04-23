<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
	Dim NumSolColeta
	Dim QtdCartuchos
	Dim QtdCartuchosRecebidos
	Dim DataSolicitacao
	Dim DataAprovacao
	Dim DataProgramada
	Dim DataEnvioTransp
	Dim DataEntPontoColeta
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
	Dim IDPontoColeta
	Dim IDTransp
	Dim StatusAprovar
	Dim StatusAtualizar

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
'		response.write sSql
'		response.end				
'		Response.Write Session("IDAdministrator")
'		Response.End()						
		Call search(sSql, arrSolicitacao, intSolicitacao)
		If intSolicitacao > -1 Then
			For i=0 To intSolicitacao
				StatusSol 							= arrSolicitacao(0,i)
				NumSolColeta 						= arrSolicitacao(1,i)
				QtdCartuchos 						= arrSolicitacao(2,i)
				QtdCartuchosRecebidos 	= arrSolicitacao(3,i)
				DataSolicitacao 				= arrSolicitacao(4,i)
				DataAprovacao 		 			= arrSolicitacao(5,i)
				DataProgramada 					= arrSolicitacao(6,i)		
				DataEnvioTransp 				= arrSolicitacao(7,i)
				DataEntPontoColeta 			= arrSolicitacao(8,i)
				DataReceb 							= arrSolicitacao(9,i)
				MotivoStatus 						= arrSolicitacao(10,i)
				Master									= arrSolicitacao(11,i)
			Next
			

		If Not isNull(DataSolicitacao) Then DataSolicitacao = DateRight(FormatDateTime(DataSolicitacao, 2))						
		If Not isNull(DataAprovacao) Then DataAprovacao = DateRight(FormatDateTime(DataAprovacao, 2)) 						
		If Not isNull(DataProgramada) Then DataProgramada = DateRight(FormatDateTime(DataProgramada, 2))						
		If Not isNull(DataEnvioTransp) Then DataEnvioTransp = DateRight(FormatDateTime(DataEnvioTransp, 2))						
		If Not isNull(DataEntPontoColeta) Then DataEntPontoColeta = DateRight(FormatDateTime(DataEntPontoColeta, 2))
		If Not isNull(DataReceb) Then DataReceb = DateRight(FormatDateTime(DataReceb, 2))

		End If
		Call GetCliente()
'		If ReqColetaDomiciliar = 1 Then
'			Call GetEnderecoColeta()
'			Call GetIDTransp()		
'		Else
			Call GetPontoColeta()
'		End If	
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
	
	Sub GetPontoColeta()
		Dim sSql, arrPontoCol, intPontoCol, i

'		sSql = "SELECT " & _ 
'				"A.idPontos_coleta, " & _
'				"A.cep_consulta_idcep_consulta, " & _
'				"A.razao_social, " & _
'				"A.cnpj, " & _
'				"A.numero_endereco, " & _
'				"A.complemento_endereco, " & _
'				"B.cep, " & _
'				"B.logradouro, " & _
'				"B.bairro, " & _
'				"B.municipio, " & _
'				"B.estado, " & _
'				"B.tipologradouro, " & _
'				"D.nome " & _
'				"FROM Pontos_coleta AS A " & _
'				"LEFT JOIN cep_consulta AS B " & _
'				"ON A.cep_consulta_idcep_consulta = B.idcep_consulta " & _
'				"LEFT JOIN Solicitacao_coleta_has_Clientes AS C " & _
'				"ON A.idPontos_coleta = C.Pontos_coleta_idPontos_coleta " & _
'				"LEFT JOIN Contatos AS D " & _
'				"ON C.Contatos_idContatos = D.idContatos " & _
'				"WHERE C.Solicitacao_coleta_idSolicitacao_coleta = "&Request.QueryString("idsolic")&" AND " & _
'				"C.Clientes_idClientes = " & IDCliente
				
		sSql =  "SELECT " & _
				"A.idPontos_coleta, " & _ 
				"A.razao_social, " & _ 
				"A.cnpj, " & _ 
				"A.numero_endereco, " & _ 
				"A.complemento_endereco, " & _ 
				"A.cep, " & _ 
				"A.logradouro, " & _ 
				"A.bairro, " & _ 
				"A.municipio, " & _ 
				"A.estado, " & _ 
				"D.nome " & _ 
				"FROM Pontos_coleta AS A " & _
				"LEFT JOIN Solicitacao_coleta_has_Clientes AS C " & _ 
				"ON A.idPontos_coleta = C.Pontos_coleta_idPontos_coleta " & _ 
				"LEFT JOIN Contatos AS D " & _ 
				"ON C.Contatos_idContatos = D.idContatos " & _ 
				"WHERE C.Solicitacao_coleta_idSolicitacao_coleta = "&Request.QueryString("idsolic")&" " & _ 
				"AND C.Clientes_idClientes = " & IDCliente
				
				'id ponto de coleta 	= 0
				'razao social 			= 1
				'cnpj 					= 2
				'numero endereco		= 3
				'complemento endereco 	= 4
				'cep					= 5
				'logradouro				= 6
				'bairro					= 7
				'municipio				= 8
				'estado					= 9
				'nome contato			= 10
				
'		Response.Write sSql
'		Response.End()		
		Call search(sSql, arrPontoCol, intPontoCol)
		
		If intPontoCol > -1 Then
			For i=0 To intPontoCol
				IDPontoColeta 	 = arrPontoCol(0,i)
				CEP			  			 = arrPontoCol(5,i)
				LogradouroColeta = trim(arrPontoCol(6,i)) & " - " & trim(arrPontoCol(7,i))
				MunEndColeta 		 = arrPontoCol(8,i)
				UFEndColeta 		 = arrPontoCol(9,i)
				NumEndColeta  	 = arrPontoCol(3,i)
				CompEndColeta 	 = arrPontoCol(4,i)
				ContatoColeta 	 = arrPontoCol(10,i)
			Next
		End If						
	End Sub
	
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
						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
						"ON A.[Clientes_idClientes] = B.[idClientes] " & _
						"WHERE A.[Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
'		response.write sSql
'		response.end						
						
		Call search(sSql, arrCliente, intCliente)
		If intCliente > -1 Then
			For i=0 To inCliente
				RazaoSocial   = arrCliente(0,i)
				NomeFantasia  = arrCliente(1,i)
				NumEndColeta  = arrCliente(4,i)
				CompEndColeta = arrCliente(3,i)
				IDCliente			= arrCliente(8,i) 
				DDDEndColeta  = arrCliente(6,i)
				TelEndColeta	= arrCliente(7,i)
				ContatoColeta = arrCliente(5,i) 	
			Next
		End If
	End Sub

	Sub GetEnderecoColeta()
		Dim sSql, arrEnd, intEnd, i
		sSql = "SELECT " & _
						"A.[cep_consulta_idcep_consulta], " & _
						"B.[cep], " & _
						"B.[logradouro], " & _
						"B.[bairro], " & _
						"B.[municipio], " & _
						"B.[estado], " & _
						"B.[tipologradouro] " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta] AS B " & _
						"ON A.[cep_consulta_idcep_consulta] = B.[idcep_consulta] " & _
						"WHERE A.[isEnderecoColeta] = 1 AND A.[Clientes_idClientes] = " & IDCliente
		Call search(sSql, arrEnd, intEnd)
		If intEnd > -1 Then
			For i=0	To intEnd
				CEP 						 = arrEnd(1,i)
				LogradouroColeta = Trim(arrEnd(6,i)) & ". " & arrEnd(2,i) & " - " & arrEnd(3,i)
				MunEndColeta 		 = arrEnd(4,i)
				UFEndColeta 		 = arrEnd(5,i)
			Next					
		End If
	End Sub
	
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
	
	Sub RequestForm()
		QtdCartuchosRecebidos = Request.Form("txtQtdCatuchosRecebidos")
		'DataAprovacao 				= Request.Form("txtDataAprovacao")
		'DataProgramada 				= Request.Form("txtDataProgramada")
		'DataReceb 						= Request.Form("txtDataRecebimento")
		'StatusSol 						= Request.Form("cbStatusSolColeta")
		'MotivoStatus 					= Request.Form("txtMotivoStatus")
		'If Request.Form("hiddenReqColetaDomiciliar") = 0 Then
		'	DataEntPontoColeta 		= Request.Form("txtDataEntregaPontoColeta")
		'	DataEnvioTransp				= Request.Form("txtDataEnvioTransportadora")
		'	NumRecTransportadora  = Request.Form("txtNumConhTransportadora")
		'	IDTransp 							= Request.Form("cbTransp")
		'Else
		'	DataEntPontoColeta 	 = Request.Form("txtDataEntregaPontoColeta")
		'	DataEnvioTransp 		 = Request.Form("txtDataEnvioTransportadora")
		'	NumRecTransportadora = Request.Form("txtNumConhTransportadora")
		'	IDTransp 						 = Request.Form("cbTransp")
		'End If
		If QtdCartuchosRecebidos = "" Then QtdCartuchosRecebidos = "NULL" 
		'If DataReceb 						 = "" Then DataReceb = "NULL" Else DataReceb = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If 
		'If DataProgramada 			 = "" Then DataProgramada = "NULL" Else DataProgramada = "CONVERT(DATETIME, '"&FormatDate(DataProgramada)&"')" End If 
		'If DataAprovacao 				 = "" Then DataAprovacao = "NULL" Else DataAprovacao = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If 
		'If DataEnvioTransp 			 = "" Then DataEnvioTransp = "NULL" Else DataEnvioTransp = "CONVERT(DATETIME, '"&FormatDate(DataEnvioTransp)&"')" End If
		'If NumRecTransportadora  = "" Then NumRecTransportadora = "NULL" 
		'If DataEntPontoColeta 	 = "" Then DataEntPontoColeta = "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
		'If IDTransp 						 = "" Then IDTransp = "NULL" 
		'If MotivoStatus 				 = "" Then MotivoStatus = "NULL"
		
'		Response.Write DataProgramada
'		Response.End()
	End Sub
	
	Sub GetTransp()
		Dim sSql, arrTransp, intTransp, i
		Dim sSelected
		sSql = "SELECT [idTransportadoras] " & _
			   ",[nome_fantasia] " & _
			   "FROM [marketingoki2].[dbo].[Transportadoras] " & _
			   "WHERE [status] = 1"
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If GetIDTransp() = arrTransp(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If	
				Response.Write "<option value="&arrTransp(0,i)&" "&sSelected&">"&arrTransp(1,i)&"</option>"
			Next
		End If	   
	End Sub
	
	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			Call UpdateSol()			
			'verificar se a qtd max foi atingida
			Call VerifyQtdMax()
		Else
			Call GetSolicitacao()
		End If
	End Sub
	
	Sub UpdateSol()
		Dim sSql, arrSol, intSol, i

		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
					 "SET [qtd_cartuchos_recebidos] = "&QtdCartuchosRecebidos&", " & _
					 "[data_entrega_pontocoleta] = '"& year(now()) & "-" & month(now()) & _
					 "-" & day(now()) &"', " & _
					 "[status_coleta_idstatus_coleta] = 8 " & _
					 "WHERE [idSolicitacao_coleta] = " & Request.Form("id")
'		Response.Write sSql
'		Response.End()		
		Call exec(sSql)
		call insereSolicitacaoBaixada(Request.Form("id"), Session("IDPonto"))

		'sSql = "INSERT INTO Solicitacao_coleta_has_Clientes (Solicitacao_coleta_idSolicitacao_coleta, typeColect, Pontos_coleta_idPontos_coleta, Contatos_idContatos, Clientes_idClientes, cep_coleta, logradouro_coleta, bairro_coleta, numero_endereco_coleta, comp_endereco_coleta, municipio_coleta, estado_coleta, ddd_resp_coleta, telefone_resp_coleta, contato_coleta) VALUES (" & Request.Form("id") & ", 0, )"
		'Call exec(sSql)
	End Sub
	
	sub insereSolicitacaoBaixada(id_sol, id_ponto)
		dim sql
		sql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_Baixadas] " & _
					   "([id_solicitacao] " & _
					   ",[id_pontocoleta] " & _
					   ",[numero_solicitacao_master] " & _
					   ",is_baixada) " & _
				 "VALUES " & _
					   "("&id_sol&" " & _
					   ","&id_ponto&" " & _
					   ",NULL " & _
					   ",0)"
		call exec(sql)					   
	end sub

	Sub VerifyQtdMax()
		Dim sql, arrVlrTot, intVlrTot
		dim soma
		dim i
		sql = "select " & _
				"b.qtd_cartuchos_recebidos, " & _
				"a.id_solicitacao " & _
				"from solicitacoes_baixadas as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.id_solicitacao = b.idsolicitacao_coleta " & _
				"where a.id_pontocoleta = "&Session("IDPonto")&" and a.is_baixada = 0"

'		response.write sql
'		response.end

		Call search(sql, arrVlrTot, intVlrTot)
'		response.write getMaxCartuchos() & " " & arrVlrTot(0,0)
'		response.end
		if intVlrTot > -1 then
			for i=0 to intVlrTot
				soma = soma + arrVlrTot(0,i)
			next
			if soma >= getMaxCartuchos() then
				dim num
				dim cont_valormax
				dim sql_update
				'gerar numero solicitacao master
				num = AddSolColetaMaster(Session("IDPonto"), soma)
'				response.write num
'				response.end
				for cont_valormax=0 to intVlrTot
					sql_update = "UPDATE [marketingoki2].[dbo].[Solicitacoes_Baixadas] " & _
								 "SET [numero_solicitacao_master] = '"&num&"', is_baixada = 1 " & _
								 "WHERE [id_solicitacao] = "&arrVlrTot(1,cont_valormax)&" and [id_pontocoleta] = " & Session("IDPonto")
'					response.write sql_update & "<br />"
					call exec(sql_update)			 
				next
'				response.end
			end if
		end if	
		Response.Write "<script>"
		Response.Write "window.opener.location.reload();"
		Response.Write "window.close()"
		Response.Write "</script>"			
	End Sub
	
	function getMaxCartuchos()
		Dim sql, arrVlrMax, intVlrMax

		sql = "SELECT Qtd_Limite_Cartuchos FROM [marketingoki2].[dbo].[Pontos_coleta] " & _
		      "WHERE idPontos_coleta = " & Session("IDPonto")
		Call search(sql, arrVlrMax, intVlrMax)
		if intVlrMax > -1 then
			getMaxCartuchos = arrVlrMax(0,0)				
		else
			getMaxCartuchos = 0
		end if
	end function
	
	function AddSolColetaMaster(ByVal IDPonto, ByVal QtdCartuchos)
		Dim oCommand
		Dim NumeroSequencial
		Dim NumeroSolicitacaoColeta
		Dim IdentifiedCharacterColeta
		Dim DateMonthColeta
		Dim DateYearColeta
		
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandType = 4
		oCommand.CommandText = "sp_addSolicitacaoMaster"

		IdentifiedCharacterColeta = "M"

		DateMonthColeta = Month(Now())
		If Len(DateMonthColeta) = 1 Then
			DateMonthColeta = "0" & DateMonthColeta
		End If
		DateYearColeta = Right(Year(Now()), 2)

		NumeroSolicitacaoColeta = IdentifiedCharacterColeta & DateYearColeta & DateMonthColeta
		NumeroSequencial = getSequencial(False)
		NumeroSolicitacaoColeta = NumeroSolicitacaoColeta & NumeroSequencial
		NumeroSolicitacaoColeta = getDigitoControle(NumeroSolicitacaoColeta)

		oCommand.Parameters("@idpontocoleta") = IDPonto
		oCommand.Parameters("@numerosolicitacaocoleta") = NumeroSolicitacaoColeta
		oCommand.Parameters("@qtdcartuchos") = QtdCartuchos
		
		oCommand.Execute()
		AddSolColetaMaster = NumeroSolicitacaoColeta
	End function	
	
	Function GetIDTransp()
		Dim sSql, arrId, intId, i
		Dim Ret
		sSql = "SELECT [Transportadoras_idTransportadoras] " & _
				 ",[numero_reconhecimento_transportadora] " & _
			  	 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
				 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(0,i)
			Next
		End If
		GetIDTransp = Ret				 
	End Function
	
	Function GetNumRecTransportadora()
		Dim sSql, arrId, intId, i
		Dim Ret
		sSql = "SELECT [Transportadoras_idTransportadoras] " & _
				 ",[numero_reconhecimento_transportadora] " & _
			  	 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
				 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(1,i)
			Next
		End If
		GetNumRecTransportadora = Ret				 
	End Function
	
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
	
	If Not CInt(StatusSol) = 1 And Not CInt(StatusSol) = 3 Then
		StatusAtualizar = True
	Else
		StatusAtualizar = False
	End If
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<style>
	label {
		font-weight:bold;
	}
</style>
<script language="javascript">
function validateForm() {
	var error = 0;

	if (document.frmEditSolicitacaoEntrega.txtQtdCatuchosRecebidos.value == "") {
		alert("Preencha o campo Quantidade de Cartuchos");
		error++;
	}

	if (error == 0) {
		document.frmEditSolicitacaoEntrega.submit();		
	}
}

function soNumeros(pFld, e) {
	var sep = 0; 
	var key = ''; 
	var i = j = 0; 
	var len = len2 = 0; 
	var strCheck = '0123456789'; 
	var aux = aux2 = ''; 
	var whichCode = (window.Event) ? e.which : e.keyCode;
	if (whichCode == 13) return true;
	if (whichCode == 0) return true;
	if (whichCode == 8) return true;
	key = String.fromCharCode(whichCode);// Valor para o código da Chave 
	if (strCheck.indexOf(key) == -1) return false; // Chave inválida 
}
</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="" name="frmEditSolicitacaoEntrega" method="POST">
		<input type="hidden" name="hiddenReqColetaDomiciliar" value="<%=ReqColetaDomiciliar%>" />
		<input type="hidden" name="id" value="<%=Request.QueryString("idsolic")%>" />
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Administrar Solicitação de Entrega</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de coleta: </label></td>
				<td><%=NumSolColeta%></td>
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
			<tr id="trcontatopontocoleta" <%If ReqColetaDomiciliar = 1 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="contatopontocoleta">Contato Coleta: </label></td>
				<td><%=ContatoColeta%></td>
			</tr>
			<!--
			<tr id="trnumpontocoleta" <%If ReqColetaDomiciliar = 1 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="numpontocoleta">Num. Ponto de Coleta: </label></td>
				<td><%=IDPontoColeta%></td>
			</tr>
			<tr id="trcepcoleta">
				<td width="35%" align="right"><label id="cepcoleta">CEP: </label></td>
				<td><%=CEP%></td>
			</tr>
			<tr id="trlogcoleta">
				<td width="35%" align="right"><label id="logcoleta">Logradouro Coleta: </label></td>
				<td><%=LogradouroColeta%></td>
			</tr>
			<tr id="trnumendcoleta">
				<td width="35%" align="right"><label id="numendcoleta">Num. end. Coleta: </label></td>
				<td><%=NumEndColeta%></td>
			</tr>
			<tr id="trcompendcoleta">
				<td width="35%" align="right"><label id="compendcoleta">Comp. end. Coleta: </label></td>
				<td><%=CompEndColeta%></td>
			</tr>
			<tr id="trmunendcoleta">
				<td width="35%" align="right"><label id="munendcoleta">Mun. end. Coleta: </label></td>
				<td><%=MunEndColeta%></td>
			</tr>
			<tr id="trufendcoleta">
				<td width="35%" align="right"><label id="ufendcoleta">UF. end. Coleta: </label></td>
				<td><%=UFEndColeta%></td>
			</tr>
			<tr id="trdddendcoleta" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="dddendcoleta">DDD. end. Coleta: </label></td>
				<td><%=DDDEndColeta%></td>
			</tr>
			<tr id="trtelendcoleta" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="telendcoleta">Tel. end. Coleta: </label></td>
				<td><%=TelEndColeta%></td>
			</tr>
			-->
			<tr id="trcontatocoleta" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="contatocoleta">Contato Coleta: </label></td>
				<td><%=ContatoColeta%></td>
			</tr>
			<tr id="trqtdcartuchos">
				<td width="35%" align="right"><label id="qtdcartuchos">Qtd. cartuchos: </label></td>
				<td><%=QtdCartuchos%></td>
			</tr>
			<tr id="trdatasolicitacao">
				<td width="35%" align="right"><label id="datasolicitacao">Data solicitação: </label></td>
				<td><%=DataSolicitacao%></td>
			</tr>
			<tr id="trqtdcartuchosrecebidos">
				<td width="35%" align="right"><label id="qtdcartuchosrecebidos">Qtd. cartuchos recebidos: </label></td>
				<td>
				<%if len(trim(QtdCartuchosRecebidos)) = 0 or isnull(QtdCartuchosRecebidos) then%>
					<input type="text" <%If isNull(QtdCartuchosRecebidos) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtQtdCatuchosRecebidos" value="<%=QtdCartuchosRecebidos%>" size="4" onKeyPress="return(soNumeros(this,event));" />
				<%else%>
					<%=QtdCartuchosRecebidos%>
				<%end if%>
				<!--
				&nbsp;<img src="img/produtos.gif" align="absmiddle" class="imgexpandeinfo" width="25" height="22" name="listprodutos" alt="Produtos" onClick="javascript:window.open('frmListProdSol.asp?idsol=<%=Request.QueryString("idsolic")%>&qtdrec='+document.frmEditSolicitacaoColetaAdm.txtQtdCatuchosRecebidos.value+'','','width=350,height=575,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" />
				-->
				</td>
			</tr>
			<!--
			<tr>
				<td width="35%" align="right"><label id="status">Status: </label></td>
				<td>
					<select name="cbStatusSolColeta" class="select">
						<option value="-1"> --- Selecione --- </option>	
						<%Call GetStatusColeta()%>
					</select>
				</td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="motivostatus">Motivo status: </label></td>
				<td><textarea name="txtMotivoStatus" style="width:250px;height:100px;"><%If Not MotivoStatus = "NULL" Then Response.Write MotivoStatus End If%></textarea></td>
			</tr>
			<tr id="trdataaprovacao">
				<td width="35%" align="right"><label id="dataaprovacao">Data aprovação: </label></td>
				<td><input type="text" <%If isNull(DataAprovacao) Then%>class="textreadonly" value="<%=DataAprovacao%>"<%ELse%>class="text" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="trdataprogramada" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="dataprogramada">Data programada: </label></td>
				<td><input type="text" <%If isNull(DataProgramada) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataProgramada" value="<%=DataProgramada%>" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="tridtransportadora" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="numidtransportadora">Transportadora: </label></td>
				<td>
					<select name="cbTransp" class="select">
						<option value="-1">Selecione uma Transportadora</option>
						<%Call GetTransp()%>
					</select>
					<img src="img/transportadoras.gif" class="imgexpandeinfo" width="25" height="25" align="absmiddle" alt="Buscar Transportadora" onClick="window.open('frmSearchTranspSol.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
				</td>
			</tr>
			<tr id="trdataenviotransportadora" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="dataenviotransportadora">Data envio transportadora: </label></td>
				<td><input type="text" <%If isNull(DataEnvioTransp) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataEnvioTransportadora" value="<%=DataEnvioTransp%>" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="trnumconhtransportadora" <%If ReqColetaDomiciliar = 0 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="numconhtransportadora">Número conhecimento da Transportadora: </label></td>
				<td><input type="text" <%If isNull(GetNumRecTransportadora()) Or GetNumRecTransportadora() = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtNumConhTransportadora" value="<%=GetNumRecTransportadora()%>" size="15" /></td>
			</tr>
			<tr id="trdataentregapontocoleta" <%If ReqColetaDomiciliar = 1 Then%>style="display:none;"<%End If%>>
				<td width="35%" align="right"><label id="dataentregapontocoleta">Data entrega Ponto Coleta: </label></td>
				<td><input type="text" <%If isNull(DataEntPontoColeta) Or DataEntPontoColeta = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataEntregaPontoColeta" value="<%=DataEntPontoColeta%>" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Data recebimento: </label></td>
				<td><input type="text" <%If isNull(DataReceb) Or DataReceb = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataRecebimento" value="<%=DataReceb%>" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="master">Master: </label></td>
				<td><%If Master = 1 Then%>Sim<%Else%>Não<%End If%></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
			-->
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<%if len(trim(QtdCartuchosRecebidos)) = 0 or isnull(QtdCartuchosRecebidos) then%>
			<tr id="btnatualizar" <% If Not StatusAtualizar Then %>style="display:none;"<%End If%>>
				<td align="center" colspan="2"><input type="button" class="btnform" name="btnAtualizar" value="Atualizar" onClick="validateForm()" /></td>
			</tr>
			<%end if%>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
