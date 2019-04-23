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
	
	Dim IdTranspHidden
	
	ReqColetaDomiciliar = 1

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
				DataReceb 							= arrSolicitacao(9,i)
				MotivoStatus 						= arrSolicitacao(10,i)
			Next
			
		If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
			If Not isNull(DataSolicitacao) Then DataSolicitacao = FormatDateTime(DataSolicitacao, 2) 						
			If Not isNull(DataAprovacao) Then DataAprovacao     = FormatDateTime(DataAprovacao, 2) 						
			If Not isNull(DataProgramada) Then DataProgramada   = FormatDateTime(DataProgramada, 2)						
			If Not isNull(DataEnvioTransp) Then DataEnvioTransp = FormatDateTime(DataEnvioTransp, 2)						
			If Not isNull(DataReceb) Then DataReceb             = FormatDateTime(DataReceb, 2)
		Else			
			If Not isNull(DataSolicitacao) Then DataSolicitacao = DateRight(FormatDateTime(DataSolicitacao, 2)) 						
			If Not isNull(DataAprovacao) Then DataAprovacao     = DateRight(FormatDateTime(DataAprovacao, 2)) 						
			If Not isNull(DataProgramada) Then DataProgramada   = DateRight(FormatDateTime(DataProgramada, 2))						
			If Not isNull(DataEnvioTransp) Then DataEnvioTransp = DateRight(FormatDateTime(DataEnvioTransp, 2))						
			If Not isNull(DataReceb) Then DataReceb             = DateRight(FormatDateTime(DataReceb, 2))
		End If	

		End If
		Call GetCliente()
		Call GetEnderecoColeta()
		Call GetIDTransp()		
	End Sub
	
	Function DateRight(sData)
		Dim Dia
		Dim Mes
		Dim Ano
		
		Dia = day(sData)
		if Dia < 10 then Dia = "0" & Dia

		Mes = month(sData)
		if Mes < 10 then Mes = "0" & Mes

		Ano = year(sData)

		DateRight = Dia & "/" & Mes & "/" & Ano
	End Function
	
	Sub GetCliente()
		Dim sSql, arrCliente, intCliente, i

'		sSql = "SELECT " & _ 
'						"B.[razao_social], " & _
'						"B.[nome_fantasia], " & _
'						"B.[cnpj], " & _ 
'						"B.[compl_endereco_coleta], " & _ 
'						"B.[numero_endereco_coleta], " & _
'						"B.[contato_respcoleta], " & _ 
'						"B.[ddd_respcoleta], " & _ 
'						"B.[telefone_respcoleta], " & _
'						"B.[idClientes] " & _
'						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
'						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
'						"ON A.[Clientes_idClientes] = B.[idClientes] " & _
'						"WHERE A.[Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")
						
		sSql = "select " & _
						"a.idpontos_coleta, " & _
						"a.razao_social, " & _ 
						"a.nome_fantasia, " & _
						"a.cnpj, " & _
						"a.complemento_endereco, " & _
						"a.numero_endereco, " & _
						"a.ddd, " & _
						"a.telefone " & _
						"from pontos_coleta as a " & _
						"left join solicitacao_coleta_has_pontos_coleta as b " & _
						"on a.idpontos_coleta = b.pontos_coleta_idpontos_coleta " & _
						"where b.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("idsolic")				
						
'		response.write sSql
'		response.end								
						
		Call search(sSql, arrCliente, intCliente)
		If intCliente > -1 Then
			For i=0 To inCliente
				RazaoSocial   = arrCliente(1,i)
				NomeFantasia  = arrCliente(2,i)
				NumEndColeta  = arrCliente(5,i)
				CompEndColeta = arrCliente(4,i)
				IDCliente	    = arrCliente(0,i) 
				DDDEndColeta  = arrCliente(6,i)
				TelEndColeta  = arrCliente(7,i)
			Next
		End If
	End Sub

	Sub GetEnderecoColeta()
		Dim sSql, arrEnd, intEnd, i

'		sSql = "SELECT " & _
'						"A.idendereco_coleta, " & _ 
'						"B.[cep], " & _ 
'						"B.[logradouro], " & _ 
'						"B.[bairro], " & _ 
'						"B.[municipio], " & _ 
'						"B.[estado], " & _ 
'						"B.[tipologradouro], " & _
'						"A.[numero_endereco_coleta], " & _ 
'						"A.[comp_endereco_coleta], " &_
'						"A.[ddd_resp_coleta], " &_
'						"A.[telefone_resp_coleta], " & _
'						"A.[contato_coleta] " & _
'						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS A " & _
'						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta] AS B " & _
'						"ON A.idendereco_coleta = B.idcep_consulta " & _ 
'						"WHERE Solicitacao_coleta_idSolicitacao_coleta = " & Request.QueryString("idsolic")						
						
		sSql = "select " & _
						"a.complemento_endereco, " & _
						"a.numero_endereco, " & _
						"a.ddd, " & _
						"a.telefone, " & _
						"a.cep, " & _
						"a.logradouro, " & _
						"a.bairro, " & _
						"a.municipio, " & _
						"a.estado " & _
						"from pontos_coleta as a " & _
						"left join solicitacao_coleta_has_pontos_coleta as b " & _
						"on a.idpontos_coleta = b.pontos_coleta_idpontos_coleta " & _
						"where b.solicitacao_coleta_idsolicitacao_coleta = " & Request.QueryString("idsolic")				
						
'		response.write sSql
'		response.end						
						
		Call search(sSql, arrEnd, intEnd)
		If intEnd > -1 Then
			For i=0	To intEnd
				CEP 						 = arrEnd(4,i)
				LogradouroColeta = arrEnd(5,i) & " - " & arrEnd(6,i)
				MunEndColeta 		 = arrEnd(7,i)
				UFEndColeta 		 = arrEnd(8,i)
				NumEndColeta  	 = arrEnd(1,i)
				CompEndColeta 	 = arrEnd(0,i)
				DDDEndColeta  	 = arrEnd(2,i)
				TelEndColeta		 = arrEnd(3,i)
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

		QtdCartuchosRecebidos 		            = Request.Form("txtQtdCatuchosRecebidos")
		DataAprovacao 						    = Request.Form("txtDataAprovacao")
		DataProgramada 						    = Request.Form("txtDataProgramada")
		DataReceb 								= Request.Form("txtDataRecebimento")
		StatusSol 								= Request.Form("cbStatusSolColeta")
		MotivoStatus 							= Request.Form("txtMotivoStatus")
		DataEnvioTransp						    = Request.Form("txtDataEnvioTransportadora")
		NumRecTransportadora  		            = Request.Form("txtNumConhTransportadora")
		IDTransp 								= Request.Form("cbTransp")

		If QtdCartuchosRecebidos 	            = "" Then QtdCartuchosRecebidos         = "NULL" 
		If DataReceb 							= "" Then DataReceb 					= "NULL" Else DataReceb 		  	 = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If 
		If DataProgramada 			 	        = "" Then DataProgramada 				= "NULL" Else DataProgramada  	 = "CONVERT(DATETIME, '"&FormatDate(DataProgramada)&"')" End If 
		If DataAprovacao 				 	    = "" Then DataAprovacao 				= "NULL" Else DataAprovacao   	 = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If 
		If DataEnvioTransp 			 	        = "" Then DataEnvioTransp 			    = "NULL" Else DataEnvioTransp 	 = "CONVERT(DATETIME, '"&FormatDate(DataEnvioTransp)&"')" End If
		If NumRecTransportadora  	            = "" Then NumRecTransportadora          = "NULL" 
		If DataEntPontoColeta 	 	            = "" Then DataEntPontoColeta 		    = "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
		If IDTransp 							= "" Then IDTransp 						= "NULL" 
		If MotivoStatus 				 	    = "" Then MotivoStatus 					= "NULL"
		
'		Response.Write DataProgramada
'		Response.End()
	End Sub
	
	Function GetIDTranspDefault()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT idtransp FROM pontos_coleta WHERE idpontos_coleta = " & IDCliente
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If Not isNull(arrTransp(0,i)) Then
					GetIDTranspDefault = arrTransp(0,i)
				End If
			Next
		End If
	End Function 
	
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
				ElseIf GetIDTranspDefault() = arrTransp(0,i) Then
					sSelected = "selected"
				Else	
					sSelected = ""
				End If	
				Response.Write "<option value="&arrTransp(0,i)&" "&sSelected&">"&arrTransp(1,i)&"</option>"
			Next
		End If	   
	End Sub
	
	Sub GetDescTransp()
		Dim sSql, arrTransp, intTransp, i
		Dim sSelected

		if len(trim(GetIDTransp())) then
			sSql = "SELECT [idTransportadoras] " & _
				   ",[nome_fantasia] " & _
				   "FROM [marketingoki2].[dbo].[Transportadoras] " & _
				   "WHERE [idTransportadoras] = " & GetIDTransp()

			Call search(sSql, arrTransp, intTransp)

			If intTransp > -1 Then
				Response.Write arrTransp(1,i)
				IdTranspHidden = arrTransp(0,i)
			End If	   
		end if
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
		
		sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
				"SET [Status_coleta_idStatus_coleta] = "&StatusSol&" " & _
				  ",[qtd_cartuchos_recebidos] = "&QtdCartuchosRecebidos&" " & _
				  ",[data_aprovacao] = "&DataAprovacao&" " & _
				  ",[data_programada] = "&DataProgramada&" " & _
				  ",[data_envio_transportadora] = "&DataEnvioTransp&" " & _
				  ",[data_entrega_pontocoleta] = NULL " & _
				  ",[data_recebimento] = "&DataReceb&" " & _
				  ",[motivo_status] = '"&MotivoStatus&"' " & _
				"WHERE [idSolicitacao_coleta] = " & Request.Form("id")
'		Response.Write sSql
'		Response.End()		
		Call exec(sSql)
		If CInt(Request.Form("hiddenReqColetaDomiciliar")) = 1 Then
			sSql = "SELECT [Solicitacao_coleta_idSolicitacao_coleta] " & _
						  ",[Transportadoras_idTransportadoras] " & _
						  ",[numero_reconhecimento_transportadora] " & _
					  "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
					  "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.Form("id")
			Call search(sSql, arrSol, intSol)		  
			If  intSol > -1 Then
				If IDTransp <> "" And IDTRansp <> "NULL" And IDTRansp <> "-1" Then
					sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _ 
							 "SET [Transportadoras_idTransportadoras] = "&IDTransp&", " & _
							 "[numero_reconhecimento_transportadora] = '"&NumRecTransportadora&"' " & _
							 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.Form("id")
'					Response.Write sSql
'					Response.End()		 
					Call exec(sSql)					 
				End If
			Else
				If IDTransp <> "" And IDTRansp <> "NULL" And IDTRansp <> "-1" Then
					sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
								   "([Solicitacao_coleta_idSolicitacao_coleta] " & _
								   ",[Transportadoras_idTransportadoras] " & _
								   ",[numero_reconhecimento_transportadora]) " & _
							 "VALUES " & _
								   "("&Request.Form("id")&" " & _
								   ","&IDTransp&" " & _
								   ",'"&NumRecTransportadora&"')"
	'				Response.Write sSql
	'				Response.End()		 
					Call exec(sSql)						   
				End If 
			End If
		End If	
		response.write "<script>window.opener.location.reload();</script>"
		Response.Write "<script>window.parent.close();</script>"							
	End Sub
	
	Function GetInfoEmailTransportadora()
		Dim Ret
		Dim sSqlCli, arrCli, intCli, i
		Dim sSqlSol, arrSol, intSol, j
		Dim NomeCliente
		Dim NumSolicitacao
		
		sSqlCli = "select * from pontos_coleta where idpontos_coleta = " & Request.Form("hiddenIdCliente")
		'Response.Write sSqlCli
		'Response.End
		
		Call search(sSqlCli, arrCli, intCli) 
		If intCli > -1 Then
			NomeCliente = arrCli(3,0)	
		End If
		
		sSqlSol = "select numero_solicitacao_coleta from solicitacao_coleta where idsolicitacao_coleta = " & Request.Form("id")
		Call search(sSqlSol, arrSol, intSol)
		If intSol > -1 Then
			NumSolicitacao = arrSol(0,0)	
		End If
		
		Ret = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">"
		Ret = Ret &	"<html xmlns=""http://www.w3.org/1999/xhtml"">"
		Ret = Ret &	"<head>"
		Ret = Ret &	"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" />"
		Ret = Ret &	"<title>Email para transportadora</title>"
		Ret = Ret &	"</head>"
		Ret = Ret &	"<body>"
		Ret = Ret &	"<div>"
		Ret = Ret &	"<table width=""100%"" border=""0"" cellspacing=""1"" cellpadding=""1"">"
		Ret = Ret &	"<tr>"
		Ret = Ret &	"<th scope=""row"">Cliente</th>"
		Ret = Ret &	"<td>"&NomeCliente&"</td>"
		Ret = Ret &	"</tr>"
		Ret = Ret &	"<tr>"
		Ret = Ret &	"<th scope=""row"">Número da Solicitação</th>"
		Ret = Ret &	"<td>"&NumSolicitacao&"</td>"
		Ret = Ret &	"</tr>"
		Ret = Ret &	"</table>"
		Ret = Ret &	"</div>"
		Ret = Ret &	"</body>"
		Ret = Ret &	"</html>"
		
		GetInfoEmailTransportadora = Ret		
		
	End Function
	
	Function GetIDTransp()
		Dim sSql, arrId, intId, i
		Dim Ret
		'sSql = "SELECT [Transportadoras_idTransportadoras] " & _
		'			 ",[numero_reconhecimento_transportadora] " & _
		'	  	 "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Transportadoras] " & _
		'			 "WHERE [Solicitacao_coleta_idSolicitacao_coleta] = " & Request.QueryString("idsolic")

		sSql = "SELECT idTransp FROM Pontos_coleta WHERE idPontos_Coleta = " & IDCliente

		'Response.Write ssql
		'Response.End
					 
		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(0,i)
			Next
		End If
		GetIDTransp = Ret				 
	End Function
	
	Function EmailTransportadora(ID)
		Dim sSql, arrTransp, intTransp, i
		
		sSql = "select email from transportadoras where idtransportadoras = " & ID
		
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			EmailTransportadora = arrTransp(0,0)
		Else	
			EmailTransportadora = ""
		End If
	End Function
	
	Function isColetaEmail()
		Dim sSql, arrTransp, intTransp, i
		Dim IDTransp
		Dim Ret
		
		If GetIDTransp() <> "" Then
			IDTransp = GetIDTransp()
		ElseIf GetIDTranspDefault() <> "" Then 	
			IDTransp = GetIDTranspDefault()
		End If
		
		If IDTransp <> "" Then
			sSql = "select iscoletaemail from transportadoras where idtransportadoras = " & IDTransp
			
			Call search(sSql, arrTransp, intTransp)
			If intTransp > -1 Then
				For i = 0 To intTransp
					Ret = arrTransp(0,i)	
				Next
			Else
				Ret = 0		
			End If
		Else
			Ret = 0
		End If	
		
		isColetaEmail = Ret
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
				if arrId(1,i) = "NULL" then
					Ret = ""
				else
					Ret = arrId(1,i)
				end if
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
		Ano = Right(sDate, 4)

		'Dia = day(sDate)
		'Mes = month(sDate)
		'Ano = year(sDate)
		
		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function
	
	Call SubmitForm()
	
	If CInt(StatusSol) = 3 Or CInt(StatusSol) = 1 Then
		StatusAprovar = True 
	Else
		StatusAprovar = False
	End if	
	
	If Not CInt(StatusSol) = 1 And Not CInt(StatusSol) = 3 and not cint(StatusSol) = 6 and not cint(StatusSol) = 4 Then
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
<script language="javascript" src="js/frmEditSolicitacaoColetaAdm.js"></script>
<SCRIPT LANGUAGE="JavaScript" SRC="js/CalendarPopup.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
var cal = new CalendarPopup();
</SCRIPT>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmEditSolicitacaoColetaAdm.asp" name="frmEditSolicitacaoColetaAdm" method="POST">
		<input type="hidden" name="hiddenReqColetaDomiciliar" value="<%=ReqColetaDomiciliar%>" />
		<input type="hidden" name="id" value="<%=Request.QueryString("idsolic")%>" />
		<input type="hidden" name="hiddenIsColetaEmail" value="<%=isColetaEmail()%>" />
		<input type="hidden" name="hiddenIdCliente" value="<%=IDCliente%>" />
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Administrar Solicitação de Coleta</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de coleta: </label></td>
				<td><%=NumSolColeta%> <img src="img/buscar.gif" class="imgexpandeinfo" align="absmiddle" alt="Buscar Solicitações que compuseram a solicitação Master" onClick="javascript:window.open('frmviewcompoemasteradm.asp?idsolic=<%=NumSolColeta%>','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"/></td>
			</tr>
			<tr id="tridcliente">
				<td width="35%" align="right"><label id="idcliente">ID. Ponto Coleta: </label></td>
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
			<tr id="trdddendcoleta">
				<td width="35%" align="right"><label id="dddendcoleta">DDD. end. Coleta: </label></td>
				<td><%=DDDEndColeta%></td>
			</tr>
			<tr id="trtelendcoleta">
				<td width="35%" align="right"><label id="telendcoleta">Tel. end. Coleta: </label></td>
				<td><%=TelEndColeta%></td>
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
				<td><input type="text" <%If isNull(QtdCartuchosRecebidos) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtQtdCatuchosRecebidos" value="<%=QtdCartuchosRecebidos%>" size="4" readonly="readonly" />&nbsp;<img src="img/produtos.gif" align="absmiddle" class="imgexpandeinfo" width="25" height="22" name="listprodutos" alt="Produtos" onClick="javascript:window.open('frmListaProdutosSolicitacao.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /></td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="status">Status: </label></td>
				<td>
					<!--
					<select name="cbStatusSolColeta" class="select">
						<option value="-1"> --- Selecione --- </option>	
						<%'Call GetStatusColeta()%>						
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
					<input type="text" <%If isNull(DataAprovacao) Then%>class="textreadonly" value="<%=DataAprovacao%>"<%ELse%>class="text" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" size="13" maxlength="10" onKeyPress="date(this)" readonly />					
				</td>
			</tr>
			<tr id="trdataenviotransportadora">
				<td width="35%" align="right"><label id="dataenviotransportadora">Data envio transportadora: </label></td>
				<td valign="bottom">
					<input type="text" <%If isNull(DataEnvioTransp) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataEnvioTransportadora" value="<%=DataEnvioTransp%>" size="13" maxlength="10" onKeyPress="date(this)" readonly />
            <A HREF="#" onClick="cal.select(document.forms['frmEditSolicitacaoColetaAdm'].txtDataEnvioTransportadora,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img align="absmiddle" src="img/btn_calendario.gif" border="0"></A> 
          </td>
			</tr>
			<tr id="tridtransportadora">
				<td width="35%" align="right">
					<label id="numidtransportadora">Transportadora: </label>
				</td>
				<td>
					<!--
					<select name="cbTransp" class="select" onChange="changeStatusKeyPress()">
						<option value="-1">Selecione uma Transportadora</option>
						<%'Call GetTransp()%>
					</select>
					<img src="img/transportadoras.gif" class="imgexpandeinfo" width="25" height="25" align="absmiddle" alt="Buscar Transportadora" onClick="window.open('frmSearchTranspColetaMaster.asp','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
					-->
					<%Call GetDescTransp()%>
					<INPUT type="hidden" id="cbTransp" name="cbTransp" value="<%=IdTranspHidden%>">
				</td>
			</tr>
			<tr id="trdataprogramada">
				<td width="35%" align="right"><label id="dataprogramada">Data programada: </label></td>
				<td valign="bottom">
					<input type="text" <%If isNull(DataProgramada) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataProgramada" value="<%=DataProgramada%>" size="13" maxlength="10" onKeyPress="date(this);changeStatusKeyPress()" readonly />
            <A HREF="#" onClick="cal.select(document.forms['frmEditSolicitacaoColetaAdm'].txtDataProgramada,'anchor1','dd/MM/yyyy'); return false;" NAME="anchor1" ID="anchor1"><img align="absmiddle" src="img/btn_calendario.gif" border="0"></A> 
          </td>
			</tr>
			<tr id="trnumconhtransportadora">
				<td width="35%" align="right"><label id="numconhtransportadora">Número conhecimento da Transportadora: </label></td>
				<td><input type="text" <%If isNull(GetNumRecTransportadora()) Or GetNumRecTransportadora() = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtNumConhTransportadora" value="<%=GetNumRecTransportadora()%>" size="15" /></td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Data recebimento: </label></td>
				<td><input type="text" <%If isNull(DataReceb) Or DataReceb = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataRecebimento" value="<%=DataReceb%>" size="13" maxlength="10" readonly="readonly" onKeyPress="date(this)" /></td>
			</tr>
			<%End If%>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr id="btnprove" <% If Not StatusAprovar Then %>style="display:none;"<%End If%>>
				<td align="right"><input type="button" class="btnform" name="btnAprovar" value="Aprovar" onClick="aprovar('<%= Request.QueryString("idsolic") %>')" /></td>
				<td align="left"><input type="button" class="btnform" name="btnReprovar" value="Reprovar" onClick="reprovar('<%= Request.QueryString("idsolic") %>')" /></td>
			</tr>
			<tr id="btnatualizar" <% If Not StatusAtualizar Then %>style="display:none;"<%End If%>>
				<td align="center" colspan="2">
					<input type="button" class="btnform" name="btnAtualizar" value="Atualizar" onClick="validateForm()" />
					<input type="button" class="btnform" name="btnReprovar" value="Cancelar" onClick="cancelar('<%= Request.QueryString("idsolic") %>')" />
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
