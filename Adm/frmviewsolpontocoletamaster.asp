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
			
		if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
			If Not isNull(DataSolicitacao) Then DataSolicitacao = DateRight(FormatDateTime(DataSolicitacao, 2)) 						
			If Not isNull(DataAprovacao) Then DataAprovacao = DateRight(FormatDateTime(DataAprovacao, 2)) 						
			If Not isNull(DataProgramada) Then DataProgramada = DateRight(FormatDateTime(DataProgramada, 2))						
			If Not isNull(DataEnvioTransp) Then DataEnvioTransp = DateRight(FormatDateTime(DataEnvioTransp, 2))						
			If Not isNull(DataEntPontoColeta) Then DataEntPontoColeta = DateRight(FormatDateTime(DataEntPontoColeta, 2))
			If Not isNull(DataReceb) Then DataReceb = DateRight(FormatDateTime(DataReceb, 2))
		else
			If Not isNull(DataSolicitacao) Then DataSolicitacao = FormatDateTime(DataSolicitacao, 2) 						
			If Not isNull(DataAprovacao) Then DataAprovacao = FormatDateTime(DataAprovacao, 2) 						
			If Not isNull(DataProgramada) Then DataProgramada = FormatDateTime(DataProgramada, 2)						
			If Not isNull(DataEnvioTransp) Then DataEnvioTransp = FormatDateTime(DataEnvioTransp, 2)						
			If Not isNull(DataEntPontoColeta) Then DataEntPontoColeta = FormatDateTime(DataEntPontoColeta, 2)
			If Not isNull(DataReceb) Then DataReceb = FormatDateTime(DataReceb, 2)
		end if	

		End If
		Call GetPontoColeta()
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
				"A.estado " & _ 
				"FROM Pontos_coleta AS A " & _
				"LEFT JOIN Solicitacao_coleta_has_Pontos_coleta AS C " & _ 
				"ON A.idPontos_coleta = C.Pontos_coleta_idPontos_coleta " & _ 
				"WHERE C.Solicitacao_coleta_idSolicitacao_coleta = "&Request.QueryString("idsolic")
'		response.write sSql
'		response.end		
				
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
	
	Sub SubmitForm()
			Call GetSolicitacao()
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

	if (document.frmEditSolicitacaoEntrega.txtQtdCatuchosRecebidos == '') {
		error++;
	}

	if (error == 0) {
		document.frmEditSolicitacaoEntrega.submit();		
	}
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
				<td id="explaintitle" colspan="2" align="center">Status da Solicitação de Coleta</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de coleta: </label></td>
				<td><%=NumSolColeta%> <img src="img/buscar.gif" class="imgexpandeinfo" align="absmiddle" alt="Buscar Solicitações que compuseram a solicitação Master" onClick="javascript:window.open('frmviewcompoemaster.asp?idsolic=<%=NumSolColeta%>','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"/></td>
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
			<tr id="trcontatocoleta">
				<td width="35%" align="right"><label id="contatocoleta">Contato Coleta: </label></td>
				<td><%=ContatoColeta%></td>
			</tr>
			-->
			<tr id="trqtdcartuchos">
				<td width="35%" align="right"><label id="qtdcartuchos">Qtd. cartuchos: </label></td>
				<td><%=QtdCartuchos%></td>
			</tr>
			<tr id="trqtdcartuchosrecebidos">
				<td width="35%" align="right"><label id="qtdcartuchosrecebidos">Qtd. cartuchos recebidos: </label></td>
				<td>
				<%if len(trim(QtdCartuchosRecebidos)) = 0 or isnull(QtdCartuchosRecebidos) then%>
					<input type="text" <%If isNull(QtdCartuchosRecebidos) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtQtdCatuchosRecebidos" value="<%=QtdCartuchosRecebidos%>" readonly="true" size="4" />
				<%else%>
					<%=QtdCartuchosRecebidos%>
				<%end if%>
				<img align="absmiddle" class="imgexpandeinfo" alt="Verificação de Produtos da Solicitação" src="img/buscar.gif" onClick="javascript:window.open('frmListaProdutosSolicitacao.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" />
				</td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="status">Status: </label></td>
				<td>
					<select name="cbStatusSolColeta" class="select" disabled="disabled">
						<option value="-1"> --- Selecione --- </option>	
						<%Call GetStatusColeta()%>
					</select>
				</td>
			</tr>
			<tr id="trdatasolicitacao">
				<td width="35%" align="right"><label id="datasolicitacao">Data solicitação: </label></td>
				<td><%=DataSolicitacao%></td>
			</tr>
			<tr id="trdataaprovacao">
				<td width="35%" align="right"><label id="dataaprovacao">Data aprovação: </label></td>
				<td><input type="text" <%If isNull(DataAprovacao) Then%>class="textreadonly" value="<%=DataAprovacao%>"<%ELse%>class="text" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" size="13" maxlength="10" readonly="true" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="trdataenviotransportadora">
				<td width="35%" align="right"><label id="dataenviotransportadora">Data envio transportadora: </label></td>
				<td><input type="text" <%If isNull(DataEnvioTransp) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataEnvioTransportadora" value="<%=DataEnvioTransp%>" size="13" maxlength="10" readonly="true" onKeyPress="date(this);" /></td>
			</tr>
			<tr id="trdataprogramada">
				<td width="35%" align="right"><label id="dataprogramada">Data programada: </label></td>
				<td><input type="text" <%If isNull(DataProgramada) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataProgramada" value="<%=DataProgramada%>" size="13" maxlength="10" readonly="true" onKeyPress="date(this);changeStatusKeyPress()" /></td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Data recebimento: </label></td>
				<td><input type="text" <%If isNull(DataReceb) Or DataReceb = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataRecebimento" value="<%=DataReceb%>" size="13" maxlength="10" readonly="true" onKeyPress="date(this)" /></td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
