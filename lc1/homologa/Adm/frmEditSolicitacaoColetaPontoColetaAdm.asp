<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim NumSolColeta
	Dim QtdCartuchos
	Dim QtdCartuchosRecebidos
	Dim DataSolicitacao
	Dim DataAprovacao
	Dim DataEntPontoColeta
	Dim DataReceb
	Dim StatusSol
	Dim MotivoStatus
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
	Dim IDCliente
	Dim IDPontoColeta
	Dim StatusAprovar
	Dim StatusAtualizar
	Dim RazaoSocialPontoColeta
	dim tipopessoaquerystring
	
	ReqColetaDomiciliar = Request.QueryString("iscoletadomiciliar")
	tipopessoaquerystring = request.querystring("tipopessoa")

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
				QtdCartuchosRecebidos 				= arrSolicitacao(3,i)
				DataSolicitacao 					= arrSolicitacao(4,i)
				DataAprovacao 		 				= arrSolicitacao(5,i)
				DataEntPontoColeta 					= arrSolicitacao(8,i)
				DataReceb 							= arrSolicitacao(9,i)
				MotivoStatus 						= arrSolicitacao(10,i)
			Next
			
		If Left(Request.ServerVariables("LOCAL_ADDR"), 3) = "127" Then
			If Not isNull(DataSolicitacao)			Then DataSolicitacao	= FormatDateTime(DataSolicitacao, 2) 						
			If Not isNull(DataAprovacao)			Then DataAprovacao		= FormatDateTime(DataAprovacao, 2) 						
			If Not isNull(DataEntPontoColeta)		Then DataEntPontoColeta = FormatDateTime(DataEntPontoColeta, 2)
			If Not isNull(DataReceb)				Then DataReceb			= FormatDateTime(DataReceb, 2)
		Else			
			If Not isNull(DataSolicitacao)			Then DataSolicitacao	= DateRight(FormatDateTime(DataSolicitacao, 2)) 						
			If Not isNull(DataAprovacao)			Then DataAprovacao		= DateRight(FormatDateTime(DataAprovacao, 2)) 						
			If Not isNull(DataEntPontoColeta)		Then DataEntPontoColeta = DateRight(FormatDateTime(DataEntPontoColeta, 2))
			If Not isNull(DataReceb)				Then DataReceb			= DateRight(FormatDateTime(DataReceb, 2))
		End If	

		End If
		Call GetCliente()
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
	
	function getCheckCliente(id)
		dim sql, arr ,intarr, i
		
		sql = "SELECT A.[idClientes] " & _
			  ",A.[status_cliente] " & _
			  "FROM [marketingoki2].[dbo].[Clientes] AS A " & _
			  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_coleta_has_clientes] AS B " & _
			  "ON B.[Clientes_idClientes] = A.[idClientes] " & _
			  "WHERE A.[status_cliente] = 1 AND B.[Solicitacao_coleta_idSolicitacao_coleta] = " & id
		
		call search(sql, arr, intarr)	  
		if intarr > -1 then
			if not isnull(arr(0,i)) and not isempty(arr(0,i)) then
				getCheckCliente = true
			else
				getCheckCliente = false
			end if
		else
			getCheckCliente = false
		end if
	end function

	
	Sub GetPontoColeta()
		Dim sSql, arrPontoCol, intPontoCol, i
			
		sSql = "SELECT " & _
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
				"D.nome, " & _ 
				"A.ddd, " & _ 
				"A.telefone " & _ 
				"FROM Pontos_coleta AS A " & _ 
				"LEFT JOIN Solicitacao_coleta_has_Clientes AS C " & _ 
				"ON A.idPontos_coleta = C.Pontos_coleta_idPontos_coleta " & _ 
				"LEFT JOIN Contatos AS D " & _ 
				"ON C.Contatos_idContatos = D.idContatos " & _ 
				"WHERE C.Solicitacao_coleta_idSolicitacao_coleta = "&Request.QueryString("idsolic")&" " & _ 
				"AND C.Clientes_idClientes = " & IDCliente
				
				'response.Write sSql
				'response.End		

				'A.idPontos_coleta		= 0
				'A.razao_social			= 1  
				'A.cnpj					= 2 
				'A.numero_endereco		= 3
				'A.complemento_endereco = 4 
				'A.cep					= 5
				'A.logradouro			= 6
				'A.bairro				= 7
				'A.municipio			= 8
				'A.estado				= 9
				'D.nome					= 10
				'A.ddd					= 11
				'A.telefone				= 12

		Call search(sSql, arrPontoCol, intPontoCol)
		
		If intPontoCol > -1 Then
			For i=0 To intPontoCol
				IDPontoColeta 	 = arrPontoCol(0,i)
				CEP			  	 = arrPontoCol(5,i)
				LogradouroColeta = arrPontoCol(6,i) & " - " & arrPontoCol(7,i)
				MunEndColeta 	 = arrPontoCol(8,i)
				UFEndColeta 	 = arrPontoCol(9,i)
				NumEndColeta  	 = arrPontoCol(3,i)
				CompEndColeta 	 = arrPontoCol(4,i)
				ContatoColeta 	 = arrPontoCol(10,i)
				DDDEndColeta	 = arrPontoCol(11,i)	
				TelEndColeta	 = arrPontoCol(12,i)
				RazaoSocialPontoColeta = arrPontoCol(1,i)	
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
		QtdCartuchosRecebidos 				  = Request.Form("txtQtdCatuchosRecebidos")
		DataAprovacao 						  = Request.Form("txtDataAprovacao")
		DataReceb 							  = Request.Form("txtDataRecebimento")
		StatusSol 							  = Request.Form("cbStatusSolColeta")
		MotivoStatus 						  = Request.Form("txtMotivoStatus")
		DataEntPontoColeta 					  = Request.Form("txtDataEntregaPontoColeta")

		If QtdCartuchosRecebidos 			  = "" Then QtdCartuchosRecebidos = "NULL" 
		If DataReceb 						  = "" Then DataReceb = "NULL" Else DataReceb = "CONVERT(DATETIME, '"&FormatDate(DataReceb)&"')" End If 
		If DataAprovacao 				 	  = "" Then DataAprovacao = "NULL" Else DataAprovacao = "CONVERT(DATETIME, '"&FormatDate(DataAprovacao)&"')" End If 
		If DataEntPontoColeta 	 			  = "" Then DataEntPontoColeta = "NULL" Else DataEntPontoColeta = "CONVERT(DATETIME, '"&FormatDate(DataEntPontoColeta)&"')" End If
		If MotivoStatus 				 	  = "" Then MotivoStatus = "NULL"
		
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
		
		if getCheckCliente(request.form("id")) then
			sSql = "UPDATE [marketingoki2].[dbo].[Solicitacao_coleta] " & _
					"SET [Status_coleta_idStatus_coleta] = "&StatusSol&" " & _
					  ",[data_aprovacao]				 = "&DataAprovacao&" " & _
					  ",[data_programada]				 = NULL " & _
					  ",[data_envio_transportadora]		 = NULL " & _
					  ",[data_entrega_pontocoleta]		 = "&DataEntPontoColeta&" " & _
					  ",[data_recebimento]				 = "&DataReceb&" " & _
					  ",[motivo_status]					 = '"&MotivoStatus&"' " & _
					"WHERE [idSolicitacao_coleta]		 = " & Request.Form("id")
	'		Response.Write sSql
	'		Response.End()		
			Call exec(sSql)
			response.write "<script>window.opener.location.reload();</script>"
			Response.Write "<script>window.parent.close();</script>"							
		end if	
	End Sub
	
	Function FormatDate(sDate)
		Dim Ano
		Dim Mes
		Dim Dia
		'Dia = Left(sDate, 2)
		'Mes = Mid(sDate, 4, 2)
		'Mes = Replace(Mes, "/" ,"")
		'If Len(Mes) = 1 Then
		'	Mes = "0" & Mes
		'End If	
		'Ano = Right(sDate, 4)
		Dia = day(sDate)
		Mes = month(sDate)
		Ano = year(sDate)
		
		FormatDate = Ano & "/" & Mes & "/" & Dia
	End Function
	
	function getSolMaster()
		dim sql, arr, intarr, i
		sql = "select " & _
				"a.id_solicitacao, " & _ 
				"a.numero_solicitacao_master, " & _
				"b.numero_solicitacao_coleta " & _
				"from solicitacoes_baixadas as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.id_solicitacao = b.idsolicitacao_coleta " & _
				"where a.is_baixada = 1 and a.id_solicitacao = " & Request.QueryString("idsolic")
		call search(sql, arr, intarr)				
		if intarr > -1 then
			for i=0 to intarr
				getSolMaster = arr(1,i)
			next	
		else
			getSolMaster = ""
		end if
	end function
	
	Call SubmitForm()
	
	'Response.Write CInt(StatusSol) & "<hr>"
	'Response.End
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
<script language="javascript" src="js/frmEditSolicitacaoColetaPontoColetaAdm.js"></script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmEditSolicitacaoColetaPontoColetaAdm.asp" name="frmEditSolicitacaoColetaPontoColetaAdm" method="POST">
		<input type="hidden" name="hiddenReqColetaDomiciliar" value="<%=ReqColetaDomiciliar%>" />
		<input type="hidden" name="id" value="<%=Request.QueryString("idsolic")%>" />
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Administrar Solicitação de Coleta</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td width="35%" align="right"><label id="numsolcoleta">Num. solic. de coleta: </label></td>
				<td><%=NumSolColeta%></td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="numsolcoletamaster">Num. solic. de coleta Master: </label></td>
				<td><%=getSolMaster()%></td>
			</tr>
			<tr id="tridcliente">
				<td width="35%" align="right"><label id="idcliente">ID. Cliente: </label></td>
				<td><%=IDCliente%></td>
			</tr>
			<%if cint(tipopessoaquerystring) <> 0 then%>
			<tr id="trrazaosocial">
				<td width="35%" align="right"><label id="razaosocial">Razão Social: </label></td>
				<td><%=RazaoSocial%></td>
			</tr>
			<tr id="trnomefantasia">
				<td width="35%" align="right"><label id="nomefantasia">Nome Fantasia: </label></td>
				<td><%=NomeFantasia%></td>
			</tr>
			<%else%>
			<tr id="trnomefantasia">
				<td width="35%" align="right"><label id="nomefantasia">Nome: </label></td>
				<td><%=NomeFantasia%></td>
			</tr>
			<%end if%>
			<tr id="trcontatopontocoleta">
				<td width="35%" align="right"><label id="contatopontocoleta">Contato Coleta: </label></td>
				<td><%=ContatoColeta%></td>
			</tr>
			<tr>
				<td width="35%" align="right"><label id="numpontocoleta">Razão Social Ponto de Coleta: </label></td>
				<td><%=RazaoSocialPontoColeta%></td>
			</tr>
			<tr id="trnumpontocoleta">
				<td width="35%" align="right"><label id="numpontocoleta">Num. Ponto de Coleta: </label></td>
				<td><%=IDPontoColeta%></td>
			</tr>
			<tr id="trcepcoleta">
				<td width="35%" align="right"><label id="cepcoleta">CEP Ponto de Coleta: </label></td>
				<td><%=CEP%></td>
			</tr>
			<tr id="trlogcoleta">
				<td width="35%" align="right"><label id="logcoleta">Logradouro Ponto Coleta: </label></td>
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
				<td width="35%" align="right"><label id="dddendcoleta">DDD. Ponto Coleta: </label></td>
				<td><%=DDDEndColeta%></td>
			</tr>
			<tr id="trtelendcoleta">
				<td width="35%" align="right"><label id="telendcoleta">Tel. Ponto Coleta: </label></td>
				<td><%=TelEndColeta%></td>
			</tr>
			<tr id="trdatasolicitacao">
				<td width="35%" align="right"><label id="datasolicitacao">Data solicitação: </label></td>
				<td><%=DataSolicitacao%></td>
			</tr>
			<tr id="trqtdcartuchos">
				<td width="35%" align="right"><label id="qtdcartuchos">Qtd. cartuchos a serem entregues: </label></td>
				<td><%=QtdCartuchos%></td>
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
				<td><input type="text" <%If isNull(DataAprovacao) Then%>class="textreadonly" value="<%=DataAprovacao%>"<%ELse%>class="text" value="<%=DataAprovacao%>" <%End If%> name="txtDataAprovacao" readonly="readonly" size="13" maxlength="10" onKeyPress="date(this)" /></td>
			</tr>
			<tr id="trdataentregapontocoleta">
				<td width="35%" align="right"><label id="dataentregapontocoleta">Data entrega Ponto Coleta: </label></td>
				<td>
					<input type="text" <%If isNull(DataEntPontoColeta) Or DataEntPontoColeta = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataEntregaPontoColeta" value="<%=DataEntPontoColeta%>" size="13" maxlength="10" onKeyPress="date(this)" readonly />					
				</td>
			</tr>
			<tr id="trqtdcartuchosrecebidos">
				<td width="35%" align="right"><label id="qtdcartuchosrecebidos">Qtd. cartuchos recebidos: </label></td>
				<td><input type="text" <%If isNull(QtdCartuchosRecebidos) Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtQtdCatuchosRecebidos" readonly="readonly" value="<%=QtdCartuchosRecebidos%>" size="4" />&nbsp;<img src="img/produtos.gif" align="absmiddle" class="imgexpandeinfo" width="25" height="22" name="listprodutos" alt="Produtos" onClick="javascript:window.open('frmListaProdutosSolicitacao.asp?idsol=<%=Request.QueryString("idsolic")%>','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /></td>
			</tr>
			<tr id="trdatarecebimento">
				<td width="35%" align="right"><label id="datarecebimento">Data recebimento: </label></td>
				<td><input type="text" <%If isNull(DataReceb) Or DataReceb = "" Then%>class="textreadonly"<%ELse%>class="text"<%End If%> name="txtDataRecebimento" value="<%=DataReceb%>" readonly="readonly" size="13" maxlength="10" onKeyPress="date(this)" /></td>
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
			<% If getCheckCliente(Request.QueryString("idsolic")) Then %>			
			<tr id="btnprove" <% If Not StatusAprovar Then %>style="display:none;"<%End If%>>
				<td align="right"><input type="button" class="btnform" name="btnAprovar" value="Aprovar" onClick="aprovar('<%= Request.QueryString("idsolic") %>')" /></td>
				<td align="left"><input type="button" class="btnform" name="btnReprovar" value="Rejeitar" onClick="reprovar('<%= Request.QueryString("idsolic") %>')" /></td>
			</tr>
			<tr id="btnatualizar" <% If Not StatusAtualizar Then %>style="display:none;"<%End If%>>
				<td align="center" colspan="2">
					<input type="button" class="btnform" name="btnAtualizar" value="Atualizar" onClick="validateForm()" />
					<input type="button" class="btnform" name="btnReprovar" value="Cancelar" onClick="cancelar('<%= Request.QueryString("idsolic") %>')" />	
				</td>
			</tr>
			<% End If %>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
