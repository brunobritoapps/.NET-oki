<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%

	'// Ponto de Coleta
  '-----------------------------
	Dim QuantidadeCartuchos
	Dim NumeroSolicitacaoColeta
	'-----------------------------
	Dim IDPontoColeta
	Dim NomeFantasiaPontoColeta
	Dim LogradouroPontoColeta
	'-----------------------------
	
	'// Coleta Domiciliar
	'-----------------------------
	Dim CEPColeta
	Dim LogradouroColeta
	Dim BairroColeta
	Dim MunicipioColeta
	Dim EstadoColeta
	Dim TipoLogradouroColeta
	Dim CompLogradouro
	Dim NumeroColeta
	Dim ContatoRespColeta
	Dim DDDContatoRespColeta
	Dim TelefoneContatoRespColeta
	'-----------------------------

	Sub RequestForm()
		QuantidadeCartuchos 				= Request.Form("txtQtdCartuchos")
		IDPontoColeta 						= Request.Form("hiddenIntChangePontoColeta")
		CEPColeta 							= Request.Form("txtCepColeta")
		CompLogradouro 						= Request.Form("txtCompLogradouroColeta")
		NumeroColeta 						= Request.Form("txtNumeroColeta")
		ContatoRespColeta 					= Request.Form("txtContatoRespColeta")
		DDDContatoRespColeta 				= Request.Form("txtDDDContatoRespColeta")
		TelefoneContatoRespColeta			= Request.Form("txtTelefoneContatoRespColeta")
		' Informações do endereço de coleta da empresa
		LogradouroColeta					= request.Form("txtLogradouroColeta")
		BairroColeta						= request.Form("txtBairroColeta")
		MunicipioColeta						= request.Form("txtMunicipioColeta")
		EstadoColeta						= request.Form("txtEstadoColeta")
		
		If Session("isColetaDomiciliar") = 0 Then
			If IDPontoColeta = "" Then
				Response.Redirect "frmAddSolicitacao.asp?MsgRet=Escolha uma Ponto de Coleta para envio da coleta"
			End If
		End If	
	End Sub

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
'			Response.Write Request.Form("hiddenActionForm")
'			Response.End()
			If Request.Form("hiddenActionForm") = 1 Then
				Call UpdateAdressColect()
				Response.Redirect "frmOperacionalCliente.asp"
			ElseIf Request.Form("hiddenActionForm") = 3 Then
				Call UpdateAdressColect()
				Call AddSolColeta()
			Else 	
				Call AddSolColeta()
			End If
		End If
	End Sub

	Sub AddSolColeta()
		Dim oCommand, rs
		Dim NumeroSequencial
		Dim NumeroSolicitacaoColeta
		Dim IdentifiedCharacterColeta
		Dim DateMonthColeta
		Dim DateYearColeta
		
'		response.write IDPontoColeta
'		response.End()
		
		Set oCommand = Server.CreateObject("ADODB.Command")
		Set rs = Server.CreateObject("ADODB.Recordset") 'jadilson
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandType = 4
		oCommand.CommandText = "sp_AddSolicitacaoColeta"

		If Session("isColetaDomiciliar") = 1 Then
			oCommand.Parameters("@IDPontoColeta")			= 0
			IdentifiedCharacterColeta						= "C"
		Else
			IdentifiedCharacterColeta						= "E"
			If Request.Form("hiddenIntChangePontoColeta") <> "" Then
				oCommand.Parameters("@IDPontoColeta")		= CInt(IDPontoColeta)
			Else
				oCommand.Parameters("@IDPontoColeta")		= CInt(Session("IDPontoColeta"))
			End If
		End If
		
		DateMonthColeta										= Month(Now())
		If Len(DateMonthColeta) = 1 Then
			DateMonthColeta									= "0" & DateMonthColeta
		End If
		DateYearColeta										= Right(Year(Now()), 2)

		NumeroSolicitacaoColeta								= IdentifiedCharacterColeta & DateYearColeta & DateMonthColeta
		NumeroSequencial									= getSequencial(False)
		NumeroSolicitacaoColeta								= NumeroSolicitacaoColeta & NumeroSequencial
		NumeroSolicitacaoColeta								= getDigitoControle(NumeroSolicitacaoColeta)
'		response.write NumeroSolicitacaoColeta
'		response.end

		oCommand.Parameters("@NumeroSolicitacaoColeta")		= NumeroSolicitacaoColeta
		oCommand.Parameters("@isColetaDomiciliar")			= CInt(Session("isColetaDomiciliar"))
		oCommand.Parameters("@QtdCartuchos")				= CInt(QuantidadeCartuchos)
		oCommand.Parameters("@IDClient")					= CInt(Session("IDCliente"))
		oCommand.Parameters("@IDContato")					= CInt(Session("IDContato"))

		'oCommand.Execute()
		rs.Open oCommand 'jadilson

		If Session("isColetaDomiciliar") = 1 Then
			Response.Write "<script>alert('Em breve entraremos em contato para " & _
										 "providenciar a coleta!');</script>"
		Else
			Response.Write "<script>alert('Em breve entraremos em contato para " & _
										 "autorizar a entrega do(s) cartucho(s) no ponto de coleta!')</script>"
		End If
		'Response.Write "<script>window.location.href='frmAddSolicitacao.asp?MsgRet=Solicitação efetuada com sucesso';</script>"
		
		Response.Write "<script>"
			Response.Write "window.location.href='frmCartaDoacaoNF.asp?Acao=0&IdSolicitacaoColeta=" & rs.fields(0) & "&Adm=1&TipoColeta=" & Session("isColetaDomiciliar") & "';"
			'Response.Write "window.open('frmCartaDoacaoNF.asp?IdSolicitacaoColeta="&rs.fields(0)&"&Acao=1&TipoPessoa=', '_blank');"
			'Response.Write "window.location.href='frmAddSolicitacao.asp?MsgRet=Solicitação efetuada com sucesso';"
		Response.Write "</script>"
		'Response.Write "TESTE PASSOU"
		
		Set oCommand = Nothing
	End Sub

	Sub GetListPontoColeta()
		Dim sSql, arrListPontoColeta, intListPontoColeta, i
		
'		sSql = "SELECT " & _ 
'						"[A].[Pontos_coleta_idPontos_coleta], " & _
'						"[B].[Nome_fantasia], " & _
'						"[B].[Numero_endereco], " & _
'						"[B].[Complemento_endereco], " & _
'						"[C].[Tipologradouro], " & _
'						"[C].[Logradouro], " & _
'						"[C].[Bairro], " & _
'						"[C].[Municipio] " & _
'						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [A] " & _
'						"LEFT JOIN [marketingoki2].[dbo].[Pontos_coleta] AS [B] " & _
'						"ON [A].[Pontos_coleta_idPontos_coleta] = [B].[idPontos_coleta] " & _
'						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta] AS [C] " & _
'						"ON [B].[cep_consulta_idcep_consulta] = [C].[idcep_consulta] " & _
'						"WHERE [A].[Clientes_idClientes] = " & Session("IDCliente") & " AND " & _
'						"[B].[status_pontocoleta] = 1"
						
		sSql = "select " & _
				"a.pontos_coleta_idpontos_coleta, " & _
				"b.nome_fantasia, " & _
				"b.numero_endereco, " & _
				"b.complemento_endereco, " & _
				"b.logradouro, " & _	
				"b.bairro, " & _
				"b.municipio " & _
				"from solicitacao_coleta_has_clientes as a " & _
				"left join pontos_coleta as b " & _
				"on a.pontos_coleta_idpontos_coleta = b.idpontos_coleta " & _
				"where a.clientes_idclientes = "&session("IDCliente")&" and b.status_pontocoleta = 1"						
						
		'idponto coleta = 0
		'nome fantasia  = 1
		'numero endereco= 2
		'comp endereco	= 3
		'logradouro		= 4
		'bairro			= 5
		'municipio		= 6
		
		Call search(sSql, arrListPontoColeta, intListPontoColeta)
		If intListPontoColeta > -1 Then
			For i=0 To intListPontoColeta
				IDPontoColeta				= arrListPontoColeta(0,i)
				NomeFantasiaPontoColeta		= arrListPontoColeta(1,i)
				LogradouroPontoColeta		= arrListPontoColeta(4,i) & ", n° " & arrListPontoColeta(2,i) & ", " & arrListPontoColeta(5,i) & ", " & arrListPontoColeta(6,i)
			Next
		End If
	End Sub

	Sub GetListDomiciliar()
		Dim sSql, arrListDomiciliar, intListDomiciliar, i
		
		sSql = "SELECT " & _
						"[A].[cep], " & _ 
						"[A].[logradouro], " & _ 
						"[A].[bairro], " & _ 
						"[A].[municipio], " & _ 
						"[A].[estado], " & _ 
						"[C].[compl_endereco_coleta], " & _
						"[C].[numero_endereco_coleta], " & _ 
						"[C].[contato_respcoleta], " & _
						"[C].[ddd_respcoleta], " & _
						"[C].[telefone_respcoleta] " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS [A] " & _ 
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _ 
						"ON [A].[Clientes_idClientes] = [C].[idClientes] " & _
						"WHERE [A].[isEnderecoColeta] = 1 AND " & _ 
						"[A].[Clientes_idClientes] = " & Session("IDCliente")

		Call search(sSql, arrListDomiciliar, intListDomiciliar)
		If intListDomiciliar > -1 Then
			For i=0 To intListDomiciliar
				CEPColeta						= arrListDomiciliar(0,i)
				LogradouroColeta				= arrListDomiciliar(1,i)
				BairroColeta					= arrListDomiciliar(2,i)
				MunicipioColeta					= arrListDomiciliar(3,i)
				EstadoColeta					= arrListDomiciliar(4,i)
				CompLogradouro					= arrListDomiciliar(5,i)
				NumeroColeta					= arrListDomiciliar(6,i)
				ContatoRespColeta				= arrListDomiciliar(7,i)
				DDDContatoRespColeta			= arrListDomiciliar(8,i)
				TelefoneContatoRespColeta		= arrListDomiciliar(9,i)
			Next
		End If
	End Sub
	
	Function GetCepEnderecoComum()
		Dim sSql, arrCep, intCep, i
		Dim CEP
		
		sSql = "SELECT " & _
						"A.cep " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 1"
						
		Call search(sSql, arrCep, intCep)
		
		If intCep > -1 Then
			For i=0 To intCep
				CEP = arrCep(0,i)
			Next
		Else
			CEP = Null	
		End If

		GetCepEnderecoComum	= CEP			

	End Function
	
	Function GetNumeroEnderecoCliente()
		Dim sSql, arrNumero, intNumero, i
		Dim Numero
		
		sSql = "SELECT [numero_endereco] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")
		
		Call search(sSql, arrNumero, intNumero)
		
		If intNumero > -1 Then
			For i=0 To intNumero
				Numero = arrNumero(0,i)
			Next
		Else
			Numero = Null	
		End If
		
		GetNumeroEnderecoCliente = Numero
		
	End Function
	
	Function GetCompLogradouroEnderecoCliente()
		Dim sSql, arrComp, intComp, i
		Dim CompLogradouro
		
		sSql = "SELECT [compl_endereco] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")
		
		Call search(sSql, arrComp, intComp)
		
		If intComp > -1 Then
			For i=0 To intComp
				GetCompLogradouroEnderecoCliente = arrComp(0,i)
			Next
		Else
			GetCompLogradouroEnderecoCliente = Null
		End If
						
	End Function
	
	Sub UpdateAdressColect()
		Dim oCommand
		
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandType = 4
		oCommand.CommandText = "sp_UpdateClienteColeta"
		
		' informações sobre o cliente e as atualizações que estavam fazendo normalmente		
		oCommand.Parameters("@IDCliente")						= Session("IDCliente")
		oCommand.Parameters("@CompEndereco")					= CompLogradouro
		oCommand.Parameters("@NumeroEndereco")					= CLng(NumeroColeta)
		oCommand.Parameters("@ContatoRespColeta")				= ContatoRespColeta
		oCommand.Parameters("@DDDContatoRespColeta")			= CInt(DDDContatoRespColeta)
		oCommand.Parameters("@TelefoneContatoRespColeta")		= CLng(TelefoneContatoRespColeta)
		
		' atualização que foi feita no endereço
		oCommand.Parameters("@cep_coleta")						= CEPColeta
		oCommand.Parameters("@logradouro_coleta")				= LogradouroColeta
		oCommand.Parameters("@bairro_coleta")					= BairroColeta
		oCommand.Parameters("@municipio_coleta")				= MunicipioColeta
		oCommand.Parameters("@estado_coleta")					= EstadoColeta
		
		oCommand.Execute()
		
		Set oCommand = Nothing
		
'		Response.Redirect "frmOperacionalCliente.asp"
		
	End Sub
	
	Function MinCartuchos()
		Dim sSql, arrMin, intMin, i
		
		sSql = "SELECT [minCartuchos] " & _
			   "FROM [marketingoki2].[dbo].[Clientes] " & _
			   "WHERE [idClientes] = " & Session("IDCliente")
		Call search(sSql, arrMin, intMin)
		If intMin > -1 Then
			For i=0 To intMin
				minCartuchos = arrMin(0,i)	
			Next			   
		End If
	End Function
	
	Function GetCategoria()
		Dim sSql, arrCategoria, intCategoria, i
		Dim minCartuchos
		
		sSql = "SELECT B.minCartuchos FROM Clientes AS A " & _
				"LEFT JOIN Categorias AS B " & _
				"ON B.idCategorias = A.Categorias_idCategorias " & _
				"WHERE A.idClientes = " & Session("IDCliente")
				
		Call search(sSql, arrCategoria, intCategoria)
		
		For i=0 To intCategoria
			minCartuchos = arrCategoria(0,i)
		Next
		
		GetCategoria = minCartuchos
	End Function
	
	If Session("isColetaDomiciliar") = 1 Then	
		Call GetListDomiciliar()
	Else
		Call GetListPontoColeta()
	End If		

	Call SubmitForm()	
	
'	Response.Write "Categoria: " & GetCategoria() & "<br />"
'	Response.Write "Min. Cartuchos: " & MinCartuchos() & "<br />"
%>

<!DOCTYPE html>

<html>
<head>
	<script src="js/frmAddSolicitacao.js"></script>
	<link rel="stylesheet" type="text/css" href="css/geral.css">
	<link rel="stylesheet" type="text/css" href="css/component.css" />
	<title><%=TITLE%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>


<%If Session("isColetaDomiciliar") = 0 Then%>
<body>
<%else%>
<body onLoad="loadInfoSameAdress()">
<%end if%>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		
			<form action="frmAddSolicitacao.asp" name="frmAddSolicitacao" method="POST">
			
			<input type="hidden" name="hiddenSessionisColetaDomiciliar" value="<%=Session("isColetaDomiciliar")%>" />
			<input type="hidden" name="hiddenIntPontoColeta" value="" />
			<input type="hidden" name="hiddenIntChangePontoColeta" value="<%=IDPontoColeta%>" />
			<input type="hidden" name="hiddenIntEnderecoCepColeta" value="" />
			<input type="hidden" name="hiddenGetCepEnderecoComum" value="<%=GetCepEnderecoComum()%>" />
			<input type="hidden" name="hiddenGetNumeroEnderecoCliente" value="<%=GetNumeroEnderecoCliente()%>" />
			<input type="hidden" name="hiddenGetCompLogradouroEnderecoCliente" value="<%=GetCompLogradouroEnderecoCliente()%>" />
			<input type="hidden" name="hiddenActionForm" value="0" />
			<input type="hidden" name="hiddenMinCartuchos" value="<%If MinCartuchos() = "" Or MinCartuchos() = 0 Then Response.Write GetCategoria() Else Response.Write MinCartuchos() End If %>" />

			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table border="0" cellpadding="0" width="100%">
						<tr> 
							<td>
								<img src="img/novacoleta.png">
							</td>
							<td align="left" class="oki-h1"> 
								&nbsp;SOLICITAÇÂO DE COLETA
							</td>
						</tr>
						<tr>
							<td >&nbsp;</td>
							<td align="left" class="oki-h2">Você pode solicitar a sua coleta no endereço da sua empresa, ou solicitar a coleta em um novo endereço.</br>
							Para quantidade de itens inferior a 5 itens, procure um ponto de coleta mais próximo.
							<!--<td colspan="3" id="explaintitle" align="center">Nova Solicitação</td>-->
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
			
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table border="0" cellpadding="0" width="100%">
						<tr>
								<td width="70%">&nbsp;</td>
								<td width="15%">&nbsp;</td>
								<td width="15%">&nbsp;</td>								
								<td width="15%">&nbsp;</td></tr>
						<tr>
								<td width="70%">&nbsp;</td>
								<td align="center" colspan="1" align="left" class="gn-menu-main">
									<a href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Salvar</a>
								</td>
								<td align="center" colspan="1" align="left" class="gn-menu-main">
									<a href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a>
								</td>				
								<td width="15%">&nbsp;</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
			
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table width="100%" id="tableAddSolicitacao" border="0">
						<tr>
						<%If Request.QueryString("MsgRet") <> "" Then%>
							<tr>
								<td colspan="3" align="center"><b style="color:#FF0000;"><%=Request.QueryString("MsgRet")%>!</b></td>
							</tr>
						<%End If%>
						</tr>
						<tr>
							<td align="left" colspan="4" class="oki-h2-negrito">&nbsp;Informações desta coleta:&nbsp;</br></td></br>
						</tr>
						<tr>
							<td width="20%" width="10%" align="right">Quantidade para coleta:</td>
							<td align="left" width="10%"><input type="number" class="oki-input" name="txtQtdCartuchos" min="1" max="1000" required ></td>
							<td width="60%">&nbsp;</td>
							<td width="10%">&nbsp;</td>
						</tr>
						<div id="radio">
						<tr>
							<td width="20%" width="10%" align="right">Coletar no endereço da empresa:</td>
							<td align="left" width="10%"><input id="radioendmesmo" onclick="preencheNovoEndereco();" type="radio" class="oki-input" name="tipoEndereco"></td>
							<td align="left" class="oki-h2" id="tagendmesmo" width="100%">
							<!--<td align="left" class="oki-h2" id="tagendmesmo" style="display: none; background:#f1f1f1" width="100%">-->
							<!--
									<%=LogradouroColeta%>,<%=NumeroColeta%></br>
									<%=CompLogradouro%></br>
									<%=BairroColeta%></br>
									<%=MunicipioColeta%> - <%=EstadoColeta%></br>
									<%=ContatoRespColeta%></br>
									<%=DDDContatoRespColeta%>-<%=TelefoneContatoRespColeta%>-->
									</td>
							<td width="10%">&nbsp;</td>
						</tr>
						<tr>
							<td width="20%" width="10%" align="right">Coletar Em Outro Endereço:</td>
							<!--<td align="left" width="10%"><input type="radio" class="oki-input" name="txtQtdCartuchos" onClick="preencheMesmoEndereco()"></td>-->
							<td align="left" width="10%"><input id="radioendnovo" type="radio" class="oki-input" name="tipoEndereco" onclick="preencheMesmoEndereco();"></td>
							<td width="60%">&nbsp;</td>
							<td width="10%">&nbsp;</td>
						</tr>
						</div>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>

			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table id="tagendnovo" width="100%" id="tableAddSolicitacao" border="0" style="display: none;" class="oki-addnovoend" >
					<td id="conteudo">
						<%If Session("isColetaDomiciliar") = 0 Then%>
							<tr>
								<td colspan="3" id="explaintitle" align="center">Busca de Pontos de Coleta mais próximos</td>
							</tr>
							<tr>
								<td colspan="3"><b style="color:#FF0000;">Ponto de Coleta onde foi efetuada a última Solicitação:-</b></td>
							</tr>
							<tr>
								<td colspan="3">
									<table cellpadding="3" cellspacing="1" width="100%" id="tablRelSolPontoColetaEdita">
										<tr>
											<th width="5%">ID</th>
											<th>Nome Fantasia</th>
											<th>Logradouro</th>
										</tr>
										<tr>
											<td><%=IDPontoColeta%></td>
											<td><%=NomeFantasiaPontoColeta%></td>
											<td><%=LogradouroPontoColeta%></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td colspan="2">CEP para pesquisa dos pontos de coleta: <input type="text" class="textreadonly" maxlength="8" name="txtCepConsultaPonto" value="" size="11" />
								<img align="absmiddle" style="cursor:pointer;" src="img/buscar.gif" name="btnBuscarCepColeta" id="btnBuscarCepColeta" alt="Buscar CEP" onClick="showClientePostoColeta()" /></td>
							</tr>
							<tr id="titTableListPontoColeta" style="display:none;">
								<td colspan="2" align="center"><b>Pontos de Coleta mais próximos</b></td>
							</tr>
							<tr>
								<td colspan="3" id="tableListPontoColeta"></td>
							</tr>
						
						<%Else%>
						
							<tr>
								<td align="right" width="25%">Cep de Coleta:</td>
								<td align="left">
									<input type="text" class="oki-input" name="txtCepColeta" value="<%=CEPColeta%>" size="10" maxlength="8" /> Formato: 99999999
									<img align="absmiddle" style="cursor:pointer;" src="img/buscar.gif" name="btnBuscarCepColeta" id="btnBuscarCepColeta" alt="Buscar CEP" onClick="loadCepColeta()" />
								</td>
							</tr>
							<tr>
								<td align="right" width="25%">Logradouro:</td>
								<td align="left"><input type="text" class="oki-input" name="txtLogradouroColeta" value="<%=LogradouroColeta%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Complemento Logradouro:</td>
								<td align="left"><input type="text" class="oki-input" name="txtCompLogradouroColeta" value="<%=CompLogradouro%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Número:</td>
								<td align="left"><input type="text" class="oki-input" name="txtNumeroColeta" value="<%=NumeroColeta%>" size="10" maxlength="8" /> *</td>
							</tr>
							<tr>
								<td align="right" width="25%">Bairro:</td>
								<td align="left"><input type="text" class="oki-input" name="txtBairroColeta" value="<%=BairroColeta%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Município:</td>
								<td align="left"><input type="text" class="oki-input" name="txtMunicipioColeta" value="<%=MunicipioColeta%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Estado:</td>
								<td align="left"><input type="text" class="oki-input" name="txtEstadoColeta" value="<%=EstadoColeta%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Contato para Coleta:</td>
								<td align="left"><input type="text" class="oki-input" name="txtContatoRespColeta" value="<%=ContatoRespColeta%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">DDD do Contato:</td>
								<td align="left"><input type="text" class="oki-input" name="txtDDDContatoRespColeta" value="<%=DDDContatoRespColeta%>" size="3" maxlength="2" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Telefone do Contato:</td>
								<td align="left"><input type="text" class="oki-input" name="txtTelefoneContatoRespColeta" value="<%=TelefoneContatoRespColeta%>" size="10" maxlength="8" /></td>
							</tr>
							<tr>
								<td colspan="2" align="center"><input type="button" class="btnform" name="btnChangeAdressColect" value="Alterar Endereço" onClick="authenticateUpdateAdress()" /></td>
							</tr>
						<%End If%>
					</td>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
			</div>
			</form>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
