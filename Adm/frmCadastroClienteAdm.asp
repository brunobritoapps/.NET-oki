<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
    'Response.Expires = -1
	'Server.ScriptTimeout = 600

	Dim RazaoSocial
	Dim NomeFantasia
	Dim CNPJ
	Dim InscEstadual
	Dim DDD
	Dim Telefone
	Dim Categoria
	Dim Grupo
	Dim IDCep
	Dim CEP
	Dim LogCliente
	Dim CompLog
	Dim Numero
	Dim Bairro
	Dim Municipio
	Dim Estado
	Dim IDCepColeta
	Dim CEPColeta
	Dim LogColeta
	Dim CompLogColeta
	Dim NumeroColeta
	Dim BairroColeta
	Dim MunicipioColeta
	Dim EstadoColeta
	Dim isColetaDomiciliar
	Dim StatusCliente
	Dim IDPontoColeta
	Dim NomePontoColeta
	Dim CNPJPontoColeta
	Dim MotivoStatus
	Dim ContatoRespColeta
	Dim DDDContatoRespColeta
	Dim	TelefoneContatoRespColeta
	Dim TipoColeta

	Sub GetClientes()
		Dim sSql, arrClientes, intClientes, i
		sSql = "SELECT " & _
						"A.[idClientes], " & _
						"A.[razao_social], " & _
						"A.[nome_fantasia], " & _
						"A.[cnpj], " & _
						"A.[inscricao_estadual], " & _
						"A.[ddd], " & _
						"A.[telefone], " & _
						"A.[compl_endereco], " & _
						"A.[compl_endereco_coleta], " & _
						"A.[numero_endereco], " & _
						"A.[numero_endereco_coleta], " & _
						"A.[contato_respcoleta], " & _
						"A.[ddd_respcoleta], " & _
						"A.[telefone_respcoleta], " & _
						"A.[numero_sequencial], " & _
						"A.[data_atualizacao_sequencial], " & _
						"A.[typeColect], " & _
						"A.[status_cliente], " & _
						"A.[bonus_type], " & _
						"B.[idCategorias], " & _
						"B.[descricao] " & _
						"FROM [marketingoki2].[dbo].[Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Categorias] AS B " & _
						"ON A.[Categorias_idCategorias] = B.[idCategorias] "

		If Request.ServerVariables("HTTP_METHOD") = "GET" And Request.QueryString("find") Then
			sSql = sSql & GetWhereClientes()
'			response.write sSql
'			response.end()
		End If

		Call search(sSql, arrClientes, intClientes)

		If intClientes > -1 Then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = UBound(arrClientes, 2) + 1 'numero total de registros

			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1

			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)

			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1

			Response.Write "<tr><td colspan=10>"
			Response.Write PaginacaoExibir(intPag, intProdsPorPag, intClientes)
			Response.Write "</td></tr>"

			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar' align='center' width='3%'><img src='img/buscar.gif' class='imgexpandeinfo' alt='Administrar Cliente' onClick=""window.location.href = '"&URL&"adm/frmEditCadastroClienteLc.asp?showclient=true&idcliente="&arrClientes(0,i)&"';""/></td>"
					Response.Write "<td class='classColorRelPar'>"&arrClientes(1,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrClientes(2,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrClientes(20,i)&"</td>"
					'Response.Write "<td class='classColorRelPar'>"&arrClientes(22,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrClientes(3,i)&"</td>"
					If arrClientes(17,i) = 0 Then
						Response.Write "<td class='classColorRelPar'> --- </td>"
					ElseIf arrClientes(17,i) = 1 Then
						Response.Write "<td class='classColorRelPar'>Aprovado</td>"
					ElseIf arrClientes(17,i) = 2 Then
						Response.Write "<td class='classColorRelPar'>Rejeitado</td>"
					ElseIf arrClientes(17,i) = 3 Then
						Response.Write "<td class='classColorRelPar'>Inativo</td>"
					End If
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar' align='center' width='3%'><img src='img/buscar.gif' class='imgexpandeinfo' alt='Administrar Cliente' onClick=""window.location.href = '"&URL&"adm/frmEditCadastroClienteLc.asp?showclient=true&idcliente="&arrClientes(0,i)&"';""/></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrClientes(1,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrClientes(2,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrClientes(20,i)&"</td>"
					'Response.Write "<td class='classColorRelImpar'>"&arrClientes(22,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrClientes(3,i)&"</td>"
					If arrClientes(17,i) = 0 Then
						Response.Write "<td class='classColorRelImpar'> --- </td>"
					ElseIf arrClientes(17,i) = 1 Then
						Response.Write "<td class='classColorRelImpar'>Aprovado</td>"
					ElseIf arrClientes(17,i) = 2 Then
						Response.Write "<td class='classColorRelImpar'>Rejeitado</td>"
					ElseIf arrClientes(17,i) = 3 Then
						Response.Write "<td class='classColorRelImpar'>Inativo</td>"
					End If
					Response.Write "</tr>"
				End If
			Next
		Else
				Response.Write "<tr><td colspan='7' align='center' class='classColorRelPar'><b>Nenhum Cliente encontrado</b></td></tr>"
		End If
	End Sub

	Sub GetCategories()
		Dim sSql, arrCategories, intCategories, i
		Dim sSelected

		sSql = "SELECT idCategorias, descricao FROM Categorias"

		Call search(sSql, arrCategories, intCategories)

		For i=0 To intCategories
			If Categoria = arrCategories(0,i) Then
				sSelected = "selected"
			Else
				sSelected = ""
			End If
			Response.Write "<option value='"&arrCategories(0,i)&"' "&sSelected&">"&arrCategories(1,i)&"</option>"
		Next

	End Sub

	Sub GetGroups()
		Dim sSql, arrGroups, intGroups, i
		Dim sSelected

		sSql = "SELECT " & _
						"[idGrupos], " & _
						"[descricao] " & _
						"FROM [marketingoki2].[dbo].[Grupos]"
		Call search(sSql, arrGroups, intGroups)
		If intGroups > -1 Then
			For i=0 To intGroups
				If Grupo = arrGroups(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value='"&arrGroups(0,i)&"' "&sSelected&">"&arrGroups(1,i)&"</option>"
			Next
		End If
	End Sub

	Function GetWhereClientes()
		Dim sSqlWhere
		'------------------
		sSqlWhere = ""

		Dim Busca
		Dim Por
		Dim Status
		Dim Categorias
		Dim sStrConcat

		Dim bSetBusca
		Dim bSetStatus

		sStrConcat = " AND "

		Busca 		    = Request.QueryString("search")
		Por 	        = Request.QueryString("changetype")
		Status 		    = Request.QueryString("status")
		Categorias      = Request.QueryString("categoria")
		'statussol       = request.QueryString("grupos")
		'------------------

		If Len(Trim(Busca)) = 0 And CInt(Por) = -1  And CInt(Status) = -1 And CInt(Categorias) = -1 Then
			sSqlWhere = ""
		Else
			sSqlWhere = sSqlWhere & " WHERE "
		End If

		If Len(Busca) > 0 Then
			bSetBusca = True
			Select Case CInt(Por)
				Case 0
					sSqlWhere = sSqlWhere & " A.[cnpj] LIKE '"&Busca&"%' "
				Case 1
					sSqlWhere = sSqlWhere & " upper(A.[nome_fantasia]) LIKE upper('"&Busca&"%') "
				Case 2
					sSqlWhere = sSqlWhere & " upper(A.[razao_social]) LIKE upper('"&Busca&"%') "
				Case -1
					sSqlWhere = sSqlWhere & " A.[cnpj] LIKE '%"&Busca&"%' OR "
					sSqlWhere = sSqlWhere & " upper(A.[nome_fantasia]) LIKE upper('"&Busca&"%') OR "
					sSqlWhere = sSqlWhere & " upper(A.[razao_social]) LIKE upper('"&Busca&"%') "
			End Select
		End If
		If CInt(Status) > -1 Then
			bSetStatus = True
			If bSetBusca Then
				sSqlWhere = sSqlWhere & sStrConcat
			End If
			sSqlWhere = sSqlWhere & " A.[status_cliente] = " & Status
		End If
		If CInt(Categorias) > -1 Then
			If bSetStatus Then
				sSqlWhere = sSqlWhere & sStrConcat
			End If
			sSqlWhere = sSqlWhere & " B.[idCategorias] = " & Categorias
		End If

		'if cint(statussol) > -1 then
		'	If bSetStatus Then
		'		sSqlWhere = sSqlWhere & sStrConcat
		'	End If
		'	sSqlWhere = sSqlWhere & " C.[idGrupos] = " & statussol
		'end if

		GetWhereClientes = sSqlWhere

	End Function

	sub getGrupos()
		dim sql, arrgrupos, intgrupos, i
		sql = "select * from grupos"
		call search(sql, arrgrupos, intgrupos)
		if intgrupos > -1 then
			for i=0 to intgrupos
				response.write "<option value="""&arrgrupos(0,i)&""">"&arrgrupos(1,i)&"</option>"
			next
		end if
	end sub
%>
<html>
<head>
<script src="../adm/js/frmCadastroClienteAdmLc.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
        <style type="text/css">
        .rotulos {
            font-family:Verdana;
            font-size:10px;
            color:black;
            font-weight:normal;
        }
        </style>
</head>

<body <%If Request.QueryString("showclient") Then%>OnLoad="AdministraCliente()"<%End If%>>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="frmCadastroClienteAdm.asp" name="frmCadastroClienteAdm" method="POST">
		<input type="hidden" name="hiddenIDCliente" value="<%=Request.QueryString("idcliente")%>" />
		<input type="hidden" name="hiddenActionIsColetaDomiciliar" value="<%=isColetaDomiciliar%>" />
		<input type="hidden" name="hiddenActionManagerProve" value="" />
		<input type="hidden" name="hiddenActionForm" value="" />
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellspacing="3" cellpadding="2" width="100%" border=0>
						<tr>
							<td id="explaintitle" align="center" colspan="3">Cadastros de Clientes</td>
						</tr>
						<tr>
							<td align="right" colspan="3"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
							<tr id="findcontato" valign="baseline">
								<td align="right" valign="baseline">
									Pesquisa por:</td>
								<td align="left" valign="baseline" colspan="2">
									<select name="typeFindCliente" class="select">
										<option value="-1">Selecione</option>
										<option value="0">CNPJ</option>
										<option value="1">Nome Fantasia</option>
										<option value="2">Razão Social</option>
									</select></td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="right" valign="baseline">
									Busca:</td>
								<td align="left" valign="baseline">
                                    <input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtFindCliente" size="60" value=""/></td>
								<td align="left" class="rotulos">
                                    Busca aproximada pelas iniciais.<br />
                                    Para CNPJ utilize a máscara com pontos Ex: 99.999.999/9999-99</td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="right" valign="baseline">
									Categorias:</td>
								<td align="left" valign="baseline" colspan="2">
									<select name="cbCategoriasFindCliente" class="select">
										<option value="-1">Todas</option>
										<%Call GetCategories()%>
									</select></td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="right" valign="baseline">
									Status:</td>
								<td align="left" valign="baseline" colspan="2">
									<select name="cbStatusFindCliente" class="select">
										<option value="-1">Todos</option>
										<option value="0">Aguardando Aprovação</option>
										<option value="1">Aprovado</option>
										<option value="2">Rejeitado</option>
										<option value="3">Inativo</option>
									</select></td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline">
									&nbsp;</td>
								<td align="right" valign="baseline" colspan="2">
									<input type="button" class="btnform" name="btnFindCliente" value="Buscar" onClick="windowLocationFind('<%=URL%>')" /></td>
							</tr>
						<tr>
							<td colspan="3">
								<table id="tableGetClientesCadastro" cellpadding="1" cellspacing="1" width="100%">
									<tr>
										<th><img src="img/check.gif" /></th>
										<th>Razão Social</th>
										<th>Nome Fantasia</th>
										<th>Categoria</th>
										<th>CNPJ</th>
										<th>Status</th>
									</tr>
									<%Call GetClientes()%>
								</table>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</form>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
