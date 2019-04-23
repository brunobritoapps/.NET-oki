<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>

<%
    

' Informações Gerais
'------------------------	
	Dim NomeInfo
	Dim UsuarioInfo
	Dim SenhaInfo
	Dim EmailInfo
	Dim MasterInfo
	Dim StatusInfo
	Dim RazaoSocialInfo
	Dim CNPJInfo
	Dim CategoriaInfo
	Dim GrupoInfo
	dim status_cliente
' Checkagem
'------------------------	
	Dim	isCheckMaster
	Dim isCheckStatus
	Dim isCheckIdClient
' Request
'------------------------	
	Dim IDContato
	Dim Nome
	Dim Usuario
	Dim Senha
	Dim Email
	Dim Master
	Dim Status
	Dim Cliente
	Dim WhatDo

    '
	Call SubmitForm()

	Sub GetContatos()
		Dim sSql, arrContatos, intContatos, i
		Dim sSelected
		Dim Impar
		
		sSql = "SELECT " & _ 
						"A.[idContatos], " & _ 
						"A.[nome], " & _ 
						"A.[usuario], " & _ 
						"A.[senha], " & _ 
						"A.[email], " & _ 
						"A.[isMaster], " & _ 
						"A.[status_contato], " & _
						"B.[idClientes], " & _
						"B.[razao_social], " & _
						"B.[cnpj], " & _
						"C.[idCategorias], " & _
						"C.[descricao], " & _
						"D.[idGrupos], " & _
						"D.[descricao] " & _
						"FROM [marketingoki2].[dbo].[Contatos] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
						"ON A.[Clientes_idClientes] = B.[idClientes] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Categorias] AS C " & _
						"ON B.[Categorias_idCategorias] = C.[idCategorias] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Grupos] AS D " & _
						"ON B.[Grupos_idGrupos] = D.[idGrupos] " & _ 
                        "WHERE B.[idClientes] = " &  Request.QueryString("idcliente") 
		
		If Request.ServerVariables("HTTP_METHOD") = "GET" And Request.QueryString("find") Then
			sSql = sSql & GetWhereFindContato()
'			response.write sSql
'			response.end
		End If					

		Call search(sSql, arrContatos, intContatos)						

		With Response
			If intContatos > -1 Then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = UBound(arrContatos, 2) + 1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			.Write "<tr><td colspan=10>"
			.Write PaginacaoExibir(intPag, intProdsPorPag, intContatos)
			.Write "</td></tr>"
	
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
					If arrContatos(0,i) = CInt(Request.QueryString("IDContato")) Then
						sSelected = "checked"
					Else
						sSelected = ""
					End If
					
					If i Mod 2 = 0 Then
						.Write "<tr>"
						.Write "<td class='classColorRelPar'><input type='radio' name='radioIntIdContato' value='"&arrContatos(0,i)&"' onClick=window.location.href='frmContatosAdmlc.asp?idcliente="&Request.QueryString("idcliente") &"&IDContato="&arrContatos(0,i)&"' "&sSelected&"/></td>"
						.Write "<td class='classColorRelPar'>"&arrContatos(1,i)&"</td>" 'nome
						.Write "<td class='classColorRelPar'>"&arrContatos(2,i)&"</td>" 'usuário
						'.Write "<td class='classColorRelPar'>"&arrContatos(3,i)&"</td>" 'senha
						.Write "<td class='classColorRelPar'>"&arrContatos(4,i)&"</td>"
						If arrContatos(5,i) = 1 Then
							.Write "<td class='classColorRelPar'>Sim</td>"
						Else
							.Write "<td class='classColorRelPar'>Não</td>"
						End If
						If arrContatos(6,i) = 1 Then
							.Write "<td class='classColorRelPar'>Ativo</td>"
						Else
							.Write "<td class='classColorRelPar'>Inativo</td>"
						End If
						.Write "</tr>"
					Else	
						.Write "<tr>"
						.Write "<td class='classColorRelImpar'><input type='radio' name='radioIntIdContato' value='"&arrContatos(0,i)&"' onClick=window.location.href='frmContatosAdmlc.asp?IDContato="&arrContatos(0,i)&"' "&sSelected&"/></td>"
						.Write "<td class='classColorRelImpar'>"&arrContatos(1,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrContatos(2,i)&"</td>"
						'.Write "<td class='classColorRelImpar'>"&arrContatos(3,i)&"</td>" 'senha
						.Write "<td class='classColorRelImpar'>"&arrContatos(4,i)&"</td>"
						If arrContatos(5,i) = 1 Then
							.Write "<td class='classColorRelImpar'>Sim</td>"
						Else
							.Write "<td class='classColorRelImpar'>Não</td>"
						End If
						If arrContatos(6,i) = 1 Then
							.Write "<td class='classColorRelImpar'>Ativo</td>"
						Else
							.Write "<td class='classColorRelImpar'>Inativo</td>"
						End If
						.Write "</tr>"
					End If
				Next
			Else
				.Write "<td colspan='7' align='center' class='classColorRelPar'><b>Nenhum Contato cadastrado</b></td>"
			End If
		End With

	End Sub
	
	Sub GetListInfoContatos()

		Dim IDContato, arrInfoContatos, intInfoContatos, i
	
		sSql = "SELECT " & _ 
						"A.[idContatos], " & _ 
						"A.[nome], " & _ 
						"A.[usuario], " & _ 
						"A.[senha], " & _ 
						"A.[email], " & _ 
						"A.[isMaster], " & _ 
						"A.[status_contato], " & _
						"B.[idClientes], " & _
						"B.[razao_social], " & _
						"B.[cnpj], " & _
						"C.[idCategorias], " & _
						"C.[descricao], " & _
						"D.[idGrupos], " & _
						"D.[descricao], " & _
						"B.[status_cliente] " & _
						"FROM [marketingoki2].[dbo].[Contatos] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS B " & _
						"ON A.[Clientes_idClientes] = B.[idClientes] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Categorias] AS C " & _
						"ON B.[Categorias_idCategorias] = C.[idCategorias] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Grupos] AS D " & _
						"ON B.[Grupos_idGrupos] = D.[idGrupos] " & _
						"WHERE A.[idContatos] = " & Request.QueryString("IDContato")
						
		Call search(sSql, arrInfoContatos, intInfoContatos)				
		
		If intInfoContatos > -1 Then
			For i=0 To intInfoContatos
				NomeInfo = arrInfoContatos(1,i)
				UsuarioInfo = arrInfoContatos(2,i)
				SenhaInfo = "******" 'arrInfoContatos(3,i)
				EmailInfo = arrInfoContatos(4,i)
				isCheckMaster = arrInfoContatos(5,i)
				If arrInfoContatos(5,i) = 1 Then
					MasterInfo = "É Master"
				Else
					MasterInfo = "Não é Master"
				End If
				isCheckStatus = arrInfoContatos(6,i)
				If arrInfoContatos(6,i) = 1 Then
					StatusInfo = "Ativo"
				Else
					StatusInfo = "Inativo"
				End If
				RazaoSocialInfo = arrInfoContatos(8,i)
				CNPJInfo = arrInfoContatos(9,i)
				CategoriaInfo = arrInfoContatos(11,i)
				GrupoInfo = arrInfoContatos(13,i)
				status_cliente = arrInfoContatos(14,i)
				isCheckIdClient = arrInfoContatos(7,i)
			Next
		End If
	End Sub
	
	If Request.QueryString("IDContato") <> "" Then
		Call GetListInfoContatos()
	End If
	
	Sub GetClient()
		Dim sSql, arrClientes, intClientes, i
		Dim sSelected
		
		sSql = "SELECT idClientes, nome_fantasia FROM Clientes"
		
		Call search(sSql, arrClientes, intClientes)
		
		If intClientes > -1 Then
			For i=0 To intClientes
				If isCheckIdClient = arrClientes(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If	
				Response.Write "<option value='"&arrClientes(0,i)&"' "&sSelected&">"&arrClientes(1,i)&"</option>"
			Next
		Else
			Response.Write "<option value='-1'>Nenhum Cliente encontrado</option>"	
		End If
	End Sub
	
	Sub SubmitForm()
		dim valida_contato
		dim valida_usuario
		On Error Resume Next
		
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			valida_usuario = validaUsuario(Usuario)
			valida_contato = validaContato(Usuario,Senha)
			If WhatDo = "INSERT" Then
				if valida_usuario or valida_contato then
					response.write "<script>alert('Usuário ou Senha já cadastrado.')</script>"
				else
					Call InsertContato()
				end if				
			ElseIf WhatDo = "UPDATE" Then
				if valida_usuario or valida_contato then
					response.write "<script>alert('Usuário ou Senha já cadastrado.')</script>"
				else
					Call UpdateContato()
				end if				
			ElseIf WhatDo = "DELETE" Then
				Call DeleteContato()
			End If
		End If
		
		If Err <> 0 Then
			Response.Write Err.Description 
		End If
	End Sub
	
	Sub RequestForm()
		IDContato = Request.Form("hiddenIDContato")
		Nome = Request.Form("txtNomeContato")
		Usuario = Request.Form("txtUsuarioContato")
		Senha = Request.Form("txtSenhaContato")
		Email = Request.Form("txtEmailContato")
		Master = Request.Form("cbMasterContato")
		Status = Request.Form("cbStatusContato")
		Cliente = Request.Form("cbClienteContato")
		WhatDo = Request.Form("hiddenTypeAction")
	End Sub
	
	Sub InsertContato()
		Dim oCommand
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandText = "sp_AddContatoADM"
		oCommand.CommandType = 4 
		
		oCommand.Parameters("@IDContato") = 0
		oCommand.Parameters("@IDClient") = Cliente
		oCommand.Parameters("@Contato") = Nome
		oCommand.Parameters("@Usuario") = Usuario
		oCommand.Parameters("@Senha") = Senha
		oCommand.Parameters("@Email") = Email
		oCommand.Parameters("@isMaster") = Master
		oCommand.Parameters("@Status") = Status
		oCommand.Parameters("@WhatDo") = WhatDo
		
		oCommand.Execute()
		
		Set oCommand = Nothing
	End Sub
	
	Sub UpdateContato()
		Dim oCommand
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandText = "sp_AddContatoADM"
		oCommand.CommandType = 4 
		
		oCommand.Parameters("@IDContato") = IDContato
		oCommand.Parameters("@IDClient") = Cliente
		oCommand.Parameters("@Contato") = Nome
		oCommand.Parameters("@Usuario") = Usuario
		oCommand.Parameters("@Senha") = Senha
		oCommand.Parameters("@Email") = Email
		oCommand.Parameters("@isMaster") = Master
		oCommand.Parameters("@Status") = Status
		oCommand.Parameters("@WhatDo") = WhatDo
		
		oCommand.Execute()
		
		Set oCommand = Nothing
	End Sub
	
	Sub DeleteContato()
		Dim oCommand
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandText = "sp_AddContatoADM"
		oCommand.CommandType = 4 
		
		oCommand.Parameters("@IDContato") = IDContato
		oCommand.Parameters("@IDClient") = Cliente
		oCommand.Parameters("@Contato") = Nome
		oCommand.Parameters("@Usuario") = Usuario
		oCommand.Parameters("@Senha") = Senha
		oCommand.Parameters("@Email") = Email
		oCommand.Parameters("@isMaster") = Master
		oCommand.Parameters("@Status") = Status
		oCommand.Parameters("@WhatDo") = WhatDo
		
		oCommand.Execute()
		
		Set oCommand = Nothing
	End Sub
	
	Sub GetCategories()
		Dim sSql, arrCategories, intCategories, i
		
		sSql = "SELECT idCategorias, descricao FROM Categorias"
		
		Call search(sSql, arrCategories, intCategories)
		
		For i=0 To intCategories
			Response.Write "<option value='"&arrCategories(0,i)&"'>"&arrCategories(1,i)&"</option>"
		Next
		
	End Sub
	
	Function GetWhereFindContato()
		Dim sSqlWhere
		'------------------
		sSqlWhere = ""
		
		Dim Busca
		Dim Por
		Dim Status
		Dim Categorias
		Dim sStrConcat
		dim grupo
		
		Dim bSetBusca
		Dim bSetStatus
		
		sStrConcat = " AND "
		
		Busca 		 = Request.QueryString("search")
		Por 		 = Request.QueryString("changetype")
		Status 		 = Request.QueryString("status")
		Categorias 	 = Request.QueryString("categoria")
		grupo = request.querystring("grupo")
		'------------------
		
		If Len(Busca) = 0 And CInt(Status) = -1 And CInt(Categorias) = -1 and cint(grupo) = -1 Then
			sSqlWhere = ""
		Else
			sSqlWhere = sSqlWhere & " WHERE "
		End If
		
		If Len(Busca) > 0 Then
			bSetBusca = True
			Select Case CInt(Por)
				Case 0			
					sSqlWhere = sSqlWhere & " B.[cnpj] LIKE '%"&Busca&"%' "
				Case 1
					sSqlWhere = sSqlWhere & " A.[nome] LIKE '%"&Busca&"%' "
				Case 2
					sSqlWhere = sSqlWhere & " B.[razao_social] LIKE '%"&Busca&"%' "	
				Case -1	
					sSqlWhere = sSqlWhere & " (A.[nome] LIKE '%"&Busca&"%' AND "
					sSqlWhere = sSqlWhere & " B.[razao_social] LIKE '%"&Busca&"%') OR "
					sSqlWhere = sSqlWhere & " B.[cnpj] LIKE '%"&Busca&"%' "
			End Select	
		End If
		If CInt(Status) > -1 Then
			bSetStatus = True
			If bSetBusca Then
				sSqlWhere = sSqlWhere & sStrConcat
			End If
			sSqlWhere = sSqlWhere & " A.[status_contato] = " & Status
		End If
		If CInt(Categorias) > -1 Then
			If bSetStatus Then
				sSqlWhere = sSqlWhere & sStrConcat
			End If
			sSqlWhere = sSqlWhere & " C.[idCategorias] = " & Categorias
		End If
		If CInt(grupo) > -1 Then
			If bSetStatus Then
				sSqlWhere = sSqlWhere & sStrConcat
			End If
			sSqlWhere = sSqlWhere & " D.[idGrupos] = " & grupo
		End If
		
		GetWhereFindContato = sSqlWhere
		
	End Function
	
	function getGrupos() 
		dim sql, grupos, intgrupos, i
		dim ret
		ret = ""
		
		sql = "select idgrupos, descricao from grupos"
		call search(sql, grupos, intgrupos)
		if intgrupos > -1 then
			for i=0 to intgrupos
				ret = ret & "<option value="""&grupos(0,i)&""">"&grupos(1,i)&"</option>"		
			next
		end if
		
		getGrupos = ret
	end function
	
	function validaContato(usuario, senha)
		dim sql, arr, intarr, i
		dim sql2, arr2, intarr2, i2
		if WhatDo = "UPDATE" then
			if (request.form("hidden_usuario_info") <> usuario) and (request.form("hidden_senha_info") <> senha) then
				sql = "SELECT [idContatos] " & _
							  ",[Clientes_idClientes] " & _
							  ",[nome] " & _
							  ",[usuario] " & _
							  ",[senha] " & _
							  ",[email] " & _
							  ",[isMaster] " & _
							  ",[status_contato] " & _
						  "FROM [marketingoki2].[dbo].[Contatos] " & _
						"where usuario = '"&usuario&"' and senha = '"&senha&"'"
				call search(sql, arr, intarr)
				if intarr > -1 then
					validaContato = true
				else
					sql2 = "SELECT [idContatos] " & _
							  ",[Clientes_idClientes] " & _
							  ",[nome] " & _
							  ",[usuario] " & _
							  ",[senha] " & _
							  ",[email] " & _
							  ",[isMaster] " & _
							  ",[status_contato] " & _
						  "FROM [marketingoki2].[dbo].[Contatos] " & _
						"where senha = '"&senha&"'"
					call search(sql2, arr2, intarr2)
					if intarr2 > -1 then
						validaContato = true
					else
						validaContato = false
					end if
				end if				
			end if
		else
			sql = "SELECT [idContatos] " & _
						  ",[Clientes_idClientes] " & _
						  ",[nome] " & _
						  ",[usuario] " & _
						  ",[senha] " & _
						  ",[email] " & _
						  ",[isMaster] " & _
						  ",[status_contato] " & _
					  "FROM [marketingoki2].[dbo].[Contatos] " & _
					"where usuario = '"&usuario&"' and senha = '"&senha&"'"
			call search(sql, arr, intarr)
			if intarr > -1 then
				validaContato = true
			else
				sql2 = "SELECT [idContatos] " & _
						  ",[Clientes_idClientes] " & _
						  ",[nome] " & _
						  ",[usuario] " & _
						  ",[senha] " & _
						  ",[email] " & _
						  ",[isMaster] " & _
						  ",[status_contato] " & _
					  "FROM [marketingoki2].[dbo].[Contatos] " & _
					"where senha = '"&senha&"'"
				call search(sql2, arr2, intarr2)
				if intarr2 > -1 then
					validaContato = true
				else
					validaContato = false
				end if
			end if				
		end if		
	end function
	
	function validaUsuario(usuario)
		dim sql, arr, intarr, i
		if WhatDo = "UPDATE" then
			if request.form("hidden_usuario_info") <> usuario then
				sql = "SELECT [idContatos] " & _
							  ",[Clientes_idClientes] " & _
							  ",[nome] " & _
							  ",[usuario] " & _
							  ",[senha] " & _
							  ",[email] " & _
							  ",[isMaster] " & _
							  ",[status_contato] " & _
						  "FROM [marketingoki2].[dbo].[Contatos] " & _
						"where usuario = '"&usuario&"'"
				call search(sql, arr, intarr)
				if intarr > -1 then
					validaUsuario = true
				else
					validaUsuario = false
				end if	
			end if
		else
			sql = "SELECT [idContatos] " & _
						  ",[Clientes_idClientes] " & _
						  ",[nome] " & _
						  ",[usuario] " & _
						  ",[senha] " & _
						  ",[email] " & _
						  ",[isMaster] " & _
						  ",[status_contato] " & _
					  "FROM [marketingoki2].[dbo].[Contatos] " & _
					"where usuario = '"&usuario&"'"
			call search(sql, arr, intarr)
			if intarr > -1 then
				validaUsuario = true
			else
				validaUsuario = false
			end if	
		end if							
	end function
	
%>


<head>
    <script src="js/frmContatosAdm.js" language="javascript"></script>
    <link rel="stylesheet" type="text/css" href="../css/geral.css">
    <title><%=TITLE%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<!DOCTYPE html>
<body>

<div class="container-novo" style="width:775px;">
	<div id="conteudo-novo">
		<form action="frmContatosAdmlc.asp" name="frmContatosAdmlc" method="POST">
			
            <input type="hidden" name="hiddenIDContato" value="<%=Request.QueryString("IDContato")%>" />
			<input type="hidden" name="hiddenTypeAction" value="" />
			<input type="hidden" name="hidden_usuario_info" value="<%= UsuarioInfo %>" />
			<input type="hidden" name="hidden_senha_info" value="<%= SenhaInfo %>" />

            <table style="width:775px;">
			<tr>
                <td width="11" background="img/Bg_LatEsq.gif" >&nbsp;</td> 
				<td id="conteudo" >
                    <table style="border-collapse: separate; border-spacing: 0px; width:100%;">
                        <tr>
                            <td colspan="3" id="explaintitle" align="center">Cadastro / Alteração de Usuários</td>
                        </tr>
                        <tr>
                            <td>&nbsp;</td>
                            <!--<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>-->
                        </tr>
                        <tr>
                        <td valign="top" colspan="2">
                        <table id="tableContatosAdm"  style="border-collapse: separate; border-spacing: 3px; width:100%;">
                            <tr>
                                <td align="left" width="8%">Nome:</td>
                                <td align="left" colspan="4" ><input name="txtNomeContato" class="textreadonly" type="text" value="<%=NomeInfo%>" size="40" maxlength="40"></td>
                            </tr>
                            <tr>
							<td align="left" width="8%">Usuário:</td>
							<td align="left" colspan="4" ><input name="txtUsuarioContato" class="textreadonly" type="text" value="<%=UsuarioInfo%>" size="20" maxlength="20" /></td>
							</tr>
							<tr>
							<td align="left" width="8%">Senha:</td>
							<td align="left" colspan="4" ><input name="txtSenhaContato" class="textreadonly" type="password" value="<%=SenhaInfo%>" size="20" maxlength="20" /></td>
							</tr>
							<tr>
							<td align="left" width="8%">Email:</td>
							<td align="left" colspan="4" ><input name="txtEmailContato" class="textreadonly" type="text" value="<%=EmailInfo%>" size="30" maxlength="255"></td>
							</tr>
							<tr>
							<td align="left" width="8%">Cliente:</td>
							<td align="left" colspan="4" >
							<select name="cbClienteContato" class="select" style="width:250px;">
							<option value="-1">Selecione um Cliente</option>
							<%Call GetClient()%>
							</select>
							</td>
							</tr>
							<tr>
							<td align="left" width="8%">Master:</td>
							<td align="left" colspan="4" >
							<select name="cbMasterContato" class="select">
							<option value="0" <%If isCheckMaster = 0 Then%>selected<%End If%>>Não</option>
							<option value="1" <%If isCheckMaster = 1 Then%>selected<%End If%>>Sim</option>
												</select>
											</td>
										</tr>
										<tr>
											<td align="left" width="8%">Ativo:</td>
											<td align="left" colspan="4">
												<select name="cbStatusContato" class="select">
													<option value="0" <%If isCheckStatus = 0 Then%>selected<%End If%>>Não</option>
													<option value="1" <%If isCheckStatus = 1 Then%>selected<%End If%>>Sim</option>
												</select>
											</td>
										</tr>
										<% If cint(status_cliente) = 1 Then %>
										<tr>
											<td>
												<%If Request.QueryString("IDContato") <> "" Then%>	
													&nbsp;<%Else%>	
													<%End If%>
											</td>
											<td>
													<input name="btnEditarContato" class="btnform" type="button" value="Editar" onClick="setActionForm('UPDATE')" /></td>
											<td>
													<input name="btnSalvarContato" class="btnform" type="button" value="Salvar" onClick="setActionForm('INSERT')" /></td>
											<td>
												<input name="btnLimpar" class="btnform" type="reset" value="Limpar" /></td>
											<td>
												&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;</td>
										</tr>
										<% End If %>
									</table>
								</td>
								<td>
									<table cellpadding="2" cellspacing="2" width="400" id="tableRelContatosAdm" align="left">
										<tr>
											<th colspan="2" id="explaintitle">Informações Gerais</th>
										</tr>
										<tr>
											<td width="20%"><b>Nome:</b></td>
											<td><%=NomeInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Usuário:</b></td>
											<td><%=UsuarioInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Senha:</b></td>
											<td><%=SenhaInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Email:</b></td>
											<td><%=EmailInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Master:</b></td>
											<td><%=MasterInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Status:</b></td>
											<td><%=StatusInfo%></td>
										</tr>
										<tr>
											<th colspan="2" id="explaintitle">Contato do Cliente</th>
										</tr>
										<tr>
											<td width="20%"><b>Razão Social:</b></td>
											<td><%=RazaoSocialInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>CNPJ:</b></td>
											<td><%=CNPJInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Categoria:</b></td>
											<td><%=CategoriaInfo%></td>
										</tr>
										<tr>
											<td width="20%"><b>Grupo:</b></td>
											<td><%=GrupoInfo%></td>
										</tr>
									</table>
								</td>
							</tr>
                            </table>

                    <table style="border-collapse: separate; border-spacing: 3px; width:100%;">
                        <tr>
                            <td colspan="3" id="explaintitle" align="center">Pesquisa:</td>
                        </tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline">
									Por:</td>
								<td align="left" valign="baseline">
									<select name="typeFindContato" class="select">
										<option value="-1">Selecione</option>
										<option value="0">CNPJ</option>
										<option value="1">Contato</option>
										<option value="2">Razão Social</option>
									</select></td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline">
									Busca:</td>
								<td align="left" valign="baseline">
                                    <input type="text" class="textreadonly" style="width:400px;" name="txtFindContato" value="" size="90" /></td>
								<td align="left" valign="baseline">
                                    &nbsp;</td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline">
									Status:</td>
								<td align="left" valign="baseline">
									<select name="cbStatusFindContato" class="select">
										<option value="-1">Todos</option>
										<option value="1">Ativo</option>
										<option value="0">Inativo</option>
									</select></td>
								<td align="left" valign="baseline">
									&nbsp;</td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline">
									Categorias:&nbsp;									
								</td>
								<td align="left" valign="baseline">
									<select name="cbCategoriasFindContato" class="select">
										<option value="-1">Todas</option>
										<%Call GetCategories()%>
									</select></td>
								<td align="left" valign="baseline">
									&nbsp;</td>
							</tr>
							<tr id="findcontato" valign="baseline">	
								<td align="left" valign="baseline">
									Grupos:</td>
								<td align="left" valign="baseline">
									<select name="cbGruposCliente" class="select">
										<option value="-1">Todos</option>
										<%=getGrupos()%>
									</select></td>
								<td align="left" valign="baseline">
									&nbsp;</td>
							</tr>
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline" colspan="3">
									<input align="right" type="button" class="btnform" name="btnFindContato" value="Buscar" onClick="windowLocationFind('<%=URL%>')" />
								</td>
							</tr>
							<tr>
								<td colspan="3">
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelContatosAdm">
										<tr>
											<th width="3%"><img src='img/check.gif' alt='Selecionar' /></th>
											<th width="30%">Nome</th>
											<th>Usuário</th>
											<th width="20%">Email</th>
											<th width="8%">Master</th>
											<th width="8%">Status</th>
										</tr>
										<%Call GetContatos()%>
										<tr>
											<th colspan="7" height="15"></th>
										</tr>
									</table>
								</td>
							</tr>
						</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		
    		</table>
        </form>
	</div>
</div>
</body>

<%Call close()%>
