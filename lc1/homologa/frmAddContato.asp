<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%
	If Session("isMaster") <> 1 Then
		Response.Redirect "frmOperacionalCliente.asp"
	End If

	Dim ID
	Dim Contato
	Dim Usuario
	Dim Senha
	Dim Email
	Dim Tipo
	Dim Status
	Dim Departamento
    Dim Telefone
    Dim DDD
    Dim Ramal

	Sub SubmitForm()
	
		Dim oCommand
		Dim sSqlQueryUpdate, arrContatoQueryUpdate, intContatoQueryUpdate, iQueryUpdate
		dim valida_contato
		dim valida_usuario
		Set oCommand = Server.CreateObject("ADODB.Command")
		
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			
			Call RequestForm()
			
			valida_usuario = validaUsuario(Usuario)
			valida_contato = validaContato(Usuario,Senha)
			if valida_usuario or valida_contato then
				response.write "<script>alert('Usuário ou senha já cadastrado.')</script>"
			else
				If Request.Form("hiddenActionForm") = "UPDATE" Then
					Call updateContato()
				Else
					oCommand.CommandTimeout = 200
					oCommand.ActiveConnection = oConn
					oCommand.CommandType = 4
					oCommand.CommandText = "sp_AddContatoLc"
					oCommand.Parameters("@IDClient") = Session("IDCliente")
					oCommand.Parameters("@Contato") = Contato
					oCommand.Parameters("@Usuario") = Usuario
					oCommand.Parameters("@Senha") = Senha
					oCommand.Parameters("@Email") = Email
					oCommand.Parameters("@Departamento") = Departamento
					oCommand.Parameters("@isMaster") = Tipo
                    oCommand.Parameters("@Telefone") = Telefone
                    oCommand.Parameters("@Ramal") = Ramal
                    oCommand.Parameters("@DDD") = DDD
					oCommand.Execute()
					Contato = ""
					Usuario = ""
					Senha = ""
					Email = ""
					Departamento = ""
                    Telefone = ""
                    Ramal = ""
					Tipo = 0
					Response.Write "<script>alert('Contato inserido com sucesso!');</script>"
				Set oCommand = Nothing
				End If
			end if
		
		Else
			If Request.QueryString("Query") = "UPDATE" Then
				sSqlQueryUpdate = "SELECT idContatos, Clientes_idClientes, nome, usuario, senha, email, isMaster, status_contato, departamento, telefone, ramal,ddd FROM Contatos WHERE idContatos = " & Request.QueryString("ID")
				Call search(sSqlQueryUpdate, arrContatoQueryUpdate, intContatoQueryUpdate)
				If intContatoQueryUpdate > -1 Then
					For iQueryUpdate=0 To intContatoQueryUpdate
						ID = arrContatoQueryUpdate(0,iQueryUpdate)
						Contato = arrContatoQueryUpdate(2,iQueryUpdate)
						Usuario = arrContatoQueryUpdate(3,iQueryUpdate)
						Senha = arrContatoQueryUpdate(4,iQueryUpdate)
						Email = arrContatoQueryUpdate(5,iQueryUpdate)
						Tipo =  arrContatoQueryUpdate(6,iQueryUpdate)
						Status = arrContatoQueryUpdate(7,iQueryUpdate)
						Departamento = arrContatoQueryUpdate(8,iQueryUpdate)
                        
                        Telefone = arrContatoQueryUpdate(9,iQueryUpdate)
                        Ramal = arrContatoQueryUpdate(10,iQueryUpdate)
                        DDD = arrContatoQueryUpdate(11,iQueryUpdate)
					Next
				Else
					Response.Write "<script>alert('Não foi possível encontrar o Contato especificado!')</script>"
					Response.Redirect "frmAddContato.asp"
				End If
			ElseIf Request.QueryString("Query") = "DELETE" Then
				Call deleteContato()
			End If
		End If
	End Sub

	Sub RequestForm()
		Contato = Request.Form("txtContatoColeta")
		Usuario = Request.Form("txtUsuario")
		Senha = Request.Form("txtSenha")
		Email = Request.Form("txtEmail")
		Departamento = Request.Form("txtDepartamento")
        Telefone = Request.Form("txtTelefone")
        Ramal = Request.Form("txtRamal")
        DDD = Request.Form("txtDDD")
		If Session("isMaster") = 0 Then
			Tipo = Request.Form("hiddenIsMaster")
		Else
			Tipo = Request.Form("radioIsMaster")
		End If
		Status = Request.Form("radioStatus")
	End Sub

	Sub geraContatos()
	
		Dim sSql, arrContatos, intContatos, i

		sSql = "SELECT nome, usuario, senha, email, departamento, isMaster, status_contato, idContatos FROM Contatos WHERE Clientes_idClientes = " & Session("IDCliente")
		Call search(sSql, arrContatos, intContatos)
		With Response
			If intContatos > -1 Then
				.Write "<input type='hidden' name='hiddenIntContatos' value='"&intContatos&"' />"
				For i=0 To intContatos
					If i Mod 2 = 0 Then
						.Write "<tr>"
						'.Write "<td align='center' class='classColorRelPar'><input type='button' value='Editar' name='checkContato' id='checkContato"&i&"' value='" & arrContatos(0,i) & "' /></td>"
						.Write "<td align='center' class='classColorRelPar'><input type='button' onclick='lcEditContato("&arrContatos(7,i)&");' value='Editar' name='checkContato' id='"&arrContatos(7,i)&"' " & "/></td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(7,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(0,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(1,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(3,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(4,i) & "</td>"
						If arrContatos(6,i) = 1 Then
						.Write "<td class='classColorRelPar'>Ativo</td>"
						Else
						.Write "<td class='classColorRelPar'>Inativo</td>"
						End If
						.Write "</tr>"
					Else
						.Write "<tr>"
						'.Write "<td align='center' class='classColorRelPar'><input type='button' value='Editar' name='checkContato' id='checkContato"&i&"' value='" & arrContatos(0,i) & "' /></td>"
						.Write "<td align='center' class='classColorRelPar'><input type='button' onclick='lcEditContato("&arrContatos(7,i)&");' value='Editar' name='checkContato' id='"&arrContatos(7,i)&"' " & "/></td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(7,i) & "</td>"						
						.Write "<td class='classColorRelPar'>" & arrContatos(0,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(1,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(3,i) & "</td>"
						.Write "<td class='classColorRelPar'>" & arrContatos(4,i) & "</td>"
						If arrContatos(6,i) = 1 Then
						.Write "<td class='classColorRelPar'>Ativo</td>"
						Else
						.Write "<td class='classColorRelPar'>Inativo</td>"
						End If
						.Write "</tr>"
					End If
				Next
			Else
				.Write "<td colspan='4'>Não existem Contatos cadastrados.</td>"
			End If
		End With
	End Sub

	Sub updateContato()
		Dim sSql
		sSql = "UPDATE Contatos SET nome='" & Contato & "', usuario='" & Usuario & "', senha='" & Senha & "', email='" & Email & "', isMaster=" & Tipo & ", status_contato=" & Status & ", departamento='" & Departamento & "' "
        sSql = sSql & ",telefone = " & Telefone & ", ramal = '" & Ramal & "', ddd = " & DDD & " "
        sSql = sSql &  " WHERE idContatos = " & Request.Form("hiddenIdContato")
        response.Write sSql
		Call exec(sSql)
		Contato = ""
		Usuario = ""
		Senha = ""
		Email = ""
		Tipo = 0
		Status = 0
		Response.Write "<script>alert('Contato atualizado com sucesso!');</script>"
	End Sub

	Sub deleteContato()
		Dim sSql

		sSql = "DELETE FROM Contatos WHERE idContatos = " & Request.QueryString("ID")
		Call exec(sSql)
		Contato = ""
		Usuario = ""
		Senha = ""
		Email = ""
		Tipo = 0
		Status = 0
		Response.Write "<script>alert('Contato deletado com sucesso!');</script>"
	End Sub

	function validaContato(usuario, senha)
		dim sql, arr, intarr, i
		dim sql2, arr2, intarr2, i2
		if Request.Form("hiddenActionForm") = "UPDATE" then
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
		if Request.Form("hiddenActionForm") = "UPDATE" then
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

	Call SubmitForm()
%>
<html>
<head>
<script language="javascript" type="text/javascript" src="js/frmAddContato.js"></script>
<script language="javascript" type="text/javascript" src="js/frmAddContatoLc.js"></script>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<form action="frmAddContato.asp" name="frmAddContato" method="POST">
					<input type="hidden" name="hiddenActionForm" value="<%=Request.QueryString("Query")%>" />
					<input type="hidden" name="hiddenIdContato" value="<%=ID%>" />
					<input name="hidden_usuario_info" type="hidden" value="<%=Usuario%>" />
					<input name="hidden_senha_info" type="hidden" value="<%=Senha%>" />
					<%If Session("isMaster") = 0 Then%>
						<input type="hidden" name="hiddenIsMaster" value="0" />
					<%End If%>
					<table cellpadding="1" cellspacing="3" width="100%" id="tableCadClienteContato" border="0">
						<tr>
							<td colspan="3" id="explaintitle" align="center">Manutenção de Contatos</td>
						</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="3" align="left" width="100%"><b id="fontred">Atenção :</b>
										<b style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-size: 13px; vertical-align: baseline; background-color: transparent; color: rgb(55, 61, 69); font-family: Arial, sans-serif; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: 14px; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Os campos com (asterisco)* são de preenchimento obrigatório.</b></td>
						</tr>
						<tr>
							<td colspan="3" align="left" width="100%">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="3" align="left" width="100%">Para incluir um novo contato, preencha todos os campos abaixo e depois clique em SALVAR.</td>
						</tr>
						<tr>
							<td colspan="3" align="left" width="100%">&nbsp;</td>
						</tr>						
						<tr>
							<td align="right" width="25%">Contato*</td>
							<td align="left" colspan="2"><input type="text" class="text" name="txtContatoColeta" value="<%=Contato%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Usuário*</td>
							<td align="left" colspan="2"><input type="text" class="text" name="txtUsuario" value="<%=Usuario%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Senha*</td>
							<td align="left" colspan="2"><input type="password" class="text" name="txtSenha" value="<%=Senha%>" size="20" maxlength="20" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">E-mail*</td>
							<td align="left" colspan="2"><input type="text" class="text" name="txtEmail" value="<%=Email%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Departamento*</td>
							<td align="left" colspan="2"><input type="text" class="text" name="txtDepartamento" value="<%=Departamento%>" size="50" maxlength="50"/></td>
						</tr>
						<%If Session("isMaster") = 1 Then%>
							<tr>
								<td align="right" width="25%">DDD*</td>
								<td>
									<input type="text" class="text" name="txtDDD" value="<%=DDD%>" size="10" maxlength="2" /></td>
								<td>
									Preencher sem o zero a esquerda Ex: 11</td>
							</tr>
							<tr>
								<td align="right" width="25%">Telefone*</td>
								<td>
									<input type="text" class="text" name="txtTelefone" value="<%=Telefone%>" size="20" maxlength="9" /></td>
								<td>
									Digite apenas números Ex:&nbsp; 999999999</td>
							</tr>
							<tr>
								<td align="right" width="25%">Ramal</td>
								<td colspan="2">
									<input type="text" class="text" name="txtRamal" value="<%=Ramal%>" size="10" maxlength="10" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Master:</td>
								<td colspan="2">
									<input type="radio" name="radioIsMaster" value="1" <%If Tipo = 1 Then%>checked="checked"<%End If%> /> Sim
									<input type="radio" name="radioIsMaster" value="0" <%If Tipo = 0 Then%>checked="checked"<%End If%> /> Não
								</td>
							</tr>
						<%Else%>
							<tr>
								<td align="right" width="25%">Master:</td>
								<td colspan="2">
									<input type="radio" name="radioIsMaster" value="1" disabled="true" <%If Tipo = 1 Then%>checked="checked"<%End If%> /> Sim
									<input type="radio" name="radioIsMaster" value="0" disabled="true" checked="checked" <%If Tipo = 0 Then%>checked="checked"<%End If%> /> Não
								</td>
							</tr>
						<%End If%>
						
						<div <%If Request.QueryString("ID") <> "" Then%>style="display:block;"<%Else%>style="display:none;"<%End If%>>
							<tr>
							<td align="right" width="25%">Situação:</td>
							<td width="75%" colspan="2">
								<input type="radio" name="radioStatus" value="1" <%If Status = 1 Then%>checked="checked"<%End If%> /> Ativo
								<input type="radio" name="radioStatus" value="0"<%If Status = 0 Then%>checked="checked"<%End If%> /> Inativo
							</td>
							</tr>
						</div>
						
						<tr>
							<td colspan="1" width="25%">&nbsp;</td>
							<td align="right" colspan="2">
								<input type="reset" class="btnform" name="btnReset" value="Limpar" />&nbsp;
								<input type="button" class="btnform" name="btnSubmit" value="Salvar" onClick="validaCadClienteContato()" />
							</td>
							<!--<td align="left" colspan="2"><input type="button" class="btnform" name="btnSubmit" value="Salvar" onClick="validaCadClienteContato()" /></td>-->
						</tr>
						
						<tr>
							<td colspan="3" id="explaintitle" align="center">
								Lista de Contatos Cadastrados
								<!--<If Session("isMaster") = 1 Then%>
									<span style="margin-left:280px;">
										<select  name="cbActionContatos" class="select" onChange="redirActionContato()">
											<option value="0"> ---- </option>
											<option value="1">Atualizar</option>
										</select>
										&nbsp Selecionado
									</span>
								<End If%>-->

							</td>
						</tr>
						<tr>
							<td colspan="3">
								<table cellpadding="3" cellspacing="3" width="100%" align="center" id="tableRelContatos">
									<tr>
										<th><!--<img src='img/check.gif' alt='Selecionar' /></th>--></th>
										<th>Id</th>
										<th>Nome</th>
										<th>Usuário</th>
										<th>e-mail</th>
										<th>Departamento</th>
										<th>Situação</th>
									</tr>
									<%Call geraContatos()%>
									<tr>
										<th colspan="7"></th>
									</tr>
								</table>
							</td>
						</tr>
					</table>
					</form>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
