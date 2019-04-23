<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%
	Dim Message
	Dim User
	Dim Password

	If Request.QueryString("IDMessage") <> "" Then
		Message = Request.QueryString("IDMessage")
	End If

	Function AuthenticateUser(ByVal User, ByVal Password)
		Dim sSql, arrUser, intUser, i
		Dim sSql2, arrPontoCliente, intPontoCliente
		Dim sSql3, arrNumeroSequencial, intNumeroSequencial

		sSql = "SELECT " & _
							"[A].[idContatos], " & _ 
							"[A].[Clientes_idClientes], " & _ 
							"[A].[nome], " & _ 
							"[A].[email], " & _ 
							"[A].[isMaster], " & _
							"[B].[typeColect], " & _ 
							"[B].[cod_cli_consolidador], " & _ 
							"[B].[cod_bonus_cli], " & _
							"[B].[inscricao_estadual] " & _
							"FROM [marketingoki2].[dbo].[Contatos] AS [A] " & _
							"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [B] " & _
							"ON [A].[Clientes_idClientes] = [B].[idClientes] " & _
							"WHERE [A].[status_contato] = 1 AND [B].[status_cliente] = 1 " & _ 
							"AND [A].[usuario] = '"& User &"' " & _ 
							"AND [A].[senha] = '"& Password &"'"

		Call search(sSql, arrUser, intUser)
		If intUser > -1 Then
			For i=0 To intUser
				Session("User") = User
				Session("Password") = Password
				Session("IDContato") = arrUser(0,i)
				Session("IDCliente") = arrUser(1,i)
				Session("NomeContato") = arrUser(2,i)
				Session("Email") = arrUser(3,i)
				Session("isMaster") = arrUser(4,i)
				Session("isColetaDomiciliar") = arrUser(5,i)
				session("cod_cli_consolidador") = arrUser(6,i)
				session("cod_bonus") = arrUser(7,i)
				session("IE") = arrUser(8,i)

				sSql3 = "SELECT numero_sequencial, data_atualizacao_sequencial FROM Clientes WHERE idClientes = " & Session("IDCliente")
				Call search(sSql3, arrNumeroSequencial, intNumeroSequencial)
				If intNumeroSequencial > -1 Then
					Session("NumeroSequencial") = arrNumeroSequencial(0,0)
					Session("DataSequencial") = arrNumeroSequencial(1,0)
				End If

				If Session("isColetaDomiciliar") = 0 Then
					sSql2 = "SELECT * FROM Solicitacao_coleta_has_Clientes WHERE Clientes_idClientes = " & Session("IDCliente")
					Call search(sSql2, arrPontoCliente, intPontoCliente)
					If intPontoCliente < -1 Then
						Session("IDPontoColeta") = arrPontoCliente(1,0)
					End If

				End If
				AuthenticateUser = True
			Next
		Else
			AuthenticateUser = False
		End If
	End Function

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Dim RetAuthenticate
			Call RequestForm()
			RetAuthenticate = AuthenticateUser(User, Password)
			If RetAuthenticate Then
				Response.Redirect "frmOperacionalCliente.asp"
			Else
				Message = "Usuário ou Senha inválidos!"
			End If
		ElseIf Request.ServerVariables("HTTP_METHOD") = "GET" Then
			If Request.QueryString("IDMessage") <> "" Then
				Message = Request.QueryString("IDMessage")
			End If
		End If
	End Sub
	
	Sub RequestForm()
		User = Request.Form("txtLogin")		
		Password = Request.Form("txtSenha")
	End Sub

	Call SubmitForm()
%>
<html>
<head>
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
					<form action="frmLoginCliente.asp" name="frmLoginCliente" method="POST">
						<table cellpadding="3" cellspacing="4" width="100%" id="tableLoginCliente">
							<tr>
								<td colspan="2" id="explaintitle">Login</td>
							</tr>
							<tr>
								<td align="right" width="49%">Usuário:</td>
								<td width="51%" align="left"><input type="text" class="text" name="txtLogin" value="" size="20" /></td>
							</tr>
							<tr>
								<td align="right" width="49%">Senha:</td>
								<td align="left"><input type="password" class="text" name="txtSenha" value="" size="20" /></td>
							</tr>
							<tr>
								<td align="right"><input type="submit" class="btnform" name="btnSubmitLogin" value="Login" /></td>
								<td align="left"><input type="reset" class="btnform" name="btnLimpaForm" value="Limpar" /></td>
							</tr>
							<%If Message <> "" Then%>
								<tr align="center">
									<td colspan="2" style="color:#FF6A6A;font-weight:bold;"><%=Message%></td>
								</tr>
							<%End If%>
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
