<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Dim User
	Dim Password
	Dim Message

	Function AutenthicateUserAdm()
		Dim sSql, arrUsuarioAdm, intUsuarioAdm, i

		sSql = "SELECT " & _ 
						"[idAdministrator], " & _ 
						"[nome] " & _ 
						"FROM [marketingoki2].[dbo].[Administrator] " & _
						"WHERE [login] = '"&User&"' AND " & _
						"[password] = '"&Password&"' " & _
						"AND [status] = 1"
	
	Call search(sSql, arrUsuarioAdm, intUsuarioAdm)

	If intUsuarioAdm > -1 Then
		For i=0 To intUsuarioAdm
			Session("IDAdministrator") = arrUsuarioAdm(0,i)
			Session("Nome") = arrUsuarioAdm(1,i)
			Session("LoginADM") = User
			Session("PasswordADM") = Password
		Next
		AutenthicateUserAdm = True
	Else
		AutenthicateUserAdm = False	
	End If
	End Function

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			If AutenthicateUserAdm()  Then
				Response.Redirect "frmOperacionalAdm.asp"
			Else
				Message = "Usuário ou senha inválidos!"
			End If
		Else
			Message = Request.QueryString("msg")	
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
<link rel="stylesheet" type="text/css" href="../css/geral.css">
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
					<form action="frmLoginAdm.asp" name="frmLoginAdm" method="POST">
						<table cellpadding="3" cellspacing="4" width="100%" id="tableLoginCliente">
							<tr>
								<td colspan="2" id="explaintitle" align="center">Login da Administração</td>
							</tr>
							<tr>
								<td align="right" width="25%">Usuário:</td>
								<td align="left"><input type="text" class="text" name="txtLogin" value="" size="20" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Senha:</td>
								<td align="left"><input type="password" class="text" name="txtSenha" value="" size="20" /></td>
							</tr>
							<tr>
								<td align="right"><input type="submit" class="btnform" name="btnSubmitLogin" value="Login" /></td>
								<td align="left"><input type="reset" class="btnform" name="btnLimpaForm" value="Limpar" /></td>
							</tr>
							<%If Message <> "" Then%>
								<tr>
									<td colspan="2" style="color:#FF6A6A;font-weight:bold;" align="center"><%=Message%></td>
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
