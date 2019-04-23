<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Dim User
	Dim Password
	Dim Message

	Function AutenthicateUserAdm()
		Dim sSql, arrUsuarioAdm, intUsuarioAdm, i

		sSql = "SELECT " & _ 
						"[idPontos_Coleta], " & _ 
						"[nome_fantasia] " & _ 
						"FROM [marketingoki2].[dbo].[Pontos_coleta] " & _
						"WHERE [usuario] = '"&User&"' AND " & _
						"[senha] = '"&Password&"' " & _
						"AND [status_pontocoleta] = 1"
	
		'Response.Write sSql & "<hr>"
		'Response.End
		
		Call search(sSql, arrUsuarioAdm, intUsuarioAdm)
	
		If intUsuarioAdm > -1 Then
			For i=0 To intUsuarioAdm
				Session("IDPonto") = arrUsuarioAdm(0,i)
				Session("Nome") = arrUsuarioAdm(1,i)
				Session("LoginADM") = User
				Session("PasswordADM") = Password
				session("CodBonus") = getBonus()
			Next
			AutenthicateUserAdm = True
		Else
			AutenthicateUserAdm = False	
		End If
	End Function
	
	function getBonus()
		dim sql, arr, intarr, i
		sql = "SELECT [idPontos_coleta] " & _
				",[bonus_type] " & _
				"FROM [marketingoki2].[dbo].[Pontos_coleta]"
		call search(sql, arr, intarr)		
		if intarr > -1 then
			for i=0 to intarr
				getBonus = arr(1,i)
			next
		else
			getBonus = ""
		end if
	end function

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			If AutenthicateUserAdm()  Then
				Response.Redirect "frmOperacionalPonto.asp"
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
					<form action="frmLoginPC.asp" name="frmLoginPC" method="POST">
						<table cellpadding="3" cellspacing="4" width="100%" id="tableLoginCliente">
							<tr>
								<td colspan="2" id="explaintitle" align="center">Login do Ponto de Coleta</td>
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
