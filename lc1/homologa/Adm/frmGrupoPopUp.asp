<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Dim CNPJ
	Dim intGrupoCliente
	
	CNPJ = Request.QueryString("cnpj")
	
	If Request.QueryString("action") = "alterar" Then
		Call AtualizarGrupo()
	End If	
	
	Sub AtualizarGrupo()
		Dim sSql
		sSql = "UPDATE Clientes SET Grupos_idGrupos = "&Request.QueryString("idgrupo")&" WHERE cnpj = '"&Request.QueryString("cnpj")&"' "
		Call exec(sSql)
		Response.Write "<script>window.opener.location.reload();</script>"
		Response.Write "<script>window.close();</script>"
	End Sub
	
	Sub GetCliente()
		Dim sSql, arrCliente, intCliente, i
		Dim contador
		
		sSql = "SELECT A.[idClientes] " & _
					  ",B.[descricao] " & _
					  ",A.[nome_fantasia] " & _
					  ",A.[cnpj] " & _
					  ",B.[idGrupos] " & _
				  "FROM [marketingoki2].[dbo].[Clientes] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Grupos] AS B " & _
				  "ON A.[Grupos_idGrupos] = B.[idGrupos] " & _
				  "WHERE A.[cnpj] LIKE '%"&Left(CNPJ, 10)&"%'"
		Call search(sSql, arrCliente, intCliente)
		If intCliente > -1 Then
			For i=0 To intCliente
				If arrCliente(3,i) <> CNPJ Then
					intGrupoCliente = intGrupoCliente + 1
					If i Mod 2 = 0 Then
						Response.Write "<tr>"
						Response.Write "<td class='classColorRelPar' width=""5%""><input type=""radio"" name=""radioThisGroup"" value="""&arrCliente(4,i)&""" onClick=""UpdateGrupo()"" /></td>"
						If arrCliente(1,i) <> "" Then
							Response.Write "<td class='classColorRelPar'>"&arrCliente(1,i)&"</td>"
						Else
							Response.Write "<td class='classColorRelPar' width=""40%"">Sem grupo associado</td>"
						End If	
						Response.Write "<td class='classColorRelPar'>"&arrCliente(2,i)&"</td>"
						Response.Write "</tr>"
					Else
						Response.Write "<tr>"
						Response.Write "<td class='classColorRelImpar' width=""5%""><input type=""radio"" name=""radioThisGroup"" value="""&arrCliente(4,i)&""" onClick=""UpdateGrupo()"" /></td>"
						If arrCliente(1,i) <> "" Then
							Response.Write "<td class='classColorRelImpar'>"&arrCliente(1,i)&"</td>"
						Else
							Response.Write "<td class='classColorRelImpar' width=""40%"">Sem grupo associado</td>"
						End If	
						Response.Write "<td class='classColorRelImpar'>"&arrCliente(2,i)&"</td>"
						Response.Write "</tr>"
					End If
				Else
					contador = contador + 1
					If contador = intCliente + 1 Then
						Response.Write "<tr><td colspan=""3"" align=""center""><b>Nenhum Cliente encontrado!</b></td></tr>"
					End If
				End If
			Next
		Else
			Response.Write "<tr><td colspan=""3""><b>Nenhum Cliente encontrado!</b></td></tr>"	
		End If
		Response.Write "<input type=""hidden"" name=""hiddenIntGrupoCliente"" value="""&intGrupoCliente&""" />"
	End Sub
	
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
<script src="js/frmGrupoPopUp.js"></script>
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo">
					<form action="frmGrupoPopUp.asp" name="frmGrupoPopUp" method="POST">
					<input type="hidden" name="cnpj" value="<%=CNPJ%>" />
						<table cellpadding="1" cellspacing="1" width="550">
							<tr>
								<td id="explaintitle" align="center">Busca para definição de Grupo do Cliente</td>
							</tr>
							<tr>
								<td width="100%" height="400" valign="top">
									<div style="overflow:scroll;height:400;">
										<table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente">
											<tr>
												<th><img src="img/check.gif"></th>
												<th>Grupo</th>
												<th>Nome Fantasia</th>
											</tr>
											<%Call GetCliente()%>
										</table>
									</div>
								</td>
							</tr>
						</table>
					</form>
				</td>
			</tr>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>
