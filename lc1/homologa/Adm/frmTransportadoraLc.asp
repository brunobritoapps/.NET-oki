<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
    
   Sub Atualiza()
	
    Dim sSql

	if cint(Request.Form("transp")) = 0 then
			sSql = "UPDATE Clientes " & _
							"SET " & _  
							"[Transportadoras_idTransportadoras] = NULL " & _ 
							"WHERE [idClientes] = " & Request.Form("id")
	else
		sSql = "UPDATE Clientes " & _
							"SET " & _  
							"[Transportadoras_idTransportadoras] = "&Request.Form("transp")&" " & _ 
							"WHERE [idClientes] = " & Request.Form("id")
		end if					
        Response.Write sSql

		Call exec(sSql)
		Response.Write "<script>window.opener.location.reload();</script>"
		Response.Write "<script>window.close();</script>"						
	End Sub

	Sub GetTransportadoras()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT " & _
						"[idTransportadoras], " & _ 
						"[nome_fantasia], " & _  
						"[cnpj] " & _
						"FROM [marketingoki2].[dbo].[Transportadoras] " & _
						"WHERE [status] = 1"
		Call search(sSql ,arrTransp, intTransp)				
		If intTransp > -1 Then
			Response.Write "<input type=""hidden"" name=""hiddenIntTransp"" value="&intTransp + 2&" />"
			For i=0 To intTransp
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class=""classColorRelPar""><input type=""radio"" name=""transp"" value="""&arrTransp(0,i)&""" OnClick=""updateTransp()"" /></td>"
					Response.Write "<td class=""classColorRelPar"">"&arrTransp(1,i)&"</td>"
					Response.Write "<td class=""classColorRelPar"">"&arrTransp(2,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class=""classColorRelImpar""><input type=""radio"" name=""transp"" value="""&arrTransp(0,i)&""" OnClick=""updateTransp()"" /></td>"
					Response.Write "<td class=""classColorRelImpar"">"&arrTransp(1,i)&"</td>"
					Response.Write "<td class=""classColorRelImpar"">"&arrTransp(2,i)&"</td>"
					Response.Write "</tr>"
				End If	
			Next
		End If
	End Sub
	
	If Request.ServerVariables("HTTP_METHOD") = "POST" Then
        Call Atualiza()
    End If

%>

<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
<script language="javascript" type="text/javascript" src="js/frmTransportadoraLc.js"></script>
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo">
					<div style="overflow:scroll;width:410px;height:317px;">
						<form action="frmTransportadoraLc.asp" name="frmTransportadoraLc" method="POST">
						<input type="hidden" name="id" value="<%If Request.QueryString("idcli") <> "" Then Response.Write Request.QueryString("idcli") Else Response.Write Request.Form("id") End If%>" />
						<table cellspacing="1" cellpadding="1" width="395" id="tablelisttransportadoras">
							<tr>
								<td id="explaintitle" colspan="3" align="center">Busca de Transportadora</td>
							</tr>
							<!--tr>
								<td id="explaintitle" colspan="3" align="right"><input name="escolher" class="btnform" type="button" value="Escolher" OnClick="updateTransp()" /></td>
							</tr-->
							<tr>
								<td id="explaintitle" colspan="3"><input type="radio" name="transp" value="0" onClick="updateTransp()" /> Nenhuma Transportadora</td>
							</tr>
							<tr>
								<th width="5%" ><img src="img/check.gif"></th>
								<th>Nome Fantasia</th>
								<th>CNPJ</th>
							</tr>
							<%Call GetTransportadoras()%>
						</table>
						</form>
					</div>
				</td>
			</tr>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>