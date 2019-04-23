<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetTransportadoras()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT " & _
						"[idTransportadoras], " & _ 
						"[nome_fantasia], " & _  
						"[cnpj] " & _
						"FROM [marketingoki2].[dbo].[Transportadoras] " & _
						"WHERE [status] = 1 and iscoletaemail = 1"
		Call search(sSql ,arrTransp, intTransp)				
		If intTransp > -1 Then
			Response.Write "<input type=""hidden"" name=""hiddenIntTransp"" value="&intTransp + 1&" />"
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
%>
<html>
<head>
<script>
	function updateTransp() {
		var error = 0;
		var valor = null;
	//	alert(parseInt(document.frmSearchTranspSol.hiddenIntTransp.value + 1));
		for (var i=0; i < parseInt(document.frmsearchtranspgetfile.hiddenIntTransp.value); i++) {
			if (parseInt(document.frmsearchtranspgetfile.hiddenIntTransp.value) == 1) {
				if (!document.frmsearchtranspgetfile.transp.checked) {
					error++;
				} else {
					valor = document.frmsearchtranspgetfile.transp.value;	
				}
			} else {
				if (!document.frmsearchtranspgetfile.transp[i].checked) {
					error++;
				}	else {
					valor = document.frmsearchtranspgetfile.transp[i].value;	
				}
			}
		}
		if (error == parseInt(document.frmsearchtranspgetfile.hiddenIntTransp.value)) {
			alert("Por favor escolha uma transportadora");
			return;
		} else {
			window.opener.frmgerafiletransp.cbTransp.value = valor;		
			redirTransp(valor);
			window.close();
		}
	}
	
	function redirTransp(valor) {
		window.opener.location = 'frmgerafiletransp.asp?id='+valor;
	}
</script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo">
					<div style="overflow:scroll;width:410px;height:317px;">
						<form action="frmsearchtranspgetfile.asp" name="frmsearchtranspgetfile" method="POST">
						<table cellspacing="1" cellpadding="1" width="395" id="tablelisttransportadoras">
							<tr>
								<td id="explaintitle" colspan="3" align="center">Busca de Transportadora</td>
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
