<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	function GetClientes()
		Dim sSql, arrCli, intCli, i, j, ret
		
		if len(trim(sSearch)) > 0 then
			sSql = "SELECT " & _
							"[idClientes], " & _ 
							"[nome_fantasia], " & _
							"[cnpj] " & _
							"FROM [marketingoki2].[dbo].[Clientes] " & _
							"WHERE [status_cliente] = 1 and (nome_fantasia like '%"& sSearch &"%' or cnpj like '%"&sSearch&"%')"
		else
			sSql = "SELECT " & _
							"[idClientes], " & _ 
							"[nome_fantasia], " & _
							"[cnpj] " & _
							"FROM [marketingoki2].[dbo].[Clientes] " & _
							"WHERE [status_cliente] = 1"
		end if				
						
		Call search(sSql, arrCli, intCli)
		If intCli > -1 Then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = intCli+1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			ret = ret & "<tr><td colspan=3><div id=pag>"
			ret = ret & PaginacaoExibir(intPag, intProdsPorPag, intCli)
			ret = ret & "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				j = j + 1
				If i Mod 2 = 0 Then
					ret = ret & "<tr>"
					ret = ret & "<td class=""classColorRelPar""><input type=""radio"" name=""radioIntCliente"" value="&arrCli(0,i)&" onClick=""updateCodConsolidador()"" /></td>"
					ret = ret & "<td class=""classColorRelPar"">"&arrCli(1,i)&"</td>"
					ret = ret & "<td class=""classColorRelPar"">"&arrCli(2,i)&"</td>"
					ret = ret & "</tr>"
				Else
					ret = ret & "<tr>"
					ret = ret & "<td class=""classColorRelImpar""><input type=""radio"" name=""radioIntCliente"" value="&arrCli(0,i)&" onClick=""updateCodConsolidador()"" /></td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arrCli(1,i)&"</td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arrCli(2,i)&"</td>"
					ret = ret & "</tr>"
				End If	
			Next
			ret = ret & "<input type=""hidden"" name=""hiddenIntCliente"" value="&j&" />"
		End If				
		GetClientes = ret
	End function
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		sSearch = request("txtSearch")
		call GetClientes()
	end if
%>
<html>
<head>
<script language="javascript" type="text/javascript">
	function updateCodConsolidador() {
		 var form = document.frmSearchCodCOnsolidador;
		 var cont = 0;
		 for (var i = 0; i < form.hiddenIntCliente.value; i++) {
		 		if (form.hiddenIntCliente.value == 1) {
					if (form.radioIntCliente.checked) {
//						alert(form.radioIntCliente.value);
						window.opener.frmCadastroClienteAdm.cbCodConsolidador.value = form.radioIntCliente.value;
					}		
				} else {
					if (form.radioIntCliente[i].checked) {
//						alert(form.radioIntCliente[i].value);
						window.opener.frmCadastroClienteAdm.cbCodConsolidador.value = form.radioIntCliente[i].value;
					}	
				}
		 }
		 window.close();
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
						<form action="frmSearchCodCOnsolidador.asp" name="frmSearchCodCOnsolidador" method="POST">
						<table cellspacing="1" cellpadding="1" width="395" id="tablelisttransportadoras">
							<tr>
								<td>
									Procurar: <INPUT type="text" id=txtSearch name=txtSearch>
									<INPUT type="submit" value="Procurar">
								</td>
							</tr>
							<tr>
								<td id="explaintitle" colspan="3" align="center">Busca de Clientes</td>
							</tr>
 							<tr>
								<th width="5%" ><img src="img/check.gif"></th>
								<th>Nome Fantasia</th>
								<th>CNPJ</th>
							</tr>
							<%=GetClientes()%>
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
