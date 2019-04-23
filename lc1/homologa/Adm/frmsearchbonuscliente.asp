<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	function getBonus()
		Dim sql, arr, intarr, i, j, ret
		
		if len(trim(sSearch)) > 0 then
			sql = "SELECT [cod_bonus] " & _
						  ",[descricao] " & _
						  ",[validade] " & _
						  ",[moeda] " & _
						  ",[aplicacao] " & _
						  ",[data_inicio_contabilizacao] " & _
					  "FROM [marketingoki2].[dbo].[Cadastro_Bonus] where aplicacao = 'CLI' " & _
					  " and (descricao like '%"&sSearch&"%' or validade like '%"&sSearch&"%')"
		else
			sql = "SELECT [cod_bonus] " & _
						  ",[descricao] " & _
						  ",[validade] " & _
						  ",[moeda] " & _
						  ",[aplicacao] " & _
						  ",[data_inicio_contabilizacao] " & _
					  "FROM [marketingoki2].[dbo].[Cadastro_Bonus] where aplicacao = 'CLI'"
		end if		  
				  
		Call search(sql ,arr, intarr)				
		If intTransp > -1 Then
			
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = intarr+1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			ret = ret & "<tr><td colspan=3><div id=pag>"
			ret = ret & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			ret = ret & "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				j = j + 1
				If i Mod 2 = 0 Then
					ret = ret & "<tr>"
					ret = ret & "<td class=""classColorRelPar""><input type=""radio"" name=""bonus"" value="""&arr(0,i)&""" OnClick=""updateBonus()"" /></td>"
					ret = ret & "<td class=""classColorRelPar"" width=""15%"">"&arr(0,i)&"</td>"
					ret = ret & "<td class=""classColorRelPar"" width=""4%"">"&arr(2,i)&"</td>"
					ret = ret & "<td class=""classColorRelPar"" width=""4%"">"&arr(3,i)&"</td>"
					ret = ret & "<td class=""classColorRelPar"">"&arr(1,i)&"</td>"
					ret = ret & "</tr>"
				Else
					ret = ret & "<tr>"
					ret = ret & "<td class=""classColorRelImpar""><input type=""radio"" name=""bonus"" value="""&arr(0,i)&""" OnClick=""updateBonus()"" /></td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arr(0,i)&"</td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arr(2,i)&"</td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arr(3,i)&"</td>"
					ret = ret & "<td class=""classColorRelImpar"">"&arr(1,i)&"</td>"
					ret = ret & "</tr>"
				End If	
			Next
			ret = ret & "<input type=""hidden"" name=""hiddenIntBonus"" value="&j&" />"
		End If
		getBonus = ret
	End function
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		sSearch = request("txtSearch")
		call getBonus()
	end if
%>
<html>
<head>
<script>
	function updateBonus() {
		var error = 0;
		var valor = null;
		for (var i=0; i < parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value); i++) {
			if (parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value) == 1) {
				if (!document.frmsearchbonuscliente.bonus.checked) {
					error++;
				} else {
					valor = document.frmsearchbonuscliente.bonus.value;	
				}
			} else {
				if (!document.frmsearchbonuscliente.bonus[i].checked) {
					error++;
				}	else {
					valor = document.frmsearchbonuscliente.bonus[i].value;	
				}
			}
		}
		if (error == parseInt(document.frmsearchbonuscliente.hiddenIntBonus.value)) {
			alert("Por favor escolha um Bônus");
			return;
		} else {
			window.opener.frmCadastroClienteAdm.cbBonus.value = valor;		
			window.close();
		}
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
					<div style="overflow:scroll;width:600px;height:317px;">
						<form action="frmsearchbonuscliente.asp" name="frmsearchbonuscliente" method="POST">
						<table cellspacing="1" cellpadding="1" width="600" id="tablelisttransportadoras">
							<tr>
								<td colspan="5">
									Procurar: <INPUT type="text" id=txtSearch name=txtSearch>
									<INPUT type="submit" value="Procurar" id=submit1 name=submit1>
								</td>
							</tr>
							<tr>
								<td id="explaintitle" colspan="5" align="center">Busca de Bônus</td>
							</tr>
							<tr>
								<th width="5%" ><img src="img/check.gif"></th>
								<th>Cód. Bônus</th>
								<th>Validade</th>
								<th>Moeda</th>
								<th>Descrição</th>
							</tr>
							<%=getBonus()%>
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
