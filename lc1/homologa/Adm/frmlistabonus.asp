<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim sSearch
	
	function getBonus()
		dim sql, arr, intarr, i, j
		dim ret
		dim style
		style = "class=""classColorRelPar"""
		ret = ""

		if len(trim(sSearch)) > 0 then
			sql = "SELECT [cod_bonus] " & _
					  ",[descricao] " & _
					  ",[validade] " & _
					  ",[moeda] " & _
					  ",[aplicacao] " & _
					  ",[data_inicio_contabilizacao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_Bonus] " & _
				  "WHERE cod_bonus like '%"&sSearch&"%' or descricao like '%"&sSearch&"%'"
		else
			sql = "SELECT [cod_bonus] " & _
					  ",[descricao] " & _
					  ",[validade] " & _
					  ",[moeda] " & _
					  ",[aplicacao] " & _
					  ",[data_inicio_contabilizacao] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_Bonus]"
		end if
			  
		call search(sql, arr, intarr)	  
		
		if intarr > -1 then
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
		
			ret = ret & "<tr><td colspan=9><div id=pag>"
			ret = ret & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			ret = ret & "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				j = j + 1
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				ret = ret & "<tr>"
				ret = ret & "<td width=""10"" "&style&"><input type=""radio"" name=""radio_bonus"" value="""&arr(0,i)&""" onclick=""updateBonus()"" /></td>"
				ret = ret & "<td width=""150"" "&style&"> <img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Veja os produtos deste Bônus"" onClick=""javascript:window.open('frmlistaprodbonus.asp?cod="&arr(0,i)&"','','width=700,height=400,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" /> - "&arr(0,i)&"</td>"
				ret = ret & "<td "&style&">"&arr(1,i)&"</td>"
				ret = ret & "</tr>"
			next
			ret = ret & "<input type=""hidden"" name=""hiddenIntBonus"" value="&j&" />"
		else
			ret = ret & "<tr><td colspan=""4"" align=""center"" class=""classColorRelPar""><b>Nenhum registro encontrado</b></td></tr>"	
		end if
		getBonus = ret
	end function
	
	if request.servervariables("HTTP_METHOD") = "POST" then
		sSearch = request("txtSearch")
		call getBonus()
	end if
%>
<html>
<head>
<script>

function updateBonus() {
	for (var i=0; i < parseInt(document.frmlistabonus.hiddenIntBonus.value); i++) {
		if (parseInt(document.frmlistabonus.hiddenIntBonus.value) == 1) {
			if (document.frmlistabonus.radio_bonus.checked) {
				window.opener.location.href = 'frmcadbonus.asp?id='+document.frmlistabonus.radio_bonus.value;
				window.close();
			}	
		} else {
			if (document.frmlistabonus.radio_bonus[i].checked) {
				window.opener.location.href = 'frmcadbonus.asp?id='+document.frmlistabonus.radio_bonus[i].value;
				window.close();
			}	
		}
	}
}

</script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<div style="width:700px;height:400px;overflow:scroll;">
			<table cellspacing="0" cellpadding="0" width="700" ID="Table1" align="left" border="0">
			<form action="#" name="frmlistabonus" method="POST">
				<tr> 
					<td id="conteudo">
						<table cellspacing="1" cellpadding="1" width="700" align="left" border="0" id="tableRelSolPendente">
							<tr>
								<td colspan="4" id="explaintitle" align="center">Lista de Bônus</td>
							</tr>
							<tr>
								<td colspan="4">
									Procurar: <INPUT type="text" id=txtSearch name=txtSearch>
									<INPUT type="submit" value="Procurar">
								</td>
							</tr>
							<tr>
								<th><img src="img/check.gif" /></th>
								<th>Cód. Bônus</th>
								<th>Descrição</th>
							</tr>
							<%=getBonus()%>
						</table>
					</td>
				</tr>
			</form>
			</table>
		</div>	
	</div>
</div>
</body>
</html>
<%Call close()%>
