<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	function getBonusProd()
		dim sql, arr, intarr, i
		dim ret
		dim style
		style = "class=""classColorRelPar"""
		ret = ""

		sql = "SELECT [idoki_prod] " & _
					  ",[qtd] " & _
					  ",[pontuacao] " & _
					  ",[cad_cod_bonus] " & _
					  ",[pontuacao_target] " & _
				  "FROM [marketingoki2].[dbo].[Cadastro_bonus_has_produtos] " & _
				  "WHERE [cad_cod_bonus] = '"&request.querystring("cod")&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 20 'numero de registros mostrados na pagina
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
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				ret = ret & "<tr>"
				ret = ret & "<td "&style&">"&arr(0,i)&"</td>"
				ret = ret & "<td "&style&">"&arr(1,i)&"</td>"
				ret = ret & "<td "&style&">"&arr(2,i)&"</td>"
				ret = ret & "<td "&style&">"&arr(4,i)&"</td>"
				ret = ret & "<td "&style&"><INPUT type=checkbox id=chkDel name=chkDel value="&arr(0,i)&"></td>"
				ret = ret & "</tr>"
			next
		else
			ret = ret & "<tr><td colspan=""3"" align=""center""><b>Nenhum registro encontrado</b></td></tr>"
		end if
		getBonusProd = ret
	end function
	
	if request.ServerVariables("HTTP_METHOD") = "POST" then
		dim v, i, sql2

		v = split(request("chkDel"), ",")
		for i=0 to ubound(v)
			sql2 = "delete from Cadastro_bonus_has_produtos WHERE idoki_prod = '"&trim(v(i))&"' AND cad_cod_bonus = '"&trim(request("cod"))&"'"
			call exec(sql2)
		next		
		Response.Redirect "frmlistaprodbonus.asp?cod="&request("cod")
	end if
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:400px;width:698px;overflow:auto;">
		<table cellspacing="0" cellpadding="0" width="700" ID="Table1" align="left">
		<form action="#" name="frmlistabonus" method="POST">
			<tr>
				<td id="conteudo">
					<table cellspacing="1" cellpadding="1" width="698" align="left" border="0"s id="tableRelSolPendente">
						<tr>
							<td colspan="5" id="explaintitle" align="center">Lista de Produtos do Bônus <%=request.querystring("cod")%></td>
						</tr>
						<tr>
							<th>ID Pród.</th>
							<th>quantidade</th>
							<th>pontuação</th>
							<th>pontuação target</th>
							<th>Deletar?</th>
						</tr>
						<%=getBonusProd()%>
						<tr>
							<td colspan="5">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="5" align=right>
								<INPUT type="hidden" value="<%=request.querystring("cod")%>" id=cod name=cod>
								<INPUT type="submit" value="Deletar">
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</form>
		</table>
	</div>
</div>
</body>
</html>
<%Call close()%>
