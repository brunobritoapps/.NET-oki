<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	function getClientes()
		dim sql, arr, intarr, i
		dim html
		dim style
		html = ""
		sql = "SELECT [idClientes] " & _
			  ",[razao_social] " & _
			  ",[cnpj] " & _
			  ",[cod_bonus_cli] " & _
			  ",[cod_cli_consolidador] " & _
			  "FROM [marketingoki2].[dbo].[Clientes] where [cod_bonus_cli] <> '' and [status_cliente] = 1 and [cod_cli_consolidador] = 0"
		if trim(request.QueryString("busca")) <> "" then
			sql = sql & " AND cnpj = '"&trim(request.QueryString("busca"))&"'"	  
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
			intNumProds = UBound(arr, 2) + 1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			html = html & "<tr><td colspan=10>"
			html = html & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			html = html & "</td></tr>"
	
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if	
				html = html & "<tr>"
				html = html & "<td "&style&"><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Visualizar Bônus"" onClick=""javascript:window.open('frmviewbonusgeradoclientesintatico.asp?id="&arr(0,i)&"','','width=750,height=650,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" /></td>"
				html = html & "<td "&style&">"&arr(0,i)&"</td>"
				html = html & "<td "&style&">"&arr(1,i)&"</td>"
				html = html & "<td "&style&">"&arr(2,i)&"</td>"
				html = html & "<td "&style&">"&arr(3,i)&"</td>"
				html = html & "</tr>"
			next
		else
			html = html & "<tr><td colspan=""4"" align=""center"">Nenhum registro encontrado.</td></tr>"
		end if

		'html = html & "<tr><td colspan=7>"
		'html = html & Paginacao(iNumPags, intarr, request("pag"), "frmBonusGeradoClientes", Request.ServerVariables("QUERY_STRING"))
		'html = html & "</td></tr>"
		
		getClientes = html	  
	end function
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
		<form action="frmOperacionalAdm.asp" name="frmOperacionalAdm" method="POST">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table width="100%" cellpadding="1" cellspacing="1">
						<tr>
							<td id="explaintitle" align="center">Bônus Gerado [Cliente]</td>
						</tr>
						<tr>
							<td align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="4" id="explaintitle" align="left">
								&nbsp;&nbsp;CNPJ: <input type="text" name="busca" class="text" value="<%= request.querystring("busca") %>" size="40" />
								<input type="button" name="btnprocurar" value="Procurar" class="btnform" onClick="window.location.href='frmbonusgeradoclientes.asp?busca=' + document.frmOperacionalAdm.busca.value + ''" />
							</td>
						</tr>
						<tr>
							<td>
								<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;">
									<tr>
										<th width="2%"><img src="img/check.gif" /></th>
										<th width="7%">ID Cliente</th>
										<th>Razão Social</th>
										<th>CNPJ</th>
										<th width="20%">Cód. Bônus</th>
									</tr>
									<%=getClientes()%>
								</table>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</form>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
