<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%
	dim cod_consolidador
	dim idcliente

	function getSolicitacoesByCliente()
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim saldo
		dim saldoTotal
		dim sDataResgate
		saldoTotal = 0
		sql = "select distinct(bonus.numero_solicitacao) " & _
			  "from bonus_gerado_clientes as bonus  " & _
			  "left join clientes as cli " & _
			  "on bonus.clientes_idclientes = cli.idclientes " & _
			  "where bonus.clientes_idclientes in (select idclientes from clientes where cod_cli_consolidador = "&session("IDCliente")&" or idclientes = "&session("IDCliente")&")"
		'response.write sql & "<hr>"
		'response.end
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
		
			html = html & "<tr><td colspan=8><div id=pag>"
			html = html & PaginacaoExibir(intPag, intProdsPorPag, intarr)
			html = html & "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo, " & _
						"moeda, " & _
						"day(data_resgate), " & _
						"month(data_resgate), " & _
						"year(data_resgate) " & _
                        "from bonus_gerado_clientes where Clientes_idClientes = " & session("IDCliente")'peterson aquino 17-5-2014 id:7
						'"from bonus_gerado_clientes where numero_solicitacao = '"&arr(0,i)&"'"

				call search(sql, arr2, intarr2)

				if intarr > -1 then
					j=0
					sDataResgate = arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)
					saldo = arr2(8,j)
					if datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) > 0 or clng(saldo) = 0 or len(sDataResgate) = 0 then
						saldoTotal = saldoTotal + clng(saldo)
						if i mod 2 = 0 then
							style = "class=""classColorRelPar"""
						else
							style = "class=""classColorRelImpar"""
						end if
						html = html & "<tr>"
						if getIDSolicitacao(arr(0,i)) <> 0 then
							html = html & "<td "&style&"><img src=""img/buscar.gif"" class=""imgexpandeinfo"" onClick=""javascript:window.open('frmviewsol.asp?idsol="&getIDSolicitacao(arr(0,i))&"','','width=720,height=600,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');"" /></td>"
						else
							html = html & "<td "&style&"><img src=""img/buscar.gif"" /></td>"
						end if
						html = html & "<td "&style&">"&arr(0,i)&"</td>"
						html = html & "<td "&style&">"&arr2(2,j)&"/"&arr2(3,j)&"/"&arr2(4,j)&"</td>"
						html = html & "<td "&style&">"&arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j)&"</td>"
						if arr2(10,j) <> "" and arr2(11,j) <> "" and arr2(12,j) <> "" then
							html = html & "<td "&style&">"&arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)&"</td>"
						else
							html = html & "<td "&style&">##/##/#####</td>"
						end if
						html = html & "<td "&style&">"&arr2(9,j)&"</td>"
						html = html & "<td "&style&">"&getPontuacaoBySolicitacao(arr(0,i))&"</td>"
						html = html & "<td "&style&">"
						if clng(saldo) <> 0 then
							html = html & saldo
						else							
							html = html & "resgatado"
						end if						
						html = html & "</td>"
						html = html & "</tr>"
					end if
				end if
			next
			getSolicitacoesByCliente = html
		else
			html = html & "<tr><td colspan=""8"" align=""center"">Nenhum bônus foi resgatado até o momento</td></tr>"
			getSolicitacoesByCliente = html
		end if
	end function

	function getIDSolicitacao(numero)
		dim sql, arr, intarr, i
		dim idnumero
		sql = "SELECT [idSolicitacao_coleta] " & _
				  ",[numero_solicitacao_coleta] " & _
			  "FROM [marketingoki2].[dbo].[Solicitacao_coleta] " & _
			  "WHERE [numero_solicitacao_coleta] = '"&numero&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getIDSolicitacao = arr(0,i)
			next
		else
			getIDSolicitacao = 0
		end if
	end function

	function getPontuacaoBySolicitacao(solicitacao)
		dim sql, arr, intarr, i
		sql = "select " & _
				"sum(pontuacao) " & _
				"from bonus_gerado_clientes where numero_solicitacao = '"&solicitacao&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				getPontuacaoBySolicitacao = clng(arr(0,i))
			next
		else
			getPontuacaoBySolicitacao = 0
		end if
	end function

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<style>
	label {
		font-weight:bold;
	}
</style>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="frmviewbonusgeradocliente.asp" name="frmviewbonusgeradocliente" method="POST">
		<table cellpadding="1" cellspacing="1" width="750" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" colspan="2" align="center">Visualizar Bônus Resgatados</td>
			</tr>
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:615px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<tr>
							<th width="2%"><img src="img/check.gif"></th>
							<th>Número Solicitação</th>
							<th>Data Geração</th>
							<th>Data Validade</th>
							<th>Data Resgate</th>
							<th>Moeda</th>
							<th>Pontuação</th>
							<th>Saldo</th>
						</tr>
						<%=getSolicitacoesByCliente()%>
					</table>
					</div>
				</td>
			</tr>
			<tr>
				<td colspan="2">&nbsp;</td>
			</tr>
			<tr>
				<td colspan="2" id="msgret" align="center">&nbsp;</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
