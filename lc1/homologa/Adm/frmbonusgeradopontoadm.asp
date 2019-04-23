<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
'	function getCliente(id)
'		dim sql, arr, intarr, i
'		dim retorno
'		
'		sql = "SELECT [idPontos_coleta] " & _
'			  "FROM [marketingoki2].[dbo].[Pontos_coleta] where idClientes = " & id
'		call search(sql, arr, intarr)	  
'		if intarr > -1 then
'			for i=0 to intarr
'				if arr(1,i) <> "" and not isnull(arr(1,i)) and not isempty(arr(1,i)) and arr(1,i) > 0 then
'					retorno = arr(1,i)
'				else
'					retorno = arr(0,i)
'				end if
'			next
'		else
'			retorno = -1
'		end if
'		getCliente = retorno
'	end function
	
'	function getDescByCliente(id)
'		dim sql, arr, intarr, i
'		sql = "SELECT [razao_social] " & _
'				  ",[cnpj] " & _
'				  ",[cod_cli_consolidador] " & _
'			  "FROM [marketingoki2].[dbo].[Clientes] " & _
'			  "where [idClientes] = " & id
'		call search(sql, arr, intarr)	  
'		if intarr > -1 then
'			for i=0 to intarr
'				cliente_consolidador = arr(0,i)
'				cnpj_consolidador = arr(1,i)
'			next	
'		end if
'	end function

	function getSolicitacoesByCliente(id)
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim saldo
		dim saldoTotal
		dim sDataResgate
		saldoTotal = 0
		if id > -1 then 
			sql = "select distinct(numero_solicitacao) from bonus_gerado_pontocoleta where Pontos_coleta_idPontos_coleta = " & id
			'response.write sql
			'response.end
			call search(sql, arr, intarr)
			if intarr > -1 then
				for i=0 to intarr
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
							"from bonus_gerado_pontocoleta where numero_solicitacao = '"&arr(0,i)&"'"
					'response.write sql
					call search(sql, arr2, intarr2)		
					if intarr > -1 then
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
'						html = html & "</tr>"

						j=0
						saldo = arr2(8,j)
						
						if isnull(saldo) then
							saldo=0
						end if
						
						sDataResgate = arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)
						if datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) < 0 and clng(saldo) <> 0 then 'and len(sDataResgate) = 2 
							saldoTotal = saldoTotal + cint(saldo)
							if i mod 2 = 0 then
								style = "class=""classColorRelPar"""
							else
								style = "class=""classColorRelImpar"""
							end if
							html = html & "<tr>"
							html = html & "<td "&style&">"&arr(0,i)&"</td>"
							html = html & "<td "&style&">"&arr2(2,j)&"/"&arr2(3,j)&"/"&arr2(4,j)&"</td>"
							html = html & "<td "&style&">"&arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j)&"</td>"
							if arr2(10,j) <> "" and arr2(11,j) <> "" and arr2(12,j) then
								html = html & "<td "&style&">"&arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)&"</td>"
							else
								html = html & "<td "&style&"></td>"
							end if	
							html = html & "<td "&style&">"&arr2(9,j)&"</td>"
							html = html & "<td "&style&">"&getPontuacaoBySolicitacao(arr(0,i))&"</td>"
							html = html & "<td "&style&">"&saldo&"</td>"
							html = html & "</tr>"
						end if	
					end if
				next
				html = html & "<tr>"
				html = html & "<td colspan=""8"" align=""right""><a href=""frmviewopcoesresgate.asp"" class=""linkOperacional"">Opções de Resgate</a>&nbsp;&nbsp;&nbsp;&nbsp;<b>Saldo Acumulado:</b> "&saldoTotal&"&nbsp;&nbsp;</td>"
				html = html & "</tr>"
				getSolicitacoesByCliente = html
			else
				html = html & "<tr><td colspan=""8"" align=""center"" class=""classColorRelPar""><b>Nenhum bônus foi gerado até o momento</b></td></tr>"
				getSolicitacoesByCliente = html
			end if
		end if	
	end function
	
	function getPontuacaoBySolicitacao(solicitacao)
		dim sql, arr, intarr, i
		sql = "select " & _
				"sum(pontuacao) " & _ 
				"from bonus_gerado_pontocoleta where numero_solicitacao = '"&solicitacao&"'"
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
<link rel="stylesheet" type="text/css" href="../css/geral.css">
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
				<td id="explaintitle" colspan="2" align="center">Visualizar Bônus Gerado</td>
			</tr>
			<tr>
				<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalPonto.asp';">&laquo Voltar</a></td>
			</tr>			
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:615px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<tr>
							<th>Número Solicitação</th>
							<th>Data Geração</th>
							<th>Data Validade</th>
							<th>Data Resgate</th>
							<th>Moeda</th>
							<th>Pontuação</th>
							<th>Saldo</th>
						</tr>
						<%=getSolicitacoesByCliente(session("IDPonto"))%>
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
