<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
'	function getCliente(id)
'		dim sql, arr, intarr, i
'		dim retorno
'		
'		sql = "SELECT [idClientes] " & _
'			  ",[cod_cli_consolidador] " & _
'			  "FROM [marketingoki2].[dbo].[Clientes] where idClientes = " & id
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
'		if getCliente(id) > -1 then 
'			sql = "select distinct(numero_solicitacao) from bonus_gerado_clientes where clientes_idclientes = " & getCliente(id)
'			response.write sql
'			response.end
'			call search(sql, arr, intarr)
'			if intarr > -1 then
'				for i=0 to intarr
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
							"idproduto " & _
							"from bonus_gerado_clientes where numero_solicitacao = '"&id&"'"
					call search(sql, arr2, intarr2)		
					if intarr > -1 then
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
'						html = html & "</tr>"
						for j=0 to intarr2
							saldo = arr2(8,j)
							if j mod 2 = 0 then
								style = "class=""classColorRelPar"""
							else
								style = "class=""classColorRelImpar"""
							end if
							html = html & "<tr>"
							html = html & "<td "&style&">"&arr2(9,j)&"</td>"
							html = html & "<td "&style&">"&arr2(0,j)&"</td>"
							html = html & "<td "&style&">"&arr2(1,j)&"</td>"
							html = html & "</tr>"
						next
'						html = html & "<tr>"
'						html = html & "<td colspan=""5"" align=""right""><b>Saldo da Pontua��o: </b>"&saldo&"</td>"
'						html = html & "</tr>"
'					else
'						html = html & "<tr><td colspan=""5"">"&arr(0,i)&"/td></tr>"
					end if
				getSolicitacoesByCliente = html
'			else
'				html = html & "<tr><td colspan=""5"" align=""center"">Nenhum b�nus foi gerado at� o momento</td></tr>"
'				getSolicitacoesByCliente = html
'			end if
'		end if	
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
				<td id="explaintitle" colspan="2" align="center">Visualizar B�nus Gerado</td>
			</tr>
			<tr>
				<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.history.back(1);">&laquo; Voltar</a></td>
			</tr>
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:615px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<tr>
							<td colspan="5" align="right"><b>Solicita��o: </b><%=request.querystring("id")%></td>
						</tr>
						<tr>
							<th>IDProduto</th>
							<th>Pontua��o</th>
							<th>Pontua��o Target</th>
						</tr>
						<%=getSolicitacoesByCliente(request.querystring("id"))%>
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
