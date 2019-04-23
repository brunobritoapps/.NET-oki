<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
'	function getCliente(id)
'		dim sql, arr, intarr, i
'		dim retorno
'		
'		sql = "SELECT [idPontos_coleta] " & _
'			  "FROM [marketingoki2].[dbo].[Pontos_coleta] where [idPontos_coleta] = " & id
'		call search(sql, arr, intarr)	  
'		if intarr > -1 then
'			for i=0 to intarr
'				retorno = arr(0,i)
'			next
'		else
'			retorno = -1
'		end if
'		getCliente = retorno
'	end function
	
	function getSolicitacoesByCliente(id)
		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim saldo
'		if getCliente(id) > -1 then 
'			sql = "select distinct(numero_solicitacao) from bonus_gerado_pontocoleta where Pontos_coleta_idPontos_coleta = " & getCliente(id)
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
							"from bonus_gerado_pontocoleta where numero_solicitacao = '"&id&"'"
					call search(sql, arr2, intarr2)		
					if intarr > -1 then
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&id&" <img src=""img/buscar.gif"" class=""imgexpandeinfo"" align=""absmiddle"" alt=""Buscar Solicitações que compuseram a solicitação Master"" onClick=""javascript:window.open('frmviewcompoemasteradm.asp?idsolic="&id&"','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');""/></th>"
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
'						html = html & "<td colspan=""5"" align=""right""><b>Saldo da Pontuação: </b>"&saldo&"</td>"
'						html = html & "</tr>"
					else
						html = html & "<tr><td colspan=""5"">"&arr(0,i)&"/td></tr>"
					end if
'				next
				getSolicitacoesByCliente = html
'			else
'				html = html & "<tr><td colspan=""5"" align=""center"">Nenhum bônus foi gerado até o momento</td></tr>"
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
				<td id="explaintitle" colspan="2" align="center">Visualizar Bônus Gerado</td>
			</tr>
			<tr>
				<td colspan="2" align="right"><a class="linkOperacional" href="javascript:window.history.back(1);">&laquo; Voltar</a></td>
			</tr>
			<tr id="trnumsolcoleta">
				<td>
					<div style="overflow:auto;width:100%;height:615px;">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableGetClientesCadastro" style="border:1px solid #000000;" >
						<tr>
							<td colspan="5" align="right"><b>Solicitação: </b><%=request.querystring("id")%>&nbsp;&nbsp;<img class="imgexpandeinfo" align="absmiddle" src="img/buscar.gif" onClick="javascript:window.open('frmviewcompoemasteradm.asp?idsolic=<%=request.querystring("id")%>','','width=650,height=250,scrollbars=no,status=no,location=no,toolbar=no,menubar=no');" /></td>
						</tr>
						<tr>
							<th>IDProduto</th>
							<th>Pontuação</th>
							<th>Pontuação Target</th>
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
