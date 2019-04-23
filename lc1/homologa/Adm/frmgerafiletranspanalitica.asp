<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	function getSolicitacaoByTransportadora(id)
		dim sql, arr, intarr, i

		if isMaster(request.querystring("numsol")) then
			sql = "select " & _
					"e.idtransportadoras, " & _   
					"e.razao_social, " & _  
					"a.numero_solicitacao_coleta, " & _  
					"a.data_envio_transportadora, " & _  
					"a.qtd_cartuchos, " & _  
					"c.razao_social, " & _  
					"c.logradouro, " & _  
					"c.numero_endereco, " & _  
					"c.bairro, " & _  
					"c.cep, " & _  
					"c.municipio, " & _  
					"c.estado " & _
					"from solicitacao_coleta as a " & _  
					"left join solicitacao_coleta_has_pontos_coleta as b " & _  
					"on a.idsolicitacao_coleta = b.solicitacao_coleta_idsolicitacao_coleta " & _  
					"left join pontos_coleta as c " & _  
					"on b.pontos_coleta_idpontos_coleta = c.idpontos_coleta " & _  
					"left join solicitacao_coleta_has_transportadoras as d " & _  
					"on a.idsolicitacao_coleta = d.solicitacao_coleta_idsolicitacao_coleta " & _  
					"left join transportadoras as e " & _  
					"on d.transportadoras_idtransportadoras = e.idtransportadoras " & _  
					"where a.status_coleta_idstatus_coleta = 2 " & _  
					"and a.ismaster = 1 " & _  
					"and e.razao_social <> 'NULL' " & _  
					"and e.idtransportadoras = " & id
		else
			sql = "select " & _
					"e.idtransportadoras, " & _
					"e.razao_social, " & _
					"a.numero_solicitacao_coleta, " & _
					"a.data_envio_transportadora, " & _
					"a.qtd_cartuchos, " & _
					"c.razao_social, " & _
					"f.logradouro, " & _
					"c.numero_endereco_coleta, " & _
					"f.bairro, " & _
					"f.cep, " & _
					"f.municipio, " & _
					"f.estado, " & _
					"b.contato_coleta, " & _
					"b.ddd_resp_coleta, " & _
					"b.telefone_resp_coleta " & _
					"from solicitacao_coleta as a " & _
					"left join solicitacao_coleta_has_clientes as b " & _
					"on a.idsolicitacao_coleta = b.solicitacao_coleta_idsolicitacao_coleta " & _
					"left join clientes as c " & _
					"on b.clientes_idclientes = c.idclientes " & _
					"left join solicitacao_coleta_has_transportadoras as d " & _
					"on a.idsolicitacao_coleta = d.solicitacao_coleta_idsolicitacao_coleta " & _
					"left join transportadoras as e " & _
					"on d.transportadoras_idtransportadoras = e.idtransportadoras " & _
					"left join cep_consulta_has_clientes as f " & _
					"on c.idclientes = f.clientes_idclientes " & _
					"where a.status_coleta_idstatus_coleta = 2 " & _
					"and b.typecolect = 1 " & _
					"and f.isenderecocoleta = 1 " & _
					"and e.razao_social <> 'NULL' " & _
					"and e.idtransportadoras = " & id
		end if			
				
				
'		response.write sql
'		response.end				
				
		call search(sql, arr, intarr)		
		if intarr > -1 then
			for i=0 to intarr
				getSolicitacaoByTransportadora = "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right"" width=""25%""><b>Cód. Transp:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(0,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Nome Transp:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(1,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Nº Sol:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(2,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Data Envio:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">##/##/####</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Qtd. Prod:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(4,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Razão Social Cliente:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(5,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>End:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(6,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Nº End:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(7,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Bairro:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(8,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>CEP:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(9,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Cidade:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(10,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Estado:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(11,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				if isMaster(id) then
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Contato:</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">"&arr(12,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<tr>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""right""><b>Telefone(Contato):</b></td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "<td align=""left"">("&arr(13,i)&") - "&arr(14,i)&"</td>"
				getSolicitacaoByTransportadora = getSolicitacaoByTransportadora & "</tr>"
				end if
			next
		else
			getSolicitacaoByTransportadora = "<tr>"			
			getSolicitacaoByTransportadora = "<td colspan=""12""><b>Nenhum registro encontrado</b></td>"			
			getSolicitacaoByTransportadora = "</tr>"			
		end if
	end function
	
	function isMaster(id)
		dim sql, arr, intarr, i
		sql = "select ismaster from solicitacao_coleta where numero_solicitacao_coleta = '"&id&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cint(arr(0,i)) = 0 then
					isMaster = false
				else
					isMaster = true
				end if
			next
		end if
	end function

	
%>
<html>
<head>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../css/geral.css" rel="stylesheet" type="text/css">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td id="conteudo" align="left">
					<div style="overflow:scroll;width:600px;height:317px;">
						<form action="frmgerafiletranspanalitica.asp" name="frmgerafiletranspanalitica" method="POST">
							<table cellpadding="1" align="left" cellspacing="1" width="100%" id="tableGetClientesCadastro">
								<tr>
									<td colspan="2" id="explaintitle" align="center">Solicitação Analítica</td>
								</tr>
								<%=getSolicitacaoByTransportadora(request.querystring("id"))%>
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
