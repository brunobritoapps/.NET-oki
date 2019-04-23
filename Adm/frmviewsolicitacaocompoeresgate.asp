<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	sub getSolicitacao()
		dim sql, arr, intarr, i
		dim style
		style = "class=""classColorRelPar"""
		
		sql = "SELECT distinct(C.[numero_solicitacao]) " & _
				  "FROM [marketingoki2].[dbo].[Solicitacoes_resgate_Clientes] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_Coleta] AS B  " & _
				  "ON A.[idsolicitacao] = B.[idsolicitacao_coleta] " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Solicitacao_Resgate_has_Solicitacao_Composicao] AS C " & _
				  "ON B.[numero_solicitacao_coleta] = C.[numero_resgate] " & _
				  "WHERE A.[numero_solicitacao_geracao] = '"&request.querystring("idsolic")&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				response.write "<tr>"
				response.write "<td "&style&">"&arr(0,i)&"</td>"
				response.write "</tr>"
			next
		else
			response.write "<tr><td colspan=""3"" align=""center"">Nenhum registro encontrado</td></tr>"
		end if		
	end sub
	
	function GetStatusColeta(id)
		Dim sSql, arrStatus, intStatus, i
		sSql = "SELECT " & _
						"[idStatus_coleta], " & _ 
						"[status_coleta] " & _
						"FROM [marketingoki2].[dbo].[Status_coleta] where [idStatus_coleta] = " & id
		Call search(sSql, arrStatus, intStatus)						
		If intStatus > -1 Then
			For i=0 To intStatus
				GetStatusColeta = arrStatus(1,i)
			Next
		End If
	End function
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo" style="height:100%;">
		<form action="" name="frmEditSolicitacaoEntrega" method="POST">
		<table cellpadding="1" cellspacing="1" width="500" align="left" id="tableEditSolicitacaoColetaAdm" border="0">
			<tr>
				<td id="explaintitle" align="center">Solicitações que Compuseram o Resgate</td>
			</tr>
			<tr>
				<td colspan="2">
					<div style="width:648px;height:250px;overflow:scroll;">
						<table cellpadding="1" cellspacing="1" width="635" align="left" id="tableRelSolPendente" border="0">
							<tr>
								<th>Nº Solicitação</th>
							</tr>
							<%call getSolicitacao()%>
						</table>
					</div>
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
