<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionPonto()%>
<%
	sub getSolicitacao()
		dim sql, arr, intarr, i
		dim style
		style = "class=""classColorRelPar"""
		
		sql = "SELECT a.[id_solicitacao] " & _
				  ",a.[id_pontocoleta] " & _
				  ",a.[numero_solicitacao_master] " & _
				  ",a.[is_baixada] " & _
				  ",b.[numero_solicitacao_coleta] " & _
				  ",b.[qtd_cartuchos] " & _
				  ",b.[status_coleta_idstatus_coleta] " & _
			  	"FROM [marketingoki2].[dbo].[Solicitacoes_Baixadas] as a " & _
				"left join solicitacao_coleta as b " & _
				"on a.[id_solicitacao] = b.idsolicitacao_coleta " & _
				"where a.[numero_solicitacao_master] = '"&request.querystring("idsolic")&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if i mod 2 = 0 then
					style = "class=""classColorRelPar"""
				else
					style = "class=""classColorRelImpar"""
				end if
				response.write "<tr>"
				response.write "<td "&style&">"&arr(4,i)&"</td>"
				response.write "<td "&style&">"&GetStatusColeta(arr(6,i))&"</td>"
				response.write "<td "&style&">"&arr(5,i)&"</td>"
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
				<td id="explaintitle" align="center">Administrar Solicitação que Compuseram o Resgate</td>
			</tr>
			<tr>
				<td colspan="2">
					<div style="width:648px;height:250px;overflow:scroll;">
						<table cellpadding="1" cellspacing="1" width="635" align="left" id="tableRelSolPendente" border="0">
							<tr>
								<th>Nº Solicitação</th>
								<th>Status</th>
								<th>Qtd. Cartuchos</th>
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
