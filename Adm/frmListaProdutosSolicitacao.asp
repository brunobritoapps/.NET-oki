<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Sub GetProductBySol()
		Dim sSql, arrProd, intProd, i
		sSql = "SELECT " & _
						"A.[Produtos_idProdutos], " & _ 
						"A.[quantidade], " & _
						"B.[descricao] " & _ 
						"FROM [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
						"ON B.[IDOki] = A.[Produtos_idProdutos] " & _
						"WHERE A.[Solicitacao_coleta_idSolicitacoes_coleta] = " & Request.QueryString("idsol")
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(0,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(0,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				End If	
			Next
		Else
			Response.Write "<tr><td colspan=""3"" class='classColorRelPar' align=""center""><b>Material ainda não recebido</b></td></tr>"	
		End If				
	End Sub
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<form action="frmListaProdutosSolicitacao.asp" name="frmListaProdutosSolicitacao" method="POST">
		<table cellspacing="0" cellpadding="0" width="600" align="left">
			<tr> 
				<td id="conteudo">
					<div style="width:100%;overflow:scroll;height:315px;">
					<table cellpadding="1" cellspacing="1" width="100%" align="left" id="tableRelSolPendente">
						<tr>
							<td id="explaintitle" colspan="3" align="center">Produtos da Solicitação</td>
						</tr>
						<tr>
							<th>Cód. Produto</th>
							<th>Descrição Produto</th>
							<th>Quantidade</th>
						</tr>
						<%Call GetProductBySol()%>
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
