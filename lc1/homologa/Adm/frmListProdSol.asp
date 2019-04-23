<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%
	Dim ID
	Dim Produtos
	Dim	ProdutosLength
	ID = Request.QueryString("idsol")
	
	Sub GetProdutos()
		Dim sSql, arrProd, intProd, i
		sSql = "SELECT " & _
					 "[idProdutos], " & _ 
					 "[descricao] " & _ 
					 "FROM [marketingoki2].[dbo].[Produtos]"	
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				Response.Write "<option value='"&arrProd(0,i)&"/"&arrProd(1,i)&"'>"&arrProd(1,i)&"</option>"
			Next
		End If			 
	End Sub
	
	Sub Selecionar()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call Requests()
	
			If Request.Form("qtdrec") = "" Then
				Response.Write "<script>alert('Por favor preencha a quantidade de cartuchos recebidos!');window.close();</script>"
			End If
	
			Dim i
			If Request.ServerVariables("HTTP_METHOD") = "POST" Then
				For i=0 To Ubound(Produtos)
					Response.Write "<tr>"
					Response.Write "<td style=""border-bottom:1px dotted #333333;"" width=""8%""><img src=""img/check.gif"" class=""imgexpandeinfo"" alt="&Mid(Produtos(i),1,InStr(1, Produtos(i), "/",1)-1)&"></td>"
					Response.Write "<td style=""border-bottom:1px dotted #333333;"">"&Mid(Produtos(i), InStr(1, Produtos(i), "/",1)+1)&" <input type=""text"" size=""5"" class=""text"" name="&i&" id="&i&" value="""" /></td>"
					Response.Write "</tr>"
				Next
			Else
				Response.Write "<tr><td style=""border-bottom:1px dotted #333333;"" colspan=""2"">Nenhum produto selecionado</td></tr>"	
			End If	
		End If	
	End Sub
	
	Sub Requests()
		Produtos = Request.Form("listProdutos")
		Produtos = Split(Produtos, ",")
		ProdutosLength = Ubound(Produtos)
		ProdutosLength = ProdutosLength + 1
	End Sub
	
	Sub Adicionar()
		Call Requests()

		Dim i, sSql

		sSql = "DELETE FROM [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos] " & _
			   "WHERE [Solicitacao_coleta_idSolicitacoes_coleta] = " & Request.Form("id")
		Call exec(sSql)			   
		
		For i=0 To Ubound(Produtos)
			sSql = "INSERT INTO [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos]( " & _
							"[Produtos_idProdutos], " & _
							"[Solicitacao_coleta_idSolicitacoes_coleta], " & _
							"[quantidade]) " & _
							"VALUES( " & _
							""&Mid(Produtos(i),1,InStr(1, Produtos(i), "/",1)-1)&", " & _ 
							""&Request.Form("id")&", " & _
							""&arrProd(i)&")"
'			Response.Write sSql
'			Response.End()				
			Call exec(sSql)							
		Next
		Response.Write "<script>window.close();</script>"
	End Sub
	
	Sub GetProdutosBySol()
		Dim sSql, arrProd, intProd, i
		Dim ID
		
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			ID = Request.Form("id")
		Else
			ID = Request.QueryString("idsol")
		End If	
		
		sSql = "SELECT A.[Produtos_idProdutos] " & _
				  ",A.[Solicitacao_coleta_idSolicitacoes_coleta] " & _
				  ",A.[quantidade] " & _
				  ",B.[descricao] " & _
				  "FROM [marketingoki2].[dbo].[Solicitacoes_coleta_has_Produtos] AS A " & _
				  "LEFT JOIN [marketingoki2].[dbo].[Produtos] AS B " & _
				  "ON A.[Produtos_idProdutos] = B.[idProdutos] " & _
				  "WHERE A.[Solicitacao_coleta_idSolicitacoes_coleta] = " & ID

'		Response.Write sSql
'		Response.End()		  
				  
		Call search(sSql, arrProd, intProd)				  
		If intProd > -1 Then
			For i=0 To intProd
				Response.Write "<tr>"
				Response.Write "<td style=""border-bottom:1px dotted #333333;"" width=""8%""><img src=""img/check.gif"" class=""imgexpandeinfo"" alt=""Excluir"" onClick=""javascript:window.location.href='frmListProdSol.asp?idsol="&Request.QueryString("idsol")&"&idprod="&arrProd(0,i)&"';"" /></td>"
				Response.Write "<td style=""border-bottom:1px dotted #333333;"">"&arrProd(3,i)&"</td>"
				Response.Write "</tr>"
			Next	
		Else
			Response.Write "<tr><td style=""border-bottom:1px dotted #333333;"" colspan=""2"">Nenhum produto encontrado</td></tr>"	
		End If
	End Sub
	
	Sub RequestsTxtQtd()
		Dim i

		For i=0 To Request.Form("produtoslength")
			arrProd(i) = Request.Form(i)	
		Next
	End Sub
	
	Sub Submit()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			If Request.Form("action") = "ADD" Then
				arrProd(Request.Form("produtoslength"))
				Call RequestsTxtQtd()
				Call Adicionar()
			End If
		End If
	End Sub
	
'	Call Requests()
	
%>
<html>
<head>
<script src="js/frmListProdSol.js" language="javascript" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<div id="conteudo">
		<form action="frmListProdSol.asp" name="frmListProdSol" method="POST">
		<input type="hidden" name="id" value="<%=ID%>" />
		<input type="hidden" name="action" value="" />
		<input type="hidden" name="qtdrec" value="<%If Request.QueryString("qtdrec") = "" Then Response.Write Request.Form("qtdrec") Else Response.Write Request.QueryString("qtdrec") End If %>" />
		<input type="hidden" name="produtoslength" value="<%= ProdutosLength %>" />
		<table cellspacing="0" cellpadding="0" width="350" align="left">
			<tr> 
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" width="100%" align="left">
						<tr>
							<td id="explaintitle" colspan="3" align="center">Produtos da Solicitação</td>
						</tr>
						<tr>
							<td align="center">
								<select name="listProdutos" class="select" multiple="multiple" style="width:300px;height:200px;">
									<%Call GetProdutos()%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<input type="button" class="btnform" name="select" value="Selecionar" onClick="selecionar()" />
								<input type="button" class="btnform" name="add" value="Adicionar" onClick="adicionar()" />
							</td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<div style="width:100%;overflow:auto;height:173px;">
								<table cellpadding="1" cellspacing="1" width="100%" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">
									<tr>
										<td colspan="2" id="explaintitle" align="center">Produtos Selecionados</td>
									</tr>
									<%Call Selecionar()%>
								</table>
								</div>
							</td>
						</tr>
						<tr>
							<td colspan="3" align="center">
								<div style="width:100%;overflow:auto;height:173px;">
								<table cellpadding="1" cellspacing="1" width="100%" style="font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;">
									<tr>
										<td colspan="2" id="explaintitle" align="center">Produtos Adicionados</td>
									</tr>
									<%Call GetProdutosBySol()%>
								</table>
								</div>
							</td>
						</tr>
					</table>
				</td>
			</tr>
		</table>
		</form>
	</div>
</div>
</body>
</html>
<%Call close()%>
