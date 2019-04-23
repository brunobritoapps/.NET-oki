<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim ID
	Dim Desc
	Dim Action
	dim idGrupo
	
	Sub Submit()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call Requests()
			If Action = "edit" Then
				Call Updates()
				Desc = ""
				idGrupo = ""
			ElseIf Action = "updategroup" Then
				Call AtualizaProd()
				Desc = ""
				idGrupo = ""
			Else
				Call Insert()
				Desc = ""
				idGrupo = ""
			End If	
		ElseIf Request.QueryString("action") = "edit" Then
			Call GetGruposById(Request.QueryString("id"))
		ElseIf Request.QueryString("action") = "delete" Then
			Call Deletes(Request.QueryString("id"))	
		End If
	End Sub
	
	Sub Requests()
		ID = Request.Form("id")
		Desc = Request.Form("txtDesc")
		Action = Request.Form("action")
		idGrupo = Trim(Request.Form("txtID"))
	End Sub
	
	Sub Insert()
		Dim sSql
		sSql = "INSERT INTO [marketingoki2].[dbo].[Grupo_produtos] " & _
			    "([descricao], [idokigrupo]) " & _
				"VALUES " & _
				"('"&Desc&"', '"&idGrupo&"')"
		Call exec(sSql)				
	End Sub
	
	Sub Updates()
		Dim sSql
		sSql = "UPDATE [marketingoki2].[dbo].[Grupo_produtos] " & _
			   "SET [descricao] = '"&Desc&"', [idokigrupo] = '"&idGrupo&"' " & _
			   "WHERE [idGrupo_produtos] = " & ID
		Call exec(sSql)			   
	End Sub
	
	Sub Deletes(ID)
		Dim sSql, arrProd, intProd
		sSql = "SELECT * FROM [marketingoki2].[dbo].[Produtos] WHERE [Grupo_produtos_idGrupo_produtos] = " & ID
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			Response.Write "<script>alert('Este Grupo contém Produtos relacionados, por favor delete os produtos primeiro!')</script>"
		Else	
			sSql = "DELETE FROM [marketingoki2].[dbo].[Grupo_produtos] " & _
				   "WHERE [idGrupo_produtos] = " & ID
			Call exec(sSql)			   
		End If
	End Sub
	
	Sub GetGruposById(ID)
		Dim sSql, arrGrupo, intGrupo, i
		sSql = "SELECT [idGrupo_produtos] " & _
			    ",[descricao] " & _
			    ",[idokigrupo] " & _
				"FROM [marketingoki2].[dbo].[Grupo_produtos] " & _
				"WHERE [idGrupo_produtos] = " & ID
		Call search(sSql, arrGrupo, intGrupo)
		If intGrupo > -1 Then
			For i=0 To intGrupo
				Desc = arrGrupo(1,i)
				idGrupo = arrGrupo(2,i)
			Next
		End If							
	End Sub
	
	Sub GetGrupos()
		Dim sSql, arrProd, intProd, i
		sSql = "SELECT [idGrupo_produtos] " & _
			    ",[descricao] " & _
			    ",[idokigrupo] " & _
				"FROM [marketingoki2].[dbo].[Grupo_produtos]"
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Editar"" onClick=""window.location.href='frmGrupoProdutosAdm.asp?id="&arrProd(0,i)&"&action=edit'"" /></td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Editar"" onClick=""window.location.href='frmGrupoProdutosAdm.asp?id="&arrProd(0,i)&"&action=edit'"" /></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(2,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(1,i)&"</td>"
					Response.Write "</tr>"
				End If
			Next
		Else
			Response.Write "<tr><td colspan=""3"" align=""center"" class=""classColorRelPar""><b>Nenhum Grupo de Produto encontrado!</b></td></tr>"	
		End If				
	End Sub
	
	Sub GetProdutosByGroup(ID)
		Dim sSql, arrProdutos, intProdutos, i
		sSql = "SELECT " & _
						"[IDOki], " & _ 
						"[descricao], " & _ 
						"[gera_bonus] " & _ 
						"FROM [marketingoki2].[dbo].[Produtos] " & _
						"WHERE [Grupo_produtos_idGrupo_produtos] = " & ID
			Call search(sSql ,arrProdutos, intProdutos)
			If intProdutos > -1 Then
				'PAGINACAO NOVA - JADILSON
				Dim intUltima, _
				    intNumProds, _
						intProdsPorPag, _
						intNumPags, _
						intPag, _
						intPorLinha, _
						j

				intProdsPorPag = 30 'numero de registros mostrados na pagina
				intNumProds = UBound(arrProdutos, 2) + 1 'numero total de registros
			
				intPag = CInt(Request("pg")) 'pagina atual da paginacao
				If intPag <= 0 Then intPag = 1
				if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1
			
				intUltima   = intProdsPorPag * intPag - 1
				If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
					
				intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
				If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
				Response.Write "<tr><td colspan=10>"
				Response.Write PaginacaoExibir(intPag, intProdsPorPag, intProdutos)
				Response.Write "</td></tr>"
				
				'if intNumPags = 0 then
				'	Response.Write "<input type=""hidden"" name=""hiddenIntProdutos"" value="&intProdutos + 1&" />" & vbcrlf
				'else
				'	Response.Write "<input type=""hidden"" name=""hiddenIntProdutos"" value="&intProdutos-(intProdsPorPag*intPag)&" />" & vbcrlf
				'end if				
				
				For i = (intProdsPorPag * (intPag - 1)) to intUltima				
					j = j + 1
				'For i=0 To intProdutos
					If i Mod 2 = 0 Then
						Response.Write "<tr>" & vbcrlf
						Response.Write "<td class='classColorRelPar'><input type=""checkbox"" name=""radioIntProduto"" value="""&trim(arrProdutos(0,i))&""" onClick=""showOnClick()"" /></td>" & vbcrlf
						Response.Write "<td class='classColorRelPar'>"&arrProdutos(0,i)&"</td>" & vbcrlf
						Response.Write "<td class='classColorRelPar'>"&arrProdutos(1,i)&"</td>" & vbcrlf
						If arrProdutos(2,i) = 0 Then
							Response.Write "<td class='classColorRelPar'>Não</td>" & vbcrlf
						Else
							Response.Write "<td class='classColorRelPar'>Sim</td>" & vbcrlf
						End If	
						Response.Write "</tr>" & vbcrlf
					Else
						Response.Write "<tr>" & vbcrlf
						Response.Write "<td class='classColorRelImpar'><input type=""checkbox"" name=""radioIntProduto"" value="""&trim(arrProdutos(0,i))&""" onClick=""showOnClick()"" /></td>" & vbcrlf
						Response.Write "<td class='classColorRelImpar'>"&arrProdutos(0,i)&"</td>" & vbcrlf
						Response.Write "<td class='classColorRelImpar'>"&arrProdutos(1,i)&"</td>" & vbcrlf
						If arrProdutos(2,i) = 0 Then
							Response.Write "<td class='classColorRelImpar'>Não</td>" & vbcrlf
						Else
							Response.Write "<td class='classColorRelImpar'>Sim</td>" & vbcrlf
						End If	
						Response.Write "</tr>" & vbcrlf
					End If	
				Next
				Response.Write "<input type=""hidden"" name=""hiddenIntProdutos"" value="&j&" />" & vbcrlf
			Else
				Response.Write "<tr><td colspan=""4"" class='classColorRelPar' align=""center""><b>Nenhum produto encontrado!</b></td></tr>"
			End If			
	End Sub
	
	Sub GetGroups()
		Dim sSql, arrGroup, intGroup, i
		sSql = "SELECT [idGrupo_produtos], [descricao] FROM [marketingoki2].[dbo].[Grupo_produtos]"
		Call search(sSql, arrGroup, intGroup)
		If intGroup > -1 Then
			For i=0 To intGroup
				Response.Write "<option value="&arrGroup(0,i)&">"&arrGroup(1,i)&"</option>"
			Next	
		End If
	End Sub
	
	Sub AtualizaProd()
		Dim Prod
		Dim iGrupo
		Dim i
		Dim sSql
		
		Prod = Split(Request.Form("radioIntProduto"), ",")
		iGrupo = Request.Form("cbGrupos")
		For i=0 To Ubound(Prod)
			sSql = "UPDATE [marketingoki2].[dbo].[Produtos] " & _
							"SET [Grupo_produtos_idGrupo_produtos] = "&iGrupo&" " & _
							"WHERE [IDOki]= '"&Trim(Prod(i))&"' "
			Call exec(sSql)				
		Next
	End Sub
	
	Call Submit()
%>
<html>
<head>
<script language="javascript" src=js/frmGrupoProdutosAdm.js></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775" >
		<form action="frmGrupoProdutosAdm.asp" name="frmGrupoProdutosAdm" method="POST">
		<input type="hidden" name="id" value="<%= Request.QueryString("id") %>" />
		<input type="hidden" name="action" value="<%=Request.QueryString("action")%>" />
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif" >&nbsp;</td> 
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableprodcad">
						<tr>
							<td colspan="2" id="explaintitle" align="center">Cadastro de Grupos de Produtos</td>
						</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td align="right">ID: </td>
							<td align="left"><input type="text" name="txtID" class="text" value="<%=idGrupo%>" size="10" /></td>
						</tr>
						<tr>
							<td align="right">Descrição: </td>
							<td align="left"><input type="text" name="txtDesc" class="text" value="<%=Desc%>" size="30" /></td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<%If Request.QueryString("action") = "edit" Then%>
						<tr>
							<td align="right"><input type="button" class="btnform" name="btnEditar" value="Editar" onClick="validate()" /></td>
							<td align="left"><input type="button" class="btnform" name="btnDeletar" value="Deletar" onClick="window.location.href='frmGrupoProdutosAdm.asp?id=<%=Request.QueryString("id")%>&action=delete'"/></td>
						</tr>
						<%Else%>
						<tr>
							<td colspan="2" align="center"><input type="button" class="btnform" name="btnSalvar" value="Salvar" onClick="validate()" /></td>
						</tr>
						<%End If%>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="2">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableRelCategories">
									<tr>
										<th width="3%"><img src="img/check.gif" class="imgexpandeinfo" /></th>
										<th width="15%">ID OKI</th>
										<th>Descrição</th>
									</tr>
									<%Call GetGrupos()%>
								</table>
							</td>
						</tr>
						<tr id="cbgrupos" style="display:none;">
							<td colspan="2" id="explaintitle">
								<select name="cbGrupos" class="select" onChange="validateChangeListener()">
									<option value="-1">Selecione</option>
									<%Call GetGroups()%>
								</select>
							</td>
						</tr>
						<%If Request.QueryString("action") = "edit" Then%>
						<tr>
							<td colspan="2">
								<table cellspacing="1" cellpadding="1" width="100%" id="tableRelCategories">
									<th width="2%"><img src="img/check.gif" /></th>
									<th width="10%">ID</th>
									<th>Descrição</th>
									<th width="10%">Gera Bônus</th>
									<%Call GetProdutosByGroup(Request.QueryString("id"))%>
								</table>
							</td>
						</tr>
						<%End If%>						
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
