<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim ID
	Dim Desc
	Dim Grupo
	Dim bBonus
	Dim Action
	Dim IDOki

	Sub Submit()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call Requests()
			If Action = "edit" Then
				Call Updates()
				Desc = ""
				bBonus = 0
				Grupo = -1
				IDOki = ""
			Else
				Call Insert()
				Desc = ""
				bBonus = 0
				Grupo = -1
				IDOki = ""
			End If	
		ElseIf Request.QueryString("action") = "edit" Then
			Call GetProdutosById(Request.QueryString("id"))
		End If
	End Sub
	
	Sub Requests()
		ID = Request.Form("id")
		Desc = Request.Form("txtDesc")
		bBonus = Request.Form("cbBool")
		Grupo = Request.Form("cbGrupos")
		IDOki = Request.Form("txtID")
		Action = Request.Form("action")
	End Sub
	
	Sub Insert()
		Dim sSql
		if checkID(IDOki) then
			sSql = "INSERT INTO [marketingoki2].[dbo].[Produtos] " & _
				   "([Grupo_produtos_idGrupo_produtos] " & _
				   ",[descricao] " & _
				   ",[gera_bonus] " & _
				   ",[IDOki]) " & _
				   "VALUES " & _
				   "("&Grupo&" " & _
				   ",'"&Desc&"' " & _
				   ","&bBonus&" " & _
				   ",'"&IDOki&"')"
			Call exec(sSql)
		else
			response.write "<script>alert('ID do Produto já cadastrado')</script>"
		end if
	End Sub
	
	Sub Updates()
		Dim sSql	
		if request.form("IDOKI") <> IDOki then
			if checkID(IDOki) then
				sSql = "UPDATE [marketingoki2].[dbo].[Produtos] " & _
					   "SET [Grupo_produtos_idGrupo_produtos] = "&Grupo&" " & _
					   ",[descricao] = '"&Desc&"' " & _
					   ",[gera_bonus] = "&bBonus&" " & _
					   ",[IDOki] = '"&IDOki&"' " & _
					   "WHERE [IDOki] = '" & ID & "'" 
				Call exec(sSql)
			else	
				response.write "<script>alert('ID do Produto já cadastrado')</script>"
			end if
		else
			sSql = "UPDATE [marketingoki2].[dbo].[Produtos] " & _
				   "SET [Grupo_produtos_idGrupo_produtos] = "&Grupo&" " & _
				   ",[descricao] = '"&Desc&"' " & _
				   ",[gera_bonus] = "&bBonus&" " & _
				   ",[IDOki] = '"&IDOki&"' " & _
				   "WHERE [IDOki] = '" & ID & "'" 
			Call exec(sSql)
		end if					
	End Sub
	
'	Sub Deletes(ID)
'		Dim sSql
'		sSql = "DELETE FROM [marketingoki2].[dbo].[Produtos] " & _
'			   "WHERE [IDOki] = '" & ID & "'"
'		Call exec(sSql)
'		Response.Redirect "frmOperacionalAdm.asp"	   
'	End Sub
	
	Sub GetGrupos()
		Dim sSql, arrProd, intProd, i
		Dim sSelected
		sSql = "SELECT [idGrupo_produtos] " & _
			    ",[descricao] " & _
				"FROM [marketingoki2].[dbo].[Grupo_produtos]"
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				If Grupo = arrProd(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value="&arrProd(0,i)&" "&sSelected&">"&arrProd(1,i)&"</option>"
			Next
		End If				
	End Sub
	
	Sub GetProdutos()
		Dim sSql , arrProd, intProd, i
		sSql = "SELECT " & _
					  "A.[Grupo_produtos_idGrupo_produtos] " & _
					  ",A.[descricao] " & _
					  ",A.[gera_bonus] " & _
					  ",B.[descricao] " & _	
					  ",A.[IDOKi] " & _
					  "FROM [marketingoki2].[dbo].[Produtos] AS A " & _
					  "LEFT JOIN [marketingoki2].[dbo].[Grupo_produtos] AS B " & _ 
					  "ON A.[Grupo_produtos_idGrupo_produtos] = B.[idGrupo_produtos]"
		
		if Trim(Request.QueryString("cbGrupo")) <> "-1"  and Trim(Request.QueryString("cbGrupo")) <> "" or len(Trim(Request.QueryString("idoki"))) > 0 then
			sSql = sSql & getWhereSQL()
		end if	
		
		Call search(sSql, arrProd, intProd)

		If intProd > -1 Then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = intProd+1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.ServerVariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			Response.Write "<tr><td colspan=9><div id=pag>"
			Response.Write PaginacaoExibir(intPag, intProdsPorPag, intProd)
			Response.Write "</div></td></tr>"
			
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Editar"" onClick=""window.location.href='frmCadProdutosAdm.asp?id="&arrProd(4,i)&"&action=edit'"" /></td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(4,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(3,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrProd(1,i)&"</td>"
					If arrProd(2,i) = 1 Then
						Response.Write "<td class='classColorRelPar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelPar'>Não</td>"
					End If	
					Response.Write "</tr>"				
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'><img src=""img/buscar.gif"" class=""imgexpandeinfo"" alt=""Editar"" onClick=""window.location.href='frmCadProdutosAdm.asp?id="&arrProd(4,i)&"&action=edit'"" /></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(4,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(3,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrProd(1,i)&"</td>"
					If arrProd(2,i) = 1 Then
						Response.Write "<td class='classColorRelImpar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelImpar'>Não</td>"
					End If	
					Response.Write "</tr>"				
				End If
			Next
			Response.Write "<tr><td colspan=9><div id=pag>"
			Response.Write PaginacaoExibir(intPag, intProdsPorPag, intProd)
			Response.Write "</div></td></tr>"
		Else
			Response.Write "<tr><td colspan=""5"" align=""center"" class=""classColorRelPar""><b>Nenhum Produto encontrado!</b></td></tr>"	
		End If	   
	End Sub
	
	Sub GetProdutosById(ID)
		Dim sSql , arrProd, intProd, i
		sSql = "SELECT " & _
					  "A.[Grupo_produtos_idGrupo_produtos] " & _
					  ",A.[descricao] " & _
					  ",A.[gera_bonus] " & _
					  ",A.[IDOKi] " & _
					  "FROM [marketingoki2].[dbo].[Produtos] AS A " & _
					  "LEFT JOIN [marketingoki2].[dbo].[Grupo_produtos] AS B " & _ 
					  "ON A.[Grupo_produtos_idGrupo_produtos] = B.[idGrupo_produtos] " & _
					  "WHERE A.[IDOKi] = '" & ID & "'"
'		Response.Write sSql
'		Response.End()				
		Call search(sSql, arrProd, intProd)
		If intProd > -1 Then
			For i=0 To intProd
				Grupo = arrProd(0,i)
				Desc = arrProd(1,i)
				bBonus = arrProd(2,i)
				IDOki = arrProd(3,i)
			Next
		End If	   
	End Sub
	
	function getWhereSQL()
		dim sql
		dim bAnd
		sql = ""
		if Trim(Request.QueryString("cbGrupo")) <> "-1" and Trim(Request.QueryString("cbGrupo")) <> "" or len(Trim(Request.QueryString("idoki"))) > 0 then
			sql = sql & " where "
			if Trim(Request.QueryString("cbGrupo")) <> "-1" then
				sql = sql & " A.[Grupo_produtos_idGrupo_produtos] = " & Trim(Request.QueryString("cbGrupo"))
				bAnd = true
			end if
			if len(Trim(Request.QueryString("idoki"))) > 0 then
				if bAnd then
					sql = sql & " and A.[IDOKi] = '" & Trim(Request.QueryString("idoki")) & "'"
				else
					sql = sql & " A.[IDOKi] = '" & Trim(Request.QueryString("idoki")) & "'"
				end if
			end if
		end if
		getWhereSQL = sql
	end function
	
	function checkID(id)
		dim sql, arr, intarr, i
		sql = "SELECT * FROM [marketingoki2].[dbo].[Produtos] where [IDOki] = '"&id&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			checkID = false
		else
			checkID = true
		end if
	end function
	
	Call Submit()
%>
<html>
<head>
<script src="js/frmCadProdutosAdm.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775" >
		<form action="frmCadProdutosAdm.asp" name="frmCadProdutosAdm" method="POST">
		<input type="hidden" name="id" value="<%=Request.QueryString("id")%>" />
		<input type="hidden" name="action" value="<%=Request.QueryString("action")%>" />
		<input type="hidden" name="IDOKI" value="<%=IDOki%>" />
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif" >&nbsp;</td> 
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableprodcad">
						<tr>
							<td colspan="2" id="explaintitle" align="center">Cadastro de Produtos</td>
						</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td align="right">ID: </td>
							<td align="left"><input name="txtID" class="text" value="<%=IDOki%>" type="text" size="10" /></td>
						</tr>
						<tr>
							<td align="right">Grupos: </td>
							<td align="left">
								<select name="cbGrupos" class="select">
									<option value="-1">[Selecione]</option>
									<%Call GetGrupos()%>
								</select>
								<img src="img/grupo_produtos.gif" width="20" height="20" class="imgexpandeinfo" align="absmiddle" onClick="window.open('frmSearchGroupProduto.asp','','width=500,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
							</td>
						</tr>
						<tr>
							<td align="right">Descrição: </td>
							<td align="left"><input type="text" name="txtDesc" class="text" value="<%=Desc%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right">Gera Bônus: </td>
							<td align="left">
								<select name="cbBool" class="select">
									<option value="0" <%If CInt(bBonus) = 0 Then%>selected<%End If%>>Não</option>
									<option value="1" <%If CInt(bBonus) = 1 Then%>selected<%End If%>>Sim</option>
								</select>
							</td>
						</tr>
						<%If Request.QueryString("action") = "edit" Then%>
						<tr>
							<td colspan="2" align="center"><input type="button" class="btnform" name="btnEditar" value="Editar" onClick="validate()" /></td>
						</tr>
						<%Else%>
						<tr>
							<td colspan="2" align="center"><input type="button" class="btnform" name="btnSalvar" value="Salvar" onClick="validate()" /></td>
						</tr>
						<%End If%>
						</form>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<form action="#" name="form2" method="get">
						<tr>
							<td colspan="2" id="explaintitle">
								&nbsp;&nbsp;
								Grupo:
								<select name="cbGrupo" style="width:130px;" class="select">
									<option value="-1">[Selecione]</option>
									<%call GetGrupos()%>
								</select>
								&nbsp;&nbsp;
								ID Produto:
								<input type="text" class="text" name="idoki" />
								<input type="submit" class="btnform" name="buttonProcurar" value="Procurar" />
							</td>
						</tr>
						</form>
						<tr>
							<td colspan="2">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableRelCategories">
									<tr>
										<th width="3%"><img src="img/check.gif"></th>
										<th width="5%">ID</th>
										<th width="30%">Grupo</th>
										<th>Descrição</th>
										<th width="15%">Gera Bônus</th>
									</tr>
									<%Call GetProdutos()%>
								</table>
							</td>
						</tr>
					</table>
				</td>
				<td width="11" background="img/Bg_LatDir.gif">&nbsp;</td>
			</tr>
		</table>
	</div>
	<!--#include file="inc/i_bottom.asp" -->
</div>
</body>
</html>
<%Call close()%>
