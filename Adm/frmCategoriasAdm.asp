<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
Dim sSql, rs
Dim lIdCategorias, sDescricao, bAtivo, bColeta, lQtdMinima
Dim Mensagem

lIdCategorias = request("IdCategorias")

If Request.ServerVariables("HTTP_METHOD") = "POST" Then
	sDescricao = request("txtDescricao")
	bAtivo = request("cmbAtivo")
	bColeta = request("cmbColeta")
	lQtdMinima = request("txtQtdMinima")

	if len(trim(lIdCategorias)) > 0 then
		sSql = "UPDATE Categorias SET Descricao='" & sDescricao & "', Ativo=" & bAtivo & ", isColetaDomiciliar=" & bColeta & ", minCartuchos=" & lQtdMinima & " WHERE IdCategorias = " & lIdCategorias
		Mensagem = "Categoria atualizada com sucesso"
		oConn.Execute(sSql)
	else
		sSql = "INSERT INTO Categorias (Descricao, Ativo, isColetaDomiciliar, minCartuchos) VALUES ('" & sDescricao & "', " & bAtivo & ", " & bColeta & ", " & lQtdMinima & ")"
		Mensagem = "Categoria inserida com sucesso"
		oConn.Execute(sSql)
	end if
Else
	if len(trim(lIdCategorias)) > 0 then
		sSql = "SELECT * FROM Categorias WHERE idCategorias = " & lIdCategorias
		
		Set rs = Server.CreateObject("Adodb.Recordset")		
		rs.Open sSql, oConn
		
		sDescricao = rs("descricao")
		bAtivo = rs("ativo")
		bColeta = rs("isColetaDomiciliar")
		lQtdMinima = rs("minCartuchos")
		
		rs.Close
		set rs = nothing
		
	end if
End If

	Sub GetCategories()
		Dim sSql, arrCategories, intCategories, i
		Dim sSelected
		
		sSql = "SELECT " & _
						"[idCategorias], " & _ 
						"[descricao], " & _ 
						"[ativo], " & _ 
						"[isColetaDomiciliar], " & _ 
						"[minCartuchos] " & _ 
						"FROM [marketingoki2].[dbo].[Categorias]"
						
		Call search(sSql, arrCategories, intCategories)
		
		With Response
			If intCategories > -1 Then
				For i=0 To intCategories
					If i Mod 2 = 0 Then
						.Write "<tr>"					
						If CInt(Request.QueryString("IdCategorias")) = arrCategories(0,i) Then
							sSelected = "checked"
						Else
							sSelected = ""
						End If
						.Write "<td class='classColorRelPar'><input type='radio' name='radioIntIdCategories' value='"&arrCategories(0,i)&"' onClick=window.location.href='frmCategoriasAdm.asp?IdCategorias="&arrCategories(0,i)&"' "&sSelected&"/></td>"
						.Write "<td class='classColorRelPar'>"&arrCategories(1,i)&"</td>"
						If arrCategories(2,i) = 1 Then
							.Write "<td class='classColorRelPar'>Ativo</td>"
						Else 
							.Write "<td class='classColorRelPar'>Inativo</td>"
						End If
						If arrCategories(3,i) = 1 Then
							.Write "<td class='classColorRelPar'>Sim</td>"
						Else
							.Write "<td class='classColorRelPar'>Não</td>"
						End If
						.Write "<td class='classColorRelPar'>"&arrCategories(4,i)&"</td>"
						.Write "</tr>"					
					Else
						.Write "<tr>"					
						If CInt(Request.QueryString("IdCategorias")) = arrCategories(0,i) Then
							sSelected = "checked"
						Else
							sSelected = ""
						End If
						.Write "<td class='classColorRelImpar'><input type='radio' name='radioIntIdCategories' value='"&arrCategories(0,i)&"' onClick=window.location.href='frmCategoriasAdm.asp?IdCategorias="&arrCategories(0,i)&"' "&sSelected&"/></td>"
						.Write "<td class='classColorRelImpar'>"&arrCategories(1,i)&"</td>"
						If arrCategories(2,i) = 1 Then
							.Write "<td class='classColorRelImpar'>Ativo</td>"
						Else 
							.Write "<td class='classColorRelImpar'>Inativo</td>"
						End If
						If arrCategories(3,i) = 1 Then
							.Write "<td class='classColorRelImpar'>Sim</td>"
						Else
							.Write "<td class='classColorRelImpar'>Não</td>"
						End If
						.Write "<td class='classColorRelImpar'>"&arrCategories(4,i)&"</td>"
						.Write "</tr>"					
					End If
				Next
			Else
				.Write "<td colspan='5' align=""center"" class=""classColorRelPar""><b>Nenhuma Categoria encontrada</b></td>"		
			End If				
		End With
		
	End Sub

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<script src="js/frmCategoriasAdm.js"></script>
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<form action="#" name="frmCategoriasAdm" method="POST">
						<table cellpadding="3" cellspacing="4" width="100%" id="tableLoginCliente" border="0">
							<tr>
								<td colspan="2" id="explaintitle" align="center">Cadastro/Alteração de Categorias</td>
							</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
							<tr>
								<td align="right" width="25%">Descrição:</td>
								<td align="left"><input type="text" class="text" name="txtDescricao" value="<%=sDescricao%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" width="25%">Ativo:</td>
								<td align="left">
									<select name="cmbAtivo" class="select">
										<option value="1"<%if bAtivo = 1 then response.write " selected"%>>Sim</option>
										<option value="0"<%if bAtivo = 0 then response.write " selected"%>>Não</option>
									</select>
								</td>
							</tr>
							<tr>
								<td align="right" width="25%">É Coleta Domiciliar:</td>
								<td align="left">
									<select name="cmbColeta" class="select">
										<option value="1"<%if bColeta = 1 then response.write " selected"%>>Sim</option>
										<option value="0"<%if bColeta = 0 then response.write " selected"%>>Não</option>
									</select>
								</td>
							</tr>
							<tr>
								<td align="right" width="25%">Quantidade Mínima:</td>
								<td align="left"><input type="text" class="text" name="txtQtdMinima" value="<%=lQtdMinima%>" size="5" /></td>
							</tr>
							<tr>
								<td align="right" colspan=2>
									<input type="button" class="btnform" name="btnSubmitLogin" value="<%if len(trim(lIdCategorias)) > 0 then Response.Write "Editar" else Response.Write "Adicionar"%>" onClick="validate()" />&nbsp;&nbsp;
									<input type="reset" class="btnform" name="btnLimpaForm" value="Limpar" />
								</td>
							</tr>
							<%If Mensagem <> "" Then%>
								<tr>
									<td align="center" colspan=2><b style="color:#FF0000;"><%=Mensagem%></b></td>
								</tr>
							<%End If%>	
							<tr id="explaintitle">
							<tr>
								<td colspan="3">
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelCategories">
										<tr>
											<th width="3%"><img src='img/check.gif' alt='Selecionar' /></th>
											<th>Descrição</th>
											<th width="10%">Ativo</th>
											<th width="15%">Coleta Domiciliar</th>
											<th width="15%">Mínimo de Cartuchos</th>
										</tr>
										<%Call GetCategories()%>
										<tr>
											<th colspan="5" height="15"></th>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</form>
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
