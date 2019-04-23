<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
Dim sSql, rs
Dim lIdGrupos, sDescricao
Dim Mensagem


lIdGrupos = request("IdGrupos")

If Request.ServerVariables("HTTP_METHOD") = "POST" Then
	sDescricao = request("txtDescricao")

	if len(trim(lIdGrupos)) > 0 then
		sSql = "UPDATE Grupos SET Descricao='" & sDescricao & "' WHERE IdGrupos = " & lIdGrupos
		Mensagem = "Grupo atualizado com sucesso"
		oConn.Execute(sSql)
	else
		sSql = "INSERT INTO Grupos (Descricao) VALUES ('" & sDescricao & "')"
		Mensagem = "Grupo inserido com sucesso"
		oConn.Execute(sSql)
	end if
Else
	if len(trim(lIdGrupos)) > 0 then
		sSql = "SELECT * FROM Grupos WHERE idGrupos = " & lIdGrupos
		
		Set rs = Server.CreateObject("Adodb.Recordset")		
		rs.Open sSql, oConn
		
		sDescricao = rs("descricao")
		
		rs.Close
		set rs = nothing
		
	end if
End If

Sub GetGrupos()
	Dim sSql, arrGrupos, intGrupos, i
	Dim sSelected
		
	sSql = "SELECT " & _
					"[idGrupos], " & _ 
					"[descricao] " & _ 
					"FROM [marketingoki2].[dbo].[Grupos]"
						
	Call search(sSql, arrGrupos, intGrupos)
		
	With Response
		If intGrupos > -1 Then
			For i=0 To intGrupos
				.Write "<tr>"					
				If CInt(Request.QueryString("IdGrupos")) = arrGrupos(0,i) Then
					sSelected = "checked"
				Else
					sSelected = ""
				End If
				.Write "<td class='classColorRelPar'><input type='radio' name='radioIntIdGrupos' value='"&arrGrupos(0,i)&"' onClick=window.location.href='frmAgrupamentosAdm.asp?IdGrupos="&arrGrupos(0,i)&"' "&sSelected&"/></td>"
				.Write "<td class='classColorRelPar'>"&arrGrupos(1,i)&"</td>"
				.Write "</tr>"					
			Next
		Else
			.Write "<td colspan='5' align=""center"" class='classColorRelPar'>Nenhum Grupo encontrado</td>"
		End If				
	End With
		
End Sub

Sub GetClientesGrupos(lGrupId)
	Dim sSqlCG, arrClientes, intClientes, i
	Dim sSelected

	With Response
		.Write "<tr style='display:none' id=GrupoChange><td colspan=3>Mudar clientes selecionados para Grupo: "
		.Write "<select name=cmbAgrupar onChange='UpdGrupoCli();' class='select'>"
		.Write "<option value="""">Selecione</option>"

		sSqlCG = "SELECT idGrupos, Descricao FROM [marketingoki2].[dbo].[Grupos] " & _
						 "WHERE idGrupos NOT IN (" & lGrupId & ") " & _
						 "ORDER BY idGrupos"
						 
		Call search(sSqlCG, arrClientes, intClientes)

		For i=0 To intClientes
			.Write "<option value="&arrClientes(0,i)&">"&arrClientes(1,i)&"</option>"
		Next
				
		.Write "</select>"
		.Write "</td></tr>"
		.Write "<tr><td colspan=3 class='classColorRelPar'>&nbsp;</td></tr>"
	end with
		
	sSqlCG = "SELECT " & _
					"[idClientes], " & _ 
					"[nome_fantasia], " & _ 					
					"[cnpj] " & _
					"FROM [marketingoki2].[dbo].[Clientes] " & _
					"WHERE Grupos_idGrupos = " & lGrupId

	If Request.QueryString("find") Then		
		sSqlCG = sSqlCG & GetWhereClientes()
	End If

	Call search(sSqlCG, arrClientes, intClientes)
		
	With Response
		If intClientes > -1 Then
			'PAGINACAO NOVA - JADILSON
			Dim intUltima, _
			    intNumProds, _
					intProdsPorPag, _
					intNumPags, _
					intPag, _
					intPorLinha

			intProdsPorPag = 30 'numero de registros mostrados na pagina
			intNumProds = UBound(arrClientes, 2) + 1 'numero total de registros
			
			intPag = CInt(Request("pg")) 'pagina atual da paginacao
			If intPag <= 0 Then intPag = 1
			if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1
			
			intUltima   = intProdsPorPag * intPag - 1
			If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
				
			intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
			If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
			.Write "<tr><td colspan=10>"
			.Write PaginacaoExibir(intPag, intProdsPorPag, intClientes)
			.Write "</td></tr>"
	
			For i = (intProdsPorPag * (intPag - 1)) to intUltima
			'For i=0 To intClientes
				.Write "<tr>"					
				.Write "<td class='classColorRelPar'><input type='checkbox' name='chkIdClientes' value='"&arrClientes(0,i)&"' onClick='ShowCombo(this)' /></td>"
				.Write "<td class='classColorRelPar'>"&arrClientes(1,i)&"</td>"
				.Write "<td class='classColorRelPar'>"&arrClientes(2,i)&"</td>"
				.Write "</tr>"					
			Next
		Else
			.Write "<td colspan='5' align='center' class='classColorRelPar'>Nenhum Cliente encontrado</td>"
		End If				
	End With
		
End Sub

Sub GetCategories()
	Dim sSql, arrCategories, intCategories, i
		
	sSql = "SELECT idCategorias, descricao FROM Categorias"
		
	Call search(sSql, arrCategories, intCategories)
		
	For i=0 To intCategories
		Response.Write "<option value='"&arrCategories(0,i)&"'>"&arrCategories(1,i)&"</option>"
	Next
		
End Sub

Function GetWhereClientes()
	Dim sSqlWhere
	'------------------
	sSqlWhere = ""
		
	Dim Busca
	Dim Por
	Dim Status
	Dim Categorias
	Dim sStrConcat
		
	Dim bSetBusca
	Dim bSetStatus
		
	sStrConcat = " AND "
		
	Busca 		 = Request.QueryString("search")
	Por 			 = Request.QueryString("changetype")
	Status 		 = Request.QueryString("status")
	Categorias = Request.QueryString("categoria")
	'------------------
		
	If Len(Busca) = 0 And CInt(Status) = -1 And CInt(Categorias) = -1 Then
		sSqlWhere = ""
	Else
		sSqlWhere = sSqlWhere & " AND "
	End If
		
	If Len(Busca) > 0 Then
		bSetBusca = True
		Select Case CInt(Por)
			Case 0			
				sSqlWhere = sSqlWhere & " [cnpj] LIKE '%"&Busca&"%' "
			Case 1
				sSqlWhere = sSqlWhere & " [nome_fantasia] LIKE '%"&Busca&"%' "
			Case 2
				sSqlWhere = sSqlWhere & " [razao_social] LIKE '%"&Busca&"%' "	
			Case -1
				sSqlWhere = sSqlWhere & " [cnpj] LIKE '%"&Busca&"%' OR"
				sSqlWhere = sSqlWhere & " [nome_fantasia] LIKE '%"&Busca&"%' AND"
				sSqlWhere = sSqlWhere & " [razao_social] LIKE '%"&Busca&"%' "	
		End Select	
	End If
	If CInt(Status) > -1 Then
		bSetStatus = True
		If bSetBusca Then
			sSqlWhere = sSqlWhere & sStrConcat
		End If
		sSqlWhere = sSqlWhere & " [status_cliente] = " & Status
	End If
	If CInt(Categorias) > -1 Then
		If bSetStatus Then
			sSqlWhere = sSqlWhere & sStrConcat
		End If
		sSqlWhere = sSqlWhere & " [Categorias_idCategorias] = " & Categorias
	End If
		
	GetWhereClientes = sSqlWhere
		
End Function
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script language="Javascript">
function UpdGrupoCli()
{
	document.frmGrupoCli.submit();
}

function ShowCombo(me){
	Objeto = document.getElementById('GrupoChange');	
	
	//alert(me.name);
	
	if (me.checked == true){
		Objeto.style.display="";
		//alert("mostra");
	}
	
	if (me.checked == false){
		Objeto.style.display="none";
		//alert("esconde");
	}	
}

function windowLocationFind(url) {
	var find = true;
	var busca = document.frmSubCliAdm.txtFindCliente.value;
	var por = document.frmSubCliAdm.typeFindCliente.value;
	var status = document.frmSubCliAdm.cbStatusFindCliente.value;
	var categoria = document.frmSubCliAdm.cbCategoriasFindCliente.value;
	window.location.href = url + "/adm/frmAgrupamentosAdm.asp?IdGrupos=<%=lIdGrupos%>&find=true&search="+busca+"&changetype="+por+"&status="+status+"&categoria="+categoria;
}
</script>

</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775" border=0>
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<form action="#" name="frmCategoriasAdm" method="POST">
						<table cellpadding="3" cellspacing="4" width="100%" id="tableLoginCliente">
							<tr>
								<td colspan="2" id="explaintitle" align="center">Cadastro/Alteração de Agrupamentos
							</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
							<tr>
								<td align="right" width="25%">Descrição:</td>
								<td align="left"><input type="text" class="text" name="txtDescricao" value="<%=sDescricao%>" size="40" /></td>
							</tr>
							<tr>
								<td align="right" colspan=2>
									<input type="submit" class="btnform" name="btnSubmitLogin" value="<%if len(trim(lIdGrupos)) > 0 then Response.Write "Editar" else Response.Write "Adicionar"%>" />&nbsp;&nbsp;
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
										</tr>
										<%Call GetGrupos()%>
										<tr>
											<th colspan="2" height="15"></th>
										</tr>
									</table>
								</td>
							</tr>
							</form>
							<%if len(trim(lIdGrupos)) > 0 then%>
							<form action="#" name="frmSubCliAdm" method="GET">
							<tr id="findcontato" valign="baseline">
								<td align="left" valign="baseline" colspan=3>
									Busca: <input type="text" class="text" name="txtFindCliente" value="" />
									&nbsp;&nbsp;&nbsp;
									Por:
									<select name="typeFindCliente" class="select">
										<option value="-1">Selecione</option>
										<option value="0">CNPJ</option>
										<option value="1">Nome Fantasia</option>
										<option value="2">Razão Social</option>
									</select>									
									Status: 
									<select name="cbStatusFindCliente" class="select">
										<option value="-1">Todos</option>
										<option value="1">Ativo</option>
										<option value="0">Inativo</option>
									</select>
									&nbsp;&nbsp;
									Categorias: 
									<select name="cbCategoriasFindCliente" class="select">
										<option value="-1">Todas</option>
										<%Call GetCategories()%>
									</select>
									&nbsp;&nbsp;
									<input type="button" class="btnform" name="btnFindCliente" value="Buscar" onClick="windowLocationFind('<%=URL%>')" />
								</td>
							</tr>
							</form>
							<%end if%>
							<tr>
								<td colspan="3">
									<table cellpadding="1" cellspacing="1" width="100%" align="center" id="tableRelCategories">
										<tr>
											<th width="3%"><img src='img/check.gif' alt='Selecionar' /></th>
											<th>Razão Social</th>
											<th>CNPJ</th>
										</tr>										
										<form name="frmGrupoCli" action="frmUpdGrupoCli.asp" method="GET">
										<%if len(trim(lIdGrupos)) > 0 then call GetClientesGrupos(lIdGrupos)%>										
										</form>
										<tr>
											<th colspan="3" height="15"></th>
										</tr>
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
