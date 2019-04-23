<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim CEP
	Dim Logradouro
	Dim Bairro
	Dim Municipio
	Dim Estado
	Dim TipoLog
	
	Sub GetCEP()
		Dim sSql, arrCEP, intCEP, i

		sSql = "SELECT [idcep_consulta] " & _
			   ",[cep] " & _
			   ",[logradouro] " & _
			   ",[bairro] " & _
			   ",[municipio] " & _
			   ",[estado] " & _
			   ",[tipologradouro] " & _
			   "FROM [marketingoki2].[dbo].[cep_consulta]"

		If Request.QueryString("find") Then

			sSql = sSql & GetWhereFindCEP()
			Call search(sSql, arrCEP, intCEP)
			
			Response.Write "<tr>"
			Response.Write "<th>CEP</th>"
			Response.Write "<th>Tipo Logradouro</th>"	
			Response.Write "<th>Logradouro</th>"			
			Response.Write "<th>Bairro</th>"	
			Response.Write "<th>Município</th>"
			Response.Write "<th>Estado</th>"			
			Response.Write "</tr>"	
			
			If intCEP > -1 Then
				'PAGINACAO NOVA - JADILSON
				Dim intUltima, _
				    intNumProds, _
						intProdsPorPag, _
						intNumPags, _
						intPag, _
						intPorLinha

				intProdsPorPag = 30 'numero de registros mostrados na pagina
				intNumProds = UBound(arrCEP, 2) + 1 'numero total de registros
			
				intPag = CInt(Request("pg")) 'pagina atual da paginacao
				If intPag <= 0 Then intPag = 1
				if request.servervariables("HTTP_METHOD") = "POST" then	intPag=1
			
				intUltima   = intProdsPorPag * intPag - 1
				If intUltima > (intNumProds - 1) Then intUltima = (intNumProds - 1)
					
				intNumPags = (intNumProds - (intNumProds mod intProdsPorPag)) / intProdsPorPag
				If (intNumPags mod intProdsPorPag) > 0 Then intNumPags = intNumPags + 1
		
				Response.Write "<tr><td colspan=10>"
				Response.Write PaginacaoExibir(intPag, intProdsPorPag, intCEP)
				Response.Write "</td></tr>"
	
				For i = (intProdsPorPag * (intPag - 1)) to intUltima
					If i Mod 2 = 0 Then
						Response.Write "<tr>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(1,i))&"</td>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(6,i))&"</td>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(2,i))&"</td>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(3,i))&"</td>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(4,i))&"</td>"
						Response.Write "<td class='classColorRelPar'>"&Trim(arrCEP(5,i))&"</td>"
						Response.Write "</tr>"
					Else
						Response.Write "<tr>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(1,i))&"</td>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(6,i))&"</td>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(2,i))&"</td>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(3,i))&"</td>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(4,i))&"</td>"
						Response.Write "<td class='classColorRelImpar'>"&Trim(arrCEP(5,i))&"</td>"
						Response.Write "</tr>"
					End If
				Next
			Else
				Response.Write "<tr><td colspan=""6"" align=""center""><b>Nenhum CEP encontrado!</b></td></tr>"	
			End If 			   
		End If
	End Sub
	
	Sub GetTipoLog() 
		Dim sSql, arrLog, intLog, i
		
		sSql = "SELECT DISTINCT [tipologradouro] " & _
				"FROM [marketingoki2].[dbo].[cep_consulta] " & _ 
				"ORDER BY [tipologradouro] ASC"
				
		Call search(sSql, arrLog, intLog)		
		If intLog > -1 Then
			For i=0 To intLog
				Response.Write "<option value="&Trim(arrLog(0,i))&">"&Trim(arrLog(0,i))&"</option>"
			Next
		End If
	End Sub
	
	Function GetWhereFindCEP()
		Dim sSqlWhere
		'------------------
		sSqlWhere = ""
		
		Dim Busca
		Dim Por
		Dim sStrConcat
		
		Dim bSetBusca
		Dim bSetStatus
		
		sStrConcat = " AND "
		
		Busca 		 = Request.QueryString("search")
		Por 		 = Request.QueryString("changetype")
		'------------------
		
		If Len(Busca) = 0 And CInt(Por) = -1 Then
			sSqlWhere = ""
		Else
			sSqlWhere = sSqlWhere & " WHERE "
		End If
		
		If Len(Busca) > 0 Then
			bSetBusca = True
			Select Case CInt(Por)
				Case 0			
					sSqlWhere = sSqlWhere & " [CEP] LIKE '%"&Busca&"%' "
				Case 1
					sSqlWhere = sSqlWhere & " [Logradouro] LIKE '%"&Busca&"%' "
				Case 2
					sSqlWhere = sSqlWhere & " [Bairro] LIKE '%"&Busca&"%' "	
				Case 3
					sSqlWhere = sSqlWhere & " [Municipio] LIKE '%"&Busca&"%' "
				Case 4	
					sSqlWhere = sSqlWhere & " [Estado] LIKE '%"&Busca&"%' "	
				Case -1	
					sSqlWhere = sSqlWhere & " [CEP] LIKE '%"&Busca&"%' "
					sSqlWhere = sSqlWhere & " [Logradouro] LIKE '%"&Busca&"%' "
					sSqlWhere = sSqlWhere & " [Bairro] LIKE '%"&Busca&"%' "	
					sSqlWhere = sSqlWhere & " [Municipio] LIKE '%"&Busca&"%' "
					sSqlWhere = sSqlWhere & " [Estado] LIKE '%"&Busca&"%' "	
			End Select	
		End If
		
		GetWhereFindCEP = sSqlWhere
		
	End Function
	
	Sub Submit()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call Requests()
			Call Inserir()
		End If
		CEP 		= ""
		Logradouro  = ""
		Bairro 		= ""
		Municipio 	= ""
		Estado 		= ""
		TipoLog 	= -1
	End Sub
	
	Sub Inserir()
		Dim sSql
		sSql = "INSERT INTO [marketingoki2].[dbo].[cep_consulta] " & _
					   "([cep] " & _
					   ",[logradouro] " & _
					   ",[bairro] " & _
					   ",[municipio] " & _
					   ",[estado] " & _
					   ",[tipologradouro]) " & _
				 "VALUES " & _
					   "('"&CEP&"' " & _
					   ",'"&Logradouro&"' " & _
					   ",'"&Bairro&"' " & _
					   ",'"&Municipio&"' " & _
					   ",'"&Estado&"' " & _
					   ",'"&TipoLog&"')"
		Call exec(sSql)					   
	End Sub
	
	Sub Requests()
		CEP 		= Request.Form("txtCep")
		Logradouro  = Request.Form("txtLogradouro")
		Bairro 		= Request.Form("txtBairro")
		Municipio 	= Request.Form("txtMunicipio")
		Estado 		= Request.Form("txtEstado")
		TipoLog 	= Request.Form("cbTipos")
	End Sub
	
	Call Submit()
%>
<html>
<head>
<script src="js/frmCepEndAdm.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775" >
		<form action="frmCepEndAdm.asp" name="frmCepEndAdm" method="POST">
			<tr>
				<td width="11" background="img/Bg_LatEsq.gif" >&nbsp;</td> 
				<td id="conteudo">
					<table cellpadding="1" cellspacing="1" width="100%" id="tableCadCliente">
						<tr>
							<td colspan="2" id="explaintitle" align="center">Cadastro de CEP´s / Endereços</td>
						</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td align="right">CEP: </td>
							<td align="left"><input name="txtCep" type="text" class="text" value="<%= CEP %>" size="30" maxlength="8" /></td>
						</tr>
						<tr>
							<td align="right">Logradouro: </td>
							<td align="left"><input name="txtLogradouro" class="text" type="text" value="<%= Logradouro %>" size="40" /></td>
						</tr>
						<tr>
							<td align="right">Bairro: </td>
							<td align="left"><input name="txtBairro" type="text" class="text" value="<%= Bairro %>" size="15" maxlength="30" /></td>
						</tr>
						<tr>
							<td align="right">Município: </td>
							<td align="left"><input name="txtMunicipio" type="text" class="text" value="<%= Municipio %>" size="35" maxlength="60" /></td>
						</tr>
						<tr>
							<td align="right">Estado: </td>
							<td align="left"><input name="txtEstado" type="text" class="text" value="<%= Estado %>" size="4" maxlength="2" /></td>
						</tr>
						<tr>
							<td align="right">Tipo Logradouro: </td>
							<td align="left">
								<select name="cbTipos" class="select">
									<option value="-1">- - - Selecione um Tipo - - -</option>
									<%Call GetTipoLog()%>
								</select>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td colspan="2" align="center"><input name="btnSalvar" type="submit" class="btnform" value="Salvar" /></td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;</td>
						</tr>
						<tr>
							<td id="explaintitle" colspan="6" align="center">Busca de CEP</td>
						</tr>
						<tr>
							<td colspan="2">
								<table cellpadding="1" cellspacing="1" width="100%" id="tableRelSolPendente">
									<tr>
										<td colspan="6" align="center">
											Busca: <input name="search" type="text" class="text" size="20" maxlength="20" />
											<select name="cbBusca" class="select">
												<option value="-1"> Todos </option>
												<option value="0">CEP</option>
												<option value="1">Logradouro</option>
												<option value="2">Bairro</option>
												<option value="3">Municipio</option>
												<option value="4">Estado</option>
											</select>
											<input name="buscar" type="button" class="btnform" value="Buscar" onClick="windowLocationFind('<%=URL%>')" />
										</td>
									</tr>
									<%Call GetCEP()%>
								</table>
							</td>
						</tr>
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
