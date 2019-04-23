<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim ID
	Dim RazaoSocial
	Dim NomeFantasia
	Dim	CNPJ
	Dim TipoBonus
	Dim Usuario
	Dim Senha
	Dim IDCep
	Dim CEP
	Dim Logradouro
	Dim CompLog
	Dim Numero
	Dim Bairro
	Dim Municipio
	Dim Estado
	Dim bStatus
	Dim Telefone
	Dim DDD
	Dim NumeroMaxCartuchos
	Dim idtransp
	dim codbonus

	Sub GetPontosColta()
		Dim sSql, arrPonto, intPonto, i
		sSql = "SELECT [idPontos_coleta] " & _
				  ",[razao_social] " & _
				  ",[nome_fantasia] " & _
				  ",[cnpj] " & _
				  ",[bonus_type] " & _
				  ",[numero_endereco] " & _
				  ",[complemento_endereco] " & _
				  ",[usuario] " & _
				  ",[senha] " & _
				  ",[status_pontocoleta] " & _
			  	"FROM [marketingoki2].[dbo].[Pontos_coleta]"
		Call search(sSql, arrPonto, intPonto)
		If intPonto > -1 Then
			For i=0 to intPonto
				With Response
					If i Mod 2 = 0 Then
						.Write "<tr>"
						.Write "<td class='classColorRelPar'><img src=""img/buscar.gif"" alt=""Editar"" class=""imgexpandeinfo"" onClick=""window.location.href='frmPontoColetaAdm.asp?id="&arrPonto(0,i)&"&action=edit'""></td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(0,i)&"</td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(1,i)&"</td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(3,i)&"</td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(4,i)&"</td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(7,i)&"</td>"
						.Write "<td class='classColorRelPar'>"&arrPonto(8,i)&"</td>"
						If arrPonto(9,i) = 1 Then
							.Write "<td class='classColorRelPar'>Ativo</td>"
						Else
							.Write "<td class='classColorRelPar'>Inativo</td>"
						End If	
						.Write "</tr>"
					Else
						.Write "<tr>"
						.Write "<td class='classColorRelImpar'><img src=""img/buscar.gif"" alt=""Editar"" class=""imgexpandeinfo"" onClick=""window.location.href='frmPontoColetaAdm.asp?id="&arrPonto(0,i)&"&action=edit'""></td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(0,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(1,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(3,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(4,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(7,i)&"</td>"
						.Write "<td class='classColorRelImpar'>"&arrPonto(8,i)&"</td>"
						If arrPonto(9,i) = 1 Then
							.Write "<td class='classColorRelImpar'>Ativo</td>"
						Else
							.Write "<td class='classColorRelImpar'>Inativo</td>"
						End If	
						.Write "</tr>"
					End If
				End With
			Next
		Else
			Response.Write "<tr><td colspan=""8"" align=""center"" class=""classColorRelPar""><b>Nenhum Ponto de coleta encontrado!</b></td></tr>"	
		End If 			  
	End Sub
	
	sub getBonus()
		dim sql, arr, intarr, i
		dim sSelected
		sql = "SELECT [cod_bonus] FROM [marketingoki2].[dbo].[Cadastro_Bonus] where aplicacao = 'PONTO'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if codbonus = arr(0,i) then
					sSelected = "selected"
				else
					sSelected = ""
				end if	
				response.write "<option value="&arr(0,i)&" "&sSelected&">"&arr(0,i)&"</option>"
			next
		end if
	end sub
	
	Sub Requests()
		ID 				= Request.Form("id")
		RazaoSocial 	= Request.Form("txtRazaoSocial")
		NomeFantasia 	= Request.Form("txtNomeFantasia")
		CNPJ 			= Request.Form("txtCNPJ")
		TipoBonus 		= Request.Form("txtTipoBonus")
		Usuario 		= Request.Form("txtUsuario")
		Senha 			= Request.Form("txtSenha")
		CEP 			= Request.Form("txtCEP")
		Logradouro		= request.Form("txtLogradouro")
		CompLog 		= Request.Form("txtCompLog")
		Bairro			= request.Form("txtBairro")
		Municipio		= request.Form("txtMunicipio")
		Estado			= request.Form("txtEstado")
		Numero 			= Request.Form("txtNumero")
		bStatus 		= Request.Form("cbStatus")
		DDD 			= Request.Form("txtDDD")
		Telefone 		= Request.Form("txtTelefone")
		NumeroMaxCartuchos = Request.Form("txtQtdCartuchos")
		idtransp = Request.Form("cbTransp")
		codbonus = request.form("cbBonus")
	End Sub
	
	Sub Insert()
		Dim sSql
		if validaUsuario(Usuario, Senha) then
			if not CnpjExists(CNPJ) then
				sSql = "INSERT INTO [marketingoki2].[dbo].[Pontos_coleta]( " & _
						"[razao_social], " & _ 
						"[nome_fantasia], " & _ 
						"[cnpj], " & _ 
						"[bonus_type], " & _ 
						"[logradouro], " & _ 
						"[numero_endereco], " & _ 
						"[complemento_endereco], " & _ 
						"[bairro], " & _ 
						"[ddd], " & _ 
						"[telefone], " & _ 
						"[cep], " & _ 
						"[municipio], " & _ 
						"[estado], " & _ 
						"[usuario], " & _ 
						"[senha], " & _ 
						"[status_pontocoleta], " & _ 
						"[Qtd_Limite_Cartuchos], " & _ 
						"[idtransp]) " & _
						"VALUES( " & _
						"'"&RazaoSocial&"', " & _ 
						"'"&NomeFantasia&"', " & _ 
						"'"&CNPJ&"', " & _ 
						"'"&codbonus&"', " & _ 
						"'"&Logradouro&"', " & _ 
						""&Numero&", " & _ 
						"'"&CompLog&"', " & _ 
						"'"&Bairro&"', " & _ 
						""&DDD&", " & _ 
						""&Telefone&", " & _ 
						"'"&CEP&"', " & _ 
						"'"&Municipio&"', " & _ 
						"'"&Estado&"', " & _ 
						"'"&Usuario&"', " & _ 
						"'"&Senha&"', " & _ 
						""&bStatus&", " & _ 
						""&NumeroMaxCartuchos&", " & _ 
						""&idtransp&")" 
		
				Call exec(sSql)	   
			end if	
		else
			response.write "<script>alert('Usuário ou senha já cadastrados!')</script>"
		end if	
	End Sub
	
	Sub Submit()
		Call Requests()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			If ID <> "" Then
				Call Updates()
				Call SearchPonto()
			Else
				Call Insert()
				ID						= ""
				RazaoSocial				= ""
				NomeFantasia			= ""
				CNPJ					= ""
				TipoBonus				= ""
				Usuario					= ""
				Senha					= ""
				CEP						= ""
				CompLog					= ""
				Numero					= ""
				bStatus					= 0
				DDD						= ""
				Telefone				= ""
				NumeroMaxCartuchos		= ""
				idtransp				= 0
				Logradouro				= ""
				Bairro					= ""
				Municipio				= ""
				Estado					= ""
				codbonus				= ""
			End If	
		Else
			If Request.QueryString("action") = "edit" Then
				Call SearchPonto()
			End If	
		End If
	End Sub
	
	Sub SearchPonto()
		Dim sSql, arrPonto, intPonto, i
		Dim arrCep, intCep, j
		sSql = "SELECT [idPontos_coleta] " & _
				  ",[razao_social] " & _
				  ",[nome_fantasia] " & _
				  ",[cnpj] " & _
				  ",[bonus_type] " & _
				  ",[numero_endereco] " & _
				  ",[complemento_endereco] " & _
				  ",[usuario] " & _
				  ",[senha] " & _
				  ",[status_pontocoleta] " & _
				  ",[ddd] " & _
				  ",[telefone] " & _
				  ",[Qtd_Limite_cartuchos] " & _
				  ",[idtransp] " & _
				  ",[cep] " & _
				  ",[logradouro] " & _
				  ",[bairro] " & _
				  ",[municipio] " & _
				  ",[estado] " & _
			  "FROM [marketingoki2].[dbo].[Pontos_coleta] " & _
			  "WHERE [idPontos_coleta] = " & Request.QueryString("id")
		Call search(sSql, arrPonto, intPonto)			  
		If intPonto > -1 Then
			For i=0 To intPonto
				ID 				= arrPonto(0,i)
				RazaoSocial 	= arrPonto(1,i)
				NomeFantasia 	= arrPonto(2,i)
				CNPJ 			= arrPonto(3,i)
				codbonus 		= arrPonto(4,i)
				Numero 			= arrPonto(5,i)
				CompLog 		= arrPonto(6,i)
				Usuario 		= arrPonto(7,i)
				Senha 			= arrPonto(8,i)
				bStatus 		= arrPonto(9,i)
				DDD 			= arrPonto(10,i)
				Telefone 		= arrPonto(11,i)
				NumeroMaxCartuchos = arrPonto(12,i)
				idtransp = arrPonto(13,i)
				CEP				= arrPonto(14,i)
				Logradouro		= arrPonto(15,i)
				Bairro			= arrPonto(16,i)
				Municipio		= arrPonto(17,i)
				Estado			= arrPonto(18,i)
			Next
		End If
	End Sub
	
	Sub Updates()
		Dim sSql
		
		if validaUsuarioUpdate(ID, Usuario, Senha) then
			sSql = "UPDATE [marketingoki2].[dbo].[Pontos_coleta] " & _
					"SET " & _ 
					"[razao_social] = '"&RazaoSocial&"', " & _ 
					"[nome_fantasia] = '"&NomeFantasia&"', " & _ 
					"[cnpj] = '"&CNPJ&"', " & _ 
					"[bonus_type] = '"&codbonus&"', " & _ 
					"[logradouro]= '"&Logradouro&"', " & _ 
					"[numero_endereco] = "&Numero&", " & _ 
					"[complemento_endereco] = '"&CompLog&"', " & _ 
					"[bairro]='"&Bairro&"', " & _ 
					"[ddd]="&DDD&", " & _ 
					"[telefone]="&Telefone&", " & _ 
					"[cep]='"&CEP&"', " & _ 
					"[municipio]='"&Municipio&"', " & _ 
					"[estado]='"&Estado&"', " & _ 
					"[usuario]='"&Usuario&"', " & _ 
					"[senha]='"&Senha&"', " & _ 
					"[status_pontocoleta]="&bStatus&", " & _ 
					"[Qtd_Limite_Cartuchos]="&NumeroMaxCartuchos&", " & _ 
					"[idtransp]="&idtransp&" " & _
					"WHERE [idPontos_coleta]= " & ID
					
			'response.write sSql
			'response.End
							
			Call exec(sSql)				 
		else
			response.write "<script>alert('Usuário ou senha já cadastrados!')</script>"
		end if	
		
	End Sub
	
	function CnpjExists(cnpj)
		Dim sSql, arrCnpj, intCnpj
		sSql = "SELECT [idPontos_coleta] " & _
			   "FROM [marketingoki2].[dbo].[Pontos_coleta] WHERE [cnpj] = '" & cnpj & "'"
'		Response.Write sSql
		Call search(sSql, arrCnpj, intCnpj)	   
		If intCnpj > -1 Then
			response.write "<script>alert('CNPJ já cadastrado')</script>"
			CnpjExists = true
		Else
			CnpjExists = false
		End If		   
	End function
	
	Function GetTransportadoras()
		Dim sSql, arrTransp, intTransp, i
		dim ret
		dim selected
		ret = ""
		sSql = "select idtransportadoras, nome_fantasia from transportadoras where status = 1"
		
		Call search(sSql, arrTransp, intTransp)
		if intTransp > -1 then
			for i=0 to intTransp
				if idtransp = arrTransp(0,i) then
					selected = "selected" 
				else
					selected = ""
				end if	 	
				ret = ret & "<option value="""&arrTransp(0,i)&""" "&selected&">"&arrTransp(1,i)&"</option>"
			next
		end if
		GetTransportadoras = ret
	End Function
	
	function validaUsuario(usuario, senha)
		dim sql, arr, intarr, i
		sql = "select usuario, senha from pontos_coleta where usuario = '"&trim(usuario)&"' or senha = '"&trim(senha)&"'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			validaUsuario = false
		else
			validaUsuario = true
		end if	
	end function
	
	function validaUsuarioUpdate(id,usuario, senha)
		dim sql, arr, intarr, i
		dim usuario_update, senha_update
		usuario_update = request.form("usuario")
		senha_update = request.form("senha")
		if usuario <> usuario_update and senha <> senha_update then
			sql = "select usuario, senha from pontos_coleta where usuario = '"&usuario&"' or senha = '"&senha&"' where idpontos_coleta <> " & id
			call search(sql, arr, intarr)
			if intarr > -1 then
				validaUsuarioUpdate = false
			else
				validaUsuarioUpdate = true
			end if
		else	
			validaUsuarioUpdate = true
		end if	
	end function
	
	Call Submit()
	
%>
<html>
<head>
<script src="js/frmPontoColetaAdm.js"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="" name="frmPontoColetaAdm" method="POST">
		<input type="hidden" name="id" value="<%= Request.QueryString("id") %>" />
		<input type="hidden" name="usuario" value="<%= Usuario %>" />
		<input type="hidden" name="senha" value="<%= Senha %>" />
		<input type="hidden" name="verifycnpj" value="<%= CNPJ %>" />
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellspacing="3" cellpadding="2" width="100%" id="tablecadtransportadoras">
						<tr>
							<td id="explaintitle" colspan="2" align="center">Pontos de Coleta</td>
						</tr>
						<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td align="right" width="20%">Razão Social: </td>
							<td align="left"><input type="text" class="text" name="txtRazaoSocial" value="<%= RazaoSocial %>" size="40" maxlength="255" /> * </td>
						</tr>
						<tr>
							<td align="right">Nome Fantasia: </td>
							<td align="left"><input type="text" class="text" name="txtNomeFantasia" value="<%= NomeFantasia %>" size="40" maxlength="255" /> * </td>
						</tr>
						<tr>
							<td align="right">CNPJ: </td>
							<td align="left"><input type="text" class="text" name="txtCNPJ" value="<%= CNPJ %>" size="22" maxlength="18" onKeyPress="cnpj_format(this)" onBlur="cnpjExists()"  /> * Ex: 88.888.888/0001-92</td>
						</tr>
						<tr>
							<td align="right">Usuário: </td>
							<td align="left"><input type="text" class="text" name="txtUsuario" value="<%= Usuario %>" size="25" maxlength="20" /> * </td>
						</tr>
						<tr>
							<td align="right">Senha: </td>
							<td align="left"><input type="password" class="text" name="txtSenha" value="<%= Senha %>" size="25" maxlength="20" /> * </td>
						</tr>
						<tr>
							<td align="right">CEP: </td>
							<td align="left"><input type="text" class="text" name="txtCEP" value="<%= CEP %>" size="10"  maxlength="8"/> <img src="img/buscar.gif" class="imgexpandeinfo" align="absmiddle" alt="Buscar" onClick="endereco()" /></td>
						</tr>
						<tr>
							<td align="right">Logradouro: </td>
							<td align="left"><input name="txtLogradouro" type="text" class="textreadonly" value="<%= Logradouro %>" size="50" /> * </td>
						</tr>
						<tr>
							<td align="right">Comp. Logradouro: </td>
							<td align="left"><input name="txtCompLog" type="text" class="text" value="<%= CompLog %>" /></td>
						</tr>
						<tr>
							<td align="right">Número: </td>
							<td align="left"><input name="txtNumero" type="text" class="text" value="<%= Numero %>" size="5" /> * </td>
						</tr>
						<tr>
							<td align="right">Bairro: </td>
							<td align="left"><input name="txtBairro" type="text" class="textreadonly" value="<%= Bairro %>" /> * </td>
						</tr>
						<tr>
							<td align="right">Município: </td>
							<td align="left"><input name="txtMunicipio" type="text" class="textreadonly" value="<%= Municipio %>" /> * </td>
						</tr>
						<tr>
							<td align="right">Estado: </td>
							<td align="left"><input name="txtEstado" type="text" class="textreadonly" size="2" value="<%= Estado %>" /> * </td>
						</tr>
						<tr>
							<td align="right">DDD: </td>
							<td align="left"><input name="txtDDD" type="text" class="textreadonly" value="<%= DDD %>" size="4" maxlength="2" /> 
							* </td>
						</tr>
						<tr>
							<td align="right">Telefone: </td>
							<td align="left"><input name="txtTelefone" type="text" class="textreadonly" value="<%= Telefone %>" size="10" maxlength="8" /> 
							* </td>
						</tr>
						<tr>
							<td align="right">Transportadora: </td>
							<td align="left">
								<select name="cbTransp" class="select">
									<option value="0">Selecione uma Transportadora</option>
									<%=GetTransportadoras()%>
								</select>
								<img src="img/transportadoras.gif" align="absmiddle" alt="Escolher Transportadora" class="imgexpandeinfo" width="25" height="25" onClick="window.open('frmSearchTranspPontoColeta.asp','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
							</td>
						</tr>
						<tr>
							<td align="right" width="33%">Cód. Bônus:</td>
							<td align="left">
								<select name="cbBonus" class="select" style="width:200px;">
									<option value="">[Selecione]</option>
									<%Call getBonus()%>
								</select>
								<img src="img/bonus.gif" class="imgexpandeinfo" width="23" height="23" align="absmiddle" alt="Buscar Bônus" onClick="window.open('frmsearchbonuspontocoleta.asp','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
							</td>
						</tr>
						<tr>
							<td align="right">Número Máximo de Cartuchos</td>
							<td align="left"><input name="txtQtdCartuchos" type="text" class="text" value="<%=NumeroMaxCartuchos%>" size="10" /> * </td>
						</tr>
						<tr>
							<td align="right">Status: </td>
							<td align="left">
								<select name="cbStatus" class="select">
									<option value="1" <% If bStatus = 1 Then %>selected<% End If %>>Ativo</option>
									<option value="0" <% If bStatus = 0 Then %>selected<% End If %>>Inativo</option>
								</select> * 
							</td>
						</tr>
						<tr>
							<td colspan="2" align="center"><input type="button" class="btnform" name="btnSalvar" <% If Request.QueryString("action") = "edit" Then %>value="Editar"<% Else %>value="Salvar"<% End If %> onClick="validate()" /></td>
						</tr>
						<tr>
							<td colspan="2">
								<table width="100%" cellpadding="1" cellspacing="1" id="tablelisttransportadoras">
									<th><img src="img/check.gif"></th>
									<th>ID</th>
									<th>Razão Social</th>
									<th>CNPJ</th>
									<th>Tipo Bônus</th>
									<th>Usuário</th>
									<th>Senha</th>
									<th>Status</th>
									<%Call GetPontosColta()%>
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
