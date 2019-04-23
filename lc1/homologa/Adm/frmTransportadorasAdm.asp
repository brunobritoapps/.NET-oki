<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>
<%
	Dim ID
	Dim RazaoSocial
	Dim NomeFantasia
	Dim CNPJ
	Dim Contato
	Dim DDD
	Dim Telefone
	Dim Fax
	Dim ColetaPorEmail
	Dim Ativo
	Dim Email

	Sub Submit()
		Call Requests()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			If ID <> "" Then
				Call Updates()
			Else	
				Call Insert()
			End If
			RazaoSocial 		 = ""
			NomeFantasia 		 = ""
			CNPJ				 = ""
			Contato 		     = ""
			DDD 				 = ""
			Telefone	 		 = ""
			Fax 				 = ""
			ColetaPorEmail 		 = 0
			Email = ""
		Else
			If Request.QueryString("action") = "edit" Then
				Call SearchTransp()
			End If
		End If							
	End Sub
	
	Sub Requests()
		ID = Request.Form("id")
		RazaoSocial 				 = Request.Form("txtRazaoSocial")
		NomeFantasia 				 = Request.Form("txtNomeFantasia")
		CNPJ						 = Request.Form("txtCNPJ")
		Contato 				     = Request.Form("txtContato")
		DDD 						 = Request.Form("txtDDD")
		Telefone	 				 = Request.Form("txtTelefone")
		Fax 						 = Request.Form("txtFax")
		ColetaPorEmail 			 	 = Request.Form("radioColetaEmail")
		Ativo 						 = Request.Form("radioAtivo")
		Email							 = Request.Form("txtEmail")	
	End Sub
	
	Sub ListTransp()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT " & _
						"[idTransportadoras], " & _ 
						"[razao_social], " & _ 
						"[nome_fantasia], " & _ 
						"[cnpj], " & _ 
						"[contato], " & _ 
						"[ddd], " & _ 
						"[telefone], " & _ 
						"[fax], " & _ 
						"[isColetaEmail], " & _ 
						"[status] " & _
						"FROM [marketingoki2].[dbo].[Transportadoras]"
						
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If i Mod 2 = 0 Then
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelPar'><img class=""imgexpandeinfo"" src=""img/buscar.gif"" alt=""Alterar"" onClick=""window.location.href='frmTransportadorasAdm.asp?id="&arrTransp(0,i)&"&action=edit'""/></td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(0,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(1,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(2,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(3,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(4,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(5,i)&"</td>"
					Response.Write "<td class='classColorRelPar'>"&arrTransp(6,i)&"</td>"
					If arrTransp(8,i) = 1 Then 
						Response.Write "<td class='classColorRelPar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelPar'>Não</td>"
					End If
					If arrTransp(9,i) = 1 Then
						Response.Write "<td class='classColorRelPar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelPar'>Não</td>"
					End If
					Response.Write "</tr>"
				Else
					Response.Write "<tr>"
					Response.Write "<td class='classColorRelImpar'><img class=""imgexpandeinfo"" src=""img/buscar.gif"" alt=""Alterar"" onClick=""window.location.href='frmTransportadorasAdm.asp?id="&arrTransp(0,i)&"&action=edit'""/></td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(0,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(1,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(2,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(3,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(4,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(5,i)&"</td>"
					Response.Write "<td class='classColorRelImpar'>"&arrTransp(6,i)&"</td>"
					If arrTransp(8,i) = 1 Then 
						Response.Write "<td class='classColorRelImpar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelImpar'>Não</td>"
					End If
					If arrTransp(9,i) = 1 Then
						Response.Write "<td class='classColorRelImpar'>Sim</td>"
					Else
						Response.Write "<td class='classColorRelImpar'>Não</td>"
					End If
					Response.Write "</tr>"
				End If
			Next
		Else
			Response.Write "<tr><td colspan=""10"" align=""center"" class=""classColorRelPar""><b>Nenhuma transportadora encontrada!</b></td></tr>"	
		End If 				
	End Sub
	
	Sub SearchTransp()
		Dim sSql, arrTransp, intTransp, i
		sSql = "SELECT " & _
						"[idTransportadoras], " & _ 
						"[razao_social], " & _ 
						"[nome_fantasia], " & _ 
						"[cnpj], " & _ 
						"[contato], " & _ 
						"[ddd], " & _ 
						"[telefone], " & _ 
						"[fax], " & _ 
						"[isColetaEmail], " & _ 
						"[status], " & _
						"[email] " & _
						"FROM [marketingoki2].[dbo].[Transportadoras] " & _
						"WHERE [idTransportadoras] = " & Request.QueryString("id")
		Call search(sSql, arrTransp, intTransp)						
		If intTransp > -1 Then
			For i=0 To intTransp
				RazaoSocial 		 = arrTransp(1,i)
				NomeFantasia 		 = arrTransp(2,i)
				CNPJ				 = arrTransp(3,i)
				Contato 		     = arrTransp(4,i)
				DDD 				 = arrTransp(5,i)
				Telefone	 		 = arrTransp(6,i)
				Fax 				 = arrTransp(7,i)
				ColetaPorEmail 		 = arrTransp(8,i)
				Ativo 				 = arrTransp(9,i)
				Email          = arrTransp(10,i)
			Next
		End If
	End Sub
	
	Sub Updates()
		Dim sSql
		sSql = "UPDATE [marketingoki2].[dbo].[Transportadoras] " & _
				   "SET [razao_social] = '"&RazaoSocial&"' " & _
					  ",[nome_fantasia] = '"&NomeFantasia&"' " & _
					  ",[cnpj] = '"&CNPJ&"' " & _
					  ",[contato] = '"&Contato&"' " & _
					  ",[ddd] = "&DDD&" " & _
					  ",[telefone] = "&Telefone&" " & _
					  ",[fax] = "&Fax&" " & _
					  ",[isColetaEmail] = "&ColetaPorEmail&" " & _
					  ",[status] = "&Ativo&" " & _
						",[email] = '"&Email&"' " & _
				 "WHERE [idTransportadoras] = " & ID
		Call exec(sSql)				 		
	End Sub
	
	Sub Deletes()
	End Sub
	
	Sub Insert()
		Dim sSql
		
		sSql = "INSERT INTO " & _
						"[marketingoki2].[dbo].[Transportadoras]( " & _
						"[razao_social], " & _ 
						"[nome_fantasia], " & _ 
						"[cnpj], " & _ 
						"[contato], " & _ 
						"[ddd], " & _ 
						"[telefone], " & _ 
						"[fax], " & _ 
						"[isColetaEmail], " & _ 
						"[status], " & _
						"[email]) " & _
						"VALUES( " & _
						"'"&RazaoSocial&"', " & _ 
						"'"&NomeFantasia&"', " & _ 
						"'"&CNPJ&"', " & _ 
						"'"&Contato&"', " & _ 
						""&DDD&", " & _ 
						""&Telefone&", " & _ 
						""&Fax&", " & _ 
						""&ColetaPorEmail&", " & _ 
						""&Ativo&", " & _
						"'"&Email&"')"
		Call exec(sSql)
	End Sub

	Call Submit()	
%>
<html>
<head>
<script src="js/frmTransportadorasAdm.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">
		<form action="frmTransportadorasAdm.asp" name="frmTransportadorasAdm" method="POST">
		<input type="hidden" name="id" value="<%= Request.QueryString("id") %>">
		<input type="hidden" name="verifycnpj" value="<%= CNPJ %>" />
			<tr> 
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<div id="painelcontrole">
						<table cellspacing="3" cellpadding="2" width="100%" id="tablecadtransportadoras">
							<tr>
								<td colspan="3" id="explaintitle" align="center">Transportadoras</td>
							</tr>
							<tr>
								<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalAdm.asp';">&laquo Voltar</a></td>
							</tr>
							<tr>
								<td width="25%" align="right">Razão Social: </td>
								<td><input type="text" name="txtRazaoSocial" class="text" value="<%= RazaoSocial %>" size="40" maxlength="255"/> * </td>
							</tr>
							<tr>
								<td width="25%" align="right">Nome Fantasia: </td>
								<td><input type="text" name="txtNomeFantasia" class="text" value="<%= NomeFantasia %>" size="40" maxlength="255" /> * </td>
							</tr>
							<tr>
								<td width="25%" align="right">CNPJ: </td>
								<td><input type="text" name="txtCNPJ" class="text" value="<%= CNPJ %>" size="22" maxlength="20" onKeyPress="cnpj_format(this)" onBlur="cnpjExists()" /> * Ex: 88.888.888/0001-92</td>
							</tr>
							<tr>
								<td width="25%" align="right">Contato: </td>
								<td><input type="text" name="txtContato" class="text" value="<%= Contato %>" size="30" maxlength="45" /> * </td>
							</tr>
							<tr>
								<td width="25%" align="right">DDD: </td>
								<td><input type="text" name="txtDDD" class="text" value="<%= DDD %>" size="3" maxlength="2" /> 
								* </td>
							</tr>
							<tr>
								<td width="25%" align="right">Telefone: </td>
								<td><input type="text" name="txtTelefone" class="text" value="<%= Telefone %>" size="10" maxlength="8" /> * </td>
							</tr>
							<tr>
								<td width="25%" align="right">Fax: </td>
								<td><input type="text" name="txtFax" class="text" value="<%= Fax %>" size="10" maxlength="8" /> * </td>
							</tr>
							<tr>
								<td width="25%" align="right">E-mail: </td>
								<td><input type="text" class="text" name="txtEmail" value="<%=Email%>" size="30" /><label id="obrig_email"></label></td>
							</tr>
							<tr>
								<td width="25%" align="right">Coleta por Email: </td>
								<td>Não <input type="radio" name="radioColetaEmail" value="0" <% If ColetaPorEmail = 0 Then %>checked<% End If %> onClick="checkObrigatorioEmail()" />Sim <input type="radio" name="radioColetaEmail" value="1" <% If ColetaPorEmail = 1 Then %>checked<% End If %> onClick="checkObrigatorioEmail()" /></td>
							</tr>
							<tr>
								<td width="25%" align="right">Ativo</td>
								<td>Não <input type="radio" name="radioAtivo" value="0" checked="checked" <% If Ativo = 0 Then %>checked<% End If %> />Sim <input type="radio" name="radioAtivo" value="1" <% If Ativo = 1 Then %>checked<% End If %> /></td>
							</tr>
							<tr>
								<td align="right"><input type="button" class="btnform" name="btnSubmit" <% If Request.QueryString("action") = "edit" Then %>value="Editar"<% Else %>value="Salvar"<% End If %> onClick="validate()" /></td>
								<td align="left"><input type="reset" class="btnform" name="btnLimpar" value="Limpar" /></td>
							</tr>
							<tr>
								<td colspan="2">
									<table cellpadding="1" cellspacing="1" width="100%" id="tablelisttransportadoras">
										<tr>
											<th><img src="img/check.gif"></th>
											<th>ID</th>
											<th>Razão Social</th>
											<th>Nome Fantasia</th>
											<th>CNPJ</th>
											<th>Contato</th>
											<th>DDD</th>
											<th>Telefone</th>
											<th>Coleta Email</th>
											<th>Ativo</th>
										</tr>
										<%Call ListTransp()%>
									</table>
								</td>
							</tr>
						</table>
					</div>
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
