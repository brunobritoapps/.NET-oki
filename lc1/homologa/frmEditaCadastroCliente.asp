
<!--#include file="_config/_config.asp" -->
<%
'|--------------------------------------------------------------------
'| Arquivo: frmEditaCadastroCliente.asp																									 
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)						 
'| Data Criação: 13/04/2007																					 
'| Data Modificação: 15/04/2007																		 
'| Descrição: Arquivo de Formulário para UPDATE de Cliente
'|--------------------------------------------------------------------
%>
<%Call open()%>
<%Call getSessionUser()%>

<%
	Dim RazaoSocial
	Dim NomeFantasia
	Dim CNPJ
	Dim InscricaoEstadual
	Dim DDD
	Dim Telefone
	Dim CompEndereco
	Dim NumeroEndereco
	Dim CEP
	Dim Logradouro
	Dim Bairro
	Dim Municipio
	Dim Estado
	Dim TipoLogradouro
	Dim tipopessoa
	
	'CEP Coleta Peterson Aquino - 04/5/2014
	Dim CompEndColeta
	Dim NumeroEndColeta
	Dim CEPColeta
	Dim LogradouroColeta
	Dim BairroColeta
	Dim MunicipioColeta
	Dim EstadoColeta
		
'	response.write Session("IDCliente")
'	response.End()
	
	Sub getInfoList()
	
		Dim sSql, arrInfoList, intInfoList, i

		sSql = "SELECT " & _ 
						"[A].[idClientes], " & _ 
						"[A].[razao_social], " & _ 
						"[A].[nome_fantasia], " & _ 
						"[A].[cnpj], " & _ 
						"[A].[inscricao_estadual], " & _ 
						"[A].[ddd], " & _ 
						"[A].[telefone], " & _ 
						"[A].[compl_endereco], " & _ 
						"[A].[compl_endereco_coleta], " & _ 
						"[A].[numero_endereco], " & _ 
						"[A].[numero_endereco_coleta], " & _
						"[B].[cep], " & _ 
						"[B].[logradouro], " & _ 
						"[B].[bairro], " & _ 
						"[B].[municipio], " & _ 
						"[B].[estado], " & _ 
						"[A].[tipopessoa] " & _
						"FROM [marketingoki2].[dbo].[Clientes] AS [A] " & _
						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS [B] " & _
						"ON [A].[idClientes] = [B].[Clientes_idClientes] " & _
						"WHERE [B].[isEnderecoComum] = 1 AND " & _
						"[A].[idClientes] = " & Session("IDCliente")
		
'		Response.Write sSql
'		Response.End()				
		
		Call search(sSql, arrInfoList, intInfoList)

		If intInfoList > -1 Then
			For i=0 To intInfoList
				RazaoSocial				= arrInfoList(1,i)
				NomeFantasia			= arrInfoList(2,i)
				CNPJ					= arrInfoList(3,i)
				InscricaoEstadual		= arrInfoList(4,i)
				DDD						= arrInfoList(5,i)
				Telefone				= arrInfoList(6,i)
				CompEndereco			= arrInfoList(7,i)
				NumeroEndereco			= arrInfoList(9,i)
				CEP						= arrInfoList(11,i)
				Logradouro				= arrInfoList(12,i)
				Bairro					= arrInfoList(13,i)
				Municipio				= arrInfoList(14,i)
				Estado					= arrInfoList(15,i)
				tipopessoa				= arrInfoList(16,i)
			Next
		End If
		
		'
		'***************************************************************************************************
		'Pega o endereço de coleta
		'Peterson 4/5/2014
		'
		Dim sSqlColeta, arrInfoListColeta, intInfoListColeta, iColeta

		sSqlColeta = "SELECT " & _ 
						"[A].[compl_endereco_coleta], " & _ 
						"[A].[numero_endereco_coleta], " & _
						"[B].[cep], " & _ 
						"[B].[logradouro], " & _ 
						"[B].[bairro], " & _ 
						"[B].[municipio], " & _ 
						"[B].[estado] " & _ 
						"FROM [marketingoki2].[dbo].[Clientes] AS [A] " & _
						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS [B] " & _
						"ON [A].[idClientes] = [B].[Clientes_idClientes] " & _
						"WHERE [B].[isEnderecoComum] = 0 AND [B].[isEnderecoColeta] = 1 AND " & _
						"[A].[idClientes] = " & Session("IDCliente")
		
		Call search(sSqlColeta, arrInfoListColeta, intInfoListColeta)

		If intInfoListColeta > -1 Then
			'Dim CompEndColeta
			'Dim NumeroEndColeta
			'Dim CEPColeta
			'Dim LogradouroColeta
			'Dim BairroColeta
			'Dim MunicipioColeta
			'Dim EstadoColeta
			
			For iColeta=0 To intInfoListColeta
				CompEndColeta			= arrInfoListColeta(0,iColeta)
				NumeroEndColeta			= arrInfoListColeta(1,iColeta)
				CEPColeta				= arrInfoListColeta(2,iColeta)
				LogradouroColeta		= arrInfoListColeta(3,iColeta)
				BairroColeta			= arrInfoListColeta(4,iColeta)
				MunicipioColeta			= arrInfoListColeta(5,iColeta)
				EstadoColeta			= arrInfoListColeta(6,iColeta)
			Next
		End If		
		
	End Sub

	Sub RequestForm()
		RazaoSocial										= Request.Form("txtRazaoSocial")
		NomeFantasia									= Request.Form("txtNomeFantasia")
		if request.form("tipopessoa") = 1 then
			CNPJ										= Request.Form("txtCNPJ")
		else
			CNPJ										= request.form("txtCPF")
		end if	
		InscricaoEstadual								= Request.Form("txtInscricaoEstadual")
		DDD												= Request.Form("txtDDD")
		Telefone										= Request.Form("txtTelefone")
		CompEndereco									= Request.Form("txtCompLogradouro")
		NumeroEndereco									= Request.Form("txtNumero")
		CEP												= Request.Form("txtCep")
		Logradouro										= request.Form("txtLogradouro")
		Bairro											= request.Form("txtBairro")
		Municipio										= request.Form("txtMunicipio")
		Estado											= request.Form("txtEstado")
		
		'
		'salva os dados do endereço para coleta
		'peterson 4-5-2014
		'
		CompEndColeta			= Request.Form("txtCompLogradouroColeta")
		NumeroEndColeta			= Request.Form("txtNumeroColeta")
		CEPColeta				= Request.Form("txtColetaCep")
		LogradouroColeta		= Request.Form("txtLogradouroColeta")
		BairroColeta			= Request.Form("txtBairroColeta")
		MunicipioColeta			= Request.Form("txtMunicipioColeta")
		EstadoColeta			= Request.Form("txtEstadoColeta")
		
'		With Response
'			.Write "RazaoSocial: " & RazaoSocial & "<br />"
'			.Write "NomeFantasia: " & NomeFantasia & "<br />"
'			.Write "CNPJ: " & CNPJ & "<br />"
'			.Write "InscricaoEstadual: " & InscricaoEstadual & "<br />"
'			.Write "DDD: " & DDD & "<br />"
'			.Write "Telefone: " & Telefone & "<br />"
'			.Write "CompEndereco: " & CompEndereco & "<br />"
'			.Write "NumeroEndereco: " & NumeroEndereco & "<br />"
'			.Write "CEP: " & CEP & "<br />"
'		End With
	End Sub

	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
			Call RequestForm()
			Call Update()
		End If
	End Sub

	Sub Update()
	
		Dim oCommand, oComCol

		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandType = 4
		oCommand.CommandText = "sp_UpdateClienteLc" 

		oCommand.Parameters("@IDCliente")				= CLng(Session("IDCliente"))
		oCommand.Parameters("@RazaoSocial")				= RazaoSocial
		oCommand.Parameters("@NomeFantasia")			= NomeFantasia
		oCommand.Parameters("@CNPJ")					= CNPJ
		oCommand.Parameters("@InscricaoEstadual")		= InscricaoEstadual
		oCommand.Parameters("@DDD")						= CInt(DDD)
		oCommand.Parameters("@Telefone")				= CLng(Telefone)
		oCommand.Parameters("@CompEndereco")			= CompEndereco
		oCommand.Parameters("@NumeroEndereco")			= CLng(NumeroEndereco)
		oCommand.Parameters("@CEP")						= CEP
		oCommand.Parameters("@logradouro")				= Logradouro
		oCommand.Parameters("@bairro")					= Bairro
		oCommand.Parameters("@municipio")				= Municipio
		oCommand.Parameters("@estado")					= Estado
		oCommand.Parameters("@isColetaDomiciliar")		= Session("isColetaDomiciliar")
		
		oCommand.Execute()
		
		'
		' peterson Atualiza Endereço Coleta
		'
		Set oComCol = Server.CreateObject("ADODB.Command")
		oComCol.CommandTimeout = 200
		oComCol.ActiveConnection = oConn
		oComCol.CommandType = 4
		oComCol.CommandText = "sp_UpdateClienteEndCol"

		oComCol.Parameters("@IDCliente")				= CLng(Session("IDCliente"))
		oComCol.Parameters("@CompEndColeta")			= CompEndColeta
		oComCol.Parameters("@NumeroEndColeta")			= CLng(NumeroEndColeta)
		oComCol.Parameters("@CEPColeta")				= CEPColeta
		oComCol.Parameters("@LogradouroColeta")			= LogradouroColeta
		oComCol.Parameters("@BairroColeta")				= BairroColeta
		oComCol.Parameters("@MunicipioColeta")			= MunicipioColeta
		oComCol.Parameters("@EstadoColeta")				= EstadoColeta
		oComCol.Parameters("@isColetaDomiciliar")		= Session("isColetaDomiciliar")
		
		oComCol.Execute()
		
		Call UpdateSessions()
		
		Response.Redirect "frmOperacionalCliente.asp"

		Set oCommand = Nothing
		Set oComCol = Nothing
		
	End Sub

	Call getInfoList()
	Call SubmitForm()
%>
<html>
<head>

<script src="http://www.sustentabilidadeoki.com.br/lc/homologa/js/frmEditaCadastroCliente.js"></script>
<script src="http://www.sustentabilidadeoki.com.br/lc/homologa/js/frmFindCep.js"></script>

<link rel="stylesheet" type="text/css" href="css/geral.css">
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
					<form action="frmEditaCadastroCliente.asp" name="frmEditaCadastroCliente" method="POST">
					<input type="hidden" name="tipopessoa" value="<%=tipopessoa%>" />
					<table cellpadding="3" cellspacing="4" width="100%" id="tableEditClienteCadastro" border="0">
						<tr>
							<td colspan="3" id="explaintitle" align="center">Cadastro da Empresa</td>
						</tr>	

						<tr>
							<!--<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>-->
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>
						</tr>
						<!--<tr>
							<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>
						</tr>-->
						<tr>
							<td colspan="2" align="left"><b id="fontred">Atenção :</b>
										<b style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-size: 13px; vertical-align: baseline; background-color: transparent; color: rgb(55, 61, 69); font-family: Arial, sans-serif; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: 14px; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Os campos com (asterisco)* são de preenchimento obrigatório.</b><td>
						</tr>
						<tr>
							<td colspan="2" align="center">&nbsp;<td>
						    &nbsp;</tr>
						<tr <%if tipopessoa = 0 then%>style="display:none;"<%end if%>>
							<td align="right" width="25%">Razão Social:</td>
							<td align="left"><input type="text" class="textreadonly" name="txtRazaoSocial" value="<%=RazaoSocial%>" size="40" /> *</td>
						</tr>
						<tr>
							<td align="right" width="25%"><%if tipopessoa = 1 then%>Nome Fantasia:<%else%>Nome:<%end if%></td>
							<td align="left"><input type="text" class="oki-input" name="txtNomeFantasia" value="<%=NomeFantasia%>" size="40" /> *</td>
						</tr>
						<tr>
							<td align="right" width="25%"><%if tipopessoa = 0 then%>CPF:<%else%>CNPJ:<%end if%></td>
							<td align="left">
								<%if tipopessoa = 1 then%>
									<input type="text" class="oki-input" id="txtCPF" name="txtCPF" value="<%=CNPJ%>" size="22" maxlength="14" style="display:none;" /> 
									<input type="text" class="oki-input" id="txtCNPJ" name="txtCNPJ" value="<%=CNPJ%>" size="22" maxlength="18" onkeypress="cnpj_format(this)"  /> 
								<%else%>	
									<input type="text" class="oki-input" id="txtCPF" name="txtCPF" value="<%=CNPJ%>" size="22" maxlength="14" onkeypress="cpf_format(this)" /> 
									<input type="text" class="oki-input" id="txtCNPJ" name="txtCNPJ" value="<%=CNPJ%>" size="22" maxlength="18" style="display:none;" /> 
								<%end if%>
									*
								<%if tipopessoa = 1 then%>Ex: 88.888.888/0001-91<%else%>Ex: 888.888.888-88<%end if%>
							</td>
						</tr>
						<tr <%if tipopessoa = 0 then%>style="display:none;"<%end if%>>
							<td align="right" width="25%">Inscrição Estadual:</td>
							<td align="left">
								<input type="text" style="text-transform: uppercase;" class="oki-input" name="txtInscricaoEstadual" value="<%=InscricaoEstadual%>" size="18" maxlength="15" />
								Digite somente números ou a palavra: ISENTO
							</td>
						</tr>

						<tr>
							<td align="right" width="25%">DDD:</td>
							<td align="left"><input type="text" class="oki-input" name="txtDDD" value="<%=DDD%>" size="3" maxlength="2" /> *</td>
						</tr>
						<tr>
							<td align="right" width="25%">Telefone:</td>
							<td align="left">
								<input type="text" class="oki-input" name="txtTelefone" value="<%=Telefone%>" size="10" maxlength="8" /> *
								Digite somente números
							</td>
						</tr>
						<tr>
							<td align="right" width="25%">CEP:</td>
							<td align="left"><input type="text" class="oki-input" name="txtCep" value="<%=CEP%>" size="10" maxlength="8" /> * Digite somente números&nbsp;
							<input type="button" class="btnform" name="btnNexBuscaEndereco" value="Buscar Endereço" onClick="showClienteEndereco()" />
							</td>
						</tr>
						<!--<tr>
							<td align="right" width="25%">Clicar para buscar endereço:</td>
							<td align="left"><input type="button" class="btnform" name="btnNexBuscaEndereco" value="Buscar Endereço" onClick="showClienteEndereco()" />
						</tr>-->
						<tr>
							<td id="loadingdisplay" style="display:none;position:fixed;left:50%;top:50%;background:#FFFFFF" align="left"><b>Loading</b> <img align="absmiddle" src="img/ajax-loader.gif" name="loading"/></td>							
						</tr>
						<!--<tr>
							<td colspan="3" id="explaintitle" align="center">Endereço da Empresa</td>
						</tr>-->
						<tr>
							<td align="right" width="25%">Logradouro:</td>
							<td align="left"><input type="text" class="oki-input" name="txtLogradouro" value="<%=Logradouro%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Complemento Logradouro:</td>
							<td align="left"><input type="text" class="oki-input" name="txtCompLogradouro" value="<%=CompEndereco%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Número:</td>
							<td align="left"><input type="text" class="oki-input" name="txtNumero" value="<%=NumeroEndereco%>" size="10" maxlength="8" /> *</td>
						</tr>
						<tr>
							<td align="right" width="25%">Bairro:</td>
							<td align="left"><input type="text" class="oki-input" name="txtBairro" value="<%=Bairro%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Município:</td>
							<td align="left"><input type="text" class="oki-input" name="txtMunicipio" value="<%=Municipio%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Estado:</td>
							<td align="left"><input type="text" class="oki-input" name="txtEstado" value="<%=Estado%>" size="2" /></td>
						</tr>
						<tr>
							<td colspan="3" id="explaintitle" align="center">Endereço Padrão para Coleta</td>
						</tr>
						
						<tr>
							<td align="right" width="25%">&nbsp;</td>
							<td><input type="checkbox" class="checkbox" id="chkMesmoEndereco" name="checkboxgroup" value="true" onClick="usaMesmoEndereco()" />Usar Endereço da Empresa&nbsp;						
							</td>
						</tr>
						
						<tr>
							<td align="right" width="25%">CEP:</td>
							<td align="left"><input type="text" class="oki-input" name="txtColetaCep" value="<%=CEPColeta%>" size="10" maxlength="8" /> * Digite somente números&nbsp;
							<input type="button" class="btnform" name="btnBuscaCepColeta" value="Buscar Endereço" onClick="showClienteEndColeta()" />
							</td>
						</tr>
						
						<tr>
							<td align="right" width="25%">Logradouro:</td>
							<td align="left"><input type="text" class="oki-input" id="txtLogradouroColeta" name="txtLogradouroColeta" value="<%=LogradouroColeta%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Complemento Logradouro:</td>
							<td align="left"><input type="text" class="oki-input" name="txtCompLogradouroColeta" value="<%=CompEndColeta%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Número:</td>
							<td align="left"><input type="text" class="oki-input" name="txtNumeroColeta" value="<%=NumeroEndColeta%>" size="10" maxlength="8" /> *</td>
						</tr>
						<tr>
							<td align="right" width="25%">Bairro:</td>
							<td align="left"><input type="text" class="oki-input" name="txtBairroColeta" value="<%=BairroColeta%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Município:</td>
							<td align="left"><input type="text" class="oki-input" name="txtMunicipioColeta" value="<%=MunicipioColeta%>" size="40" /></td>
						</tr>
						<tr>
							<td align="right" width="25%">Estado:</td>
							<td align="left"><input type="text" class="oki-input" name="txtEstadoColeta" value="<%=EstadoColeta%>" size="2" /></td>
						</tr>
						<!--<tr>
							<td align="right"><input type="reset" class="btnform" name="btnLimpar" value="Limpar" /></td>
							<td align="left" colspan="2"><input type="button" class="btnform" name="btnSalvar" value="Salvar" onClick="validate()" /></td>							
						</tr>-->
						</form>
						<tr>
							<td align="right" width="25%">&nbsp;</td>
							<td align="right"><input type="button" class="btnform" name="btnSubmitSolicitacao1" value="Salvar" onClick="validate()" /></td>
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
