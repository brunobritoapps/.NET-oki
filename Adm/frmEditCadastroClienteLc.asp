<!--#include file="../_config/_config.asp" -->
<%Call open()%>
<%Call GetSessionAdm()%>

<%
	Dim RazaoSocial
	Dim NomeFantasia
	Dim CNPJ
	Dim InscEstadual
	Dim DDD
	Dim Telefone
	Dim Categoria
	Dim Grupo
	Dim IDCep
	Dim CEP
	Dim LogCliente
	Dim CompLog
	Dim Numero
	Dim Bairro
	Dim Municipio
	Dim Estado
	Dim IDCepColeta
	Dim CEPColeta
	Dim LogColeta
	Dim CompLogColeta
	Dim NumeroColeta
	Dim BairroColeta
	Dim MunicipioColeta
	Dim EstadoColeta
	Dim isColetaDomiciliar
	Dim StatusCliente
	Dim IDPontoColeta
	Dim NomePontoColeta
	Dim CNPJPontoColeta
	Dim MotivoStatus
	Dim ContatoRespColeta
	Dim DDDContatoRespColeta
	Dim	TelefoneContatoRespColeta
	Dim TipoColeta
	Dim TipoPessoa
	Dim MinCartuchos
	Dim CodCliConsolidador
	' Endereço do Cliente
	dim cep_cliente
	dim logradouro_cliente
	dim bairro_cliente
	dim municipio_cliente
	dim estado_cliente
	' Endereço de Coleta
	dim cep_coleta
	dim logradouro_coleta
	dim bairro_coleta
	dim municipio_coleta
	dim estado_coleta

    ' Bonus
	dim cod_bonus

    'Saldo
    dim saldoTotalBonus


	function getSaldoByCliente(id)

		dim sql, arr, intarr, i
		dim j, arr2, intarr2
		dim html
		dim style
		dim saldo
		dim saldoTotal
		dim sDataResgate

		if len(trim(saldoTotalBonus)) then
			saldoTotal = saldoTotalBonus
		else
			saldoTotal = 0
		end if

		sql = "select distinct(bonus.numero_solicitacao) " & _
			  "from bonus_gerado_clientes as bonus  " & _
			  "left join clientes as cli " & _
			  "on bonus.clientes_idclientes = cli.idclientes " & _
			  "where bonus.clientes_idclientes in (select idclientes from clientes where idclientes = "& id &")"
			'response.write sql & "<hr>"
			'response.end

		call search(sql, arr, intarr)

		if intarr > -1 then
			for i=0 to intarr
				sql = "select " & _
						"pontuacao, " & _
						"pontuacao_atingir, " & _
						"day(data_geracao) as dia_geracao, " & _
						"month(data_geracao) as mes_geracao, " & _
						"year(data_geracao) as ano_geracao, " & _
						"day(data_validade), " & _
						"month(data_validade), " & _
						"year(data_validade), " & _
						"saldo, " & _
						"moeda, " & _
						"day(data_resgate), " & _
						"month(data_resgate), " & _
						"year(data_resgate) " & _
						"from bonus_gerado_clientes where numero_solicitacao = '"&arr(0,i)&"'"
					'response.write sql & "<br />"
					'Response.End
				call search(sql, arr2, intarr2)
				if intarr > -1 then
'						response.write datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) & "<br />"
'						html = html & "<tr>"
'						html = html & "<th colspan=""5"">"&arr(0,i)&"</th>"
'						html = html & "</tr>"
					j=0
					sDataResgate = arr2(10,j)&"/"&arr2(11,j)&"/"&arr2(12,j)
					saldo = arr2(8,j)
					if datediff("d", arr2(5,j)&"/"&arr2(6,j)&"/"&arr2(7,j), now()) < 0 then 'and len(sDataResgate) = 2 and clng(saldo) <> 0 then
						saldoTotal = saldoTotal + clng(saldo)
					end if
				end if
			next
			getSaldoByCliente = saldoTotal
		else
			getSaldoByCliente = saldoTotal
		end if
	end function

	Sub GetCategories()
		Dim sSql, arrCategories, intCategories, i
		Dim sSelected

		sSql = "SELECT idCategorias, descricao FROM Categorias"

		Call search(sSql, arrCategories, intCategories)

		For i=0 To intCategories
			If Categoria = arrCategories(0,i) Then
				sSelected = "selected"
			Else
				sSelected = ""
			End If
			Response.Write "<option value='"&arrCategories(0,i)&"' "&sSelected&">"&arrCategories(1,i)&"</option>"
		Next

	End Sub

	Sub GetGroups()
		Dim sSql, arrGroups, intGroups, i
		Dim sSelected

		sSql = "SELECT " & _
						"[idGrupos], " & _
						"[descricao] " & _
						"FROM [marketingoki2].[dbo].[Grupos]"
		Call search(sSql, arrGroups, intGroups)
		If intGroups > -1 Then
			For i=0 To intGroups
				If Grupo = arrGroups(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value='"&arrGroups(0,i)&"' "&sSelected&">"&arrGroups(1,i)&"</option>"
			Next
		End If
	End Sub

	Sub GetClientById(ID)

		Dim sSql, arrCliente, intCliente, i
		Dim sSqlEndereco, arrClienteEndereco, intClienteEndereco, j
		Dim sSqlEnderecoCep, arrEnderecoCep, intEnderecoCep, k
		Dim sSqlEnderecoColeta, arrEnderecoColeta, intEnderecoColeta, l
		Dim sSqlEnderecoCepColeta, arrEnderecoCepColeta, intEnderecoCepColeta, m
		Dim sSqlPontoColeta, arrPontoColeta, intPontoColeta, n

		sSql = "SELECT " & _
						"A.[idClientes], " & _
						"A.[razao_social], " & _
						"A.[nome_fantasia], " & _
						"A.[cnpj], " & _
						"A.[inscricao_estadual], " & _
						"A.[ddd], " & _
						"A.[telefone], " & _
						"A.[compl_endereco], " & _
						"A.[compl_endereco_coleta], " & _
						"A.[numero_endereco], " & _
						"A.[numero_endereco_coleta], " & _
						"A.[contato_respcoleta], " & _
						"A.[ddd_respcoleta], " & _
						"A.[telefone_respcoleta], " & _
						"A.[numero_sequencial], " & _
						"A.[data_atualizacao_sequencial], " & _
						"A.[typeColect], " & _
						"A.[status_cliente], " & _
						"A.[bonus_type], " & _
						"B.[idCategorias], " & _
						"B.[descricao], " & _
						"C.[idGrupos], " & _
						"C.[descricao], " & _
						"A.[motivo_status], " & _
						"A.[contato_respcoleta], " & _
						"A.[ddd_respcoleta], " & _
						"A.[telefone_respcoleta], " & _
						"A.[tipopessoa], " & _
						"A.[minCartuchos], " & _
						"A.[cod_cli_consolidador], " & _
						"A.[cod_bonus_cli] " & _
						"FROM [marketingoki2].[dbo].[Clientes] AS A " & _
						"LEFT JOIN [marketingoki2].[dbo].[Categorias] AS B " & _
						"ON A.[Categorias_idCategorias] = B.[idCategorias] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Grupos] AS C " & _
						"ON A.[Grupos_idGrupos] = C.[idGrupos] " & _
						"WHERE A.[idClientes] = " & ID

		Call search(sSql, arrCliente, intCliente)
		If intCliente > -1 Then
			For i=0 To intCliente
				RazaoSocial 							= arrCliente(1,i)
				NomeFantasia 							= arrCliente(2,i)
				CNPJ 									= arrCliente(3,i)
				InscEstadual 							= arrCliente(4,i)
				DDD 									= arrCliente(5,i)
				Telefone 								= arrCliente(6,i)
				CompLog 								= arrCliente(7,i)
				CompLogColeta 							= arrCliente(8,i)
				Numero 									= arrCliente(9,i)
				NumeroColeta 							= arrCliente(10,i)
				isColetaDomiciliar 						= arrCliente(16,i)
				StatusCliente 							= arrCliente(17,i)
				Categoria 								= arrCliente(19,i)
				Grupo 									= arrCliente(21,i)
				MotivoStatus 							= arrCliente(23,i)
				ContatoRespColeta 						= arrCliente(24,i)
				DDDContatoRespColeta 					= arrCliente(25,i)
				TelefoneContatoRespColeta				= arrCliente(26,i)
				TipoPessoa 								= arrCliente(27,i)
				MinCartuchos 							= arrCliente(28,i)
				CodCliConsolidador						= arrCliente(29,i)
				cod_bonus								= arrCliente(30,i)
			Next
		End If


		sSqlEndereco = "SELECT " & _
										"[cep_consulta_idcep_consulta], " & _
										"[Clientes_idClientes], " & _
										"[isEnderecoColeta], " & _
										"[isEnderecoComum] " & _
										"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] " & _
										"WHERE [Clientes_idClientes] = " & ID & " AND " & _
										"[isEnderecoComum] = 1"

		Call search(sSqlEndereco, arrClienteEndereco, intClienteEndereco)

		If intClienteEndereco > -1 Then
			For k=0 To intClienteEndereco
				IDCep = arrClienteEndereco(0,k)
			Next
		Else
			IDCep = 0
		End If

		sSqlEnderecoCep = "SELECT " & _
							"[cep_consulta_idcep_consulta], " & _
							"[cep], " & _
							"[logradouro], " & _
							"[bairro], " & _
							"[municipio], " & _
							"[estado], " & _
							"[isEnderecoColeta], " & _
							"[isEnderecoComum] " & _
							"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] " & _
							"WHERE [cep_consulta_idcep_consulta] = " & IDCep

		Call search(sSqlEnderecoCep, arrEnderecoCep, intEnderecoCep)

		If intEnderecoCep > -1 Then
			For j=0 To intEnderecoCep
				CEP = arrEnderecoCep(1,j)
				LogCliente = arrEnderecoCep(2,j)
				Bairro = arrEnderecoCep(3,j)
				Municipio = arrEnderecoCep(4,j)
				Estado = arrEnderecoCep(5,j)
			Next
		Else
				CEP = ""
				LogCliente = ""
				Bairro = ""
				Municipio = ""
				Estado = ""
		End If

        '
        'não checar mais se é coleta domiciliar = 1
        'peterson aquino - 18-5-2014
        '
	    sSqlEnderecoCepColeta = "SELECT " & _
									"[cep_consulta_idcep_consulta], " & _
									"[cep], " & _
									"[logradouro], " & _
									"[bairro], " & _
									"[municipio], " & _
									"[estado], " & _
									"[isEnderecoColeta], " & _
									"[isEnderecoComum] " & _
									"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] " & _
									"WHERE [Clientes_idClientes] = " & ID & " AND " & _
									"[isEnderecoColeta] = 1"

    	Call search(sSqlEnderecoCepColeta, arrEnderecoCepColeta, intEnderecoCepColeta)

		If intEnderecoCepColeta > -1 Then
			For m=0 To intEnderecoCepColeta
				CEPColeta = arrEnderecoCepColeta(1,m)
				LogColeta = arrEnderecoCepColeta(2,m)
				BairroColeta = arrEnderecoCepColeta(3,m)
				MunicipioColeta = arrEnderecoCepColeta(4,m)
				EstadoColeta = arrEnderecoCepColeta(5,m)
            Next
		Else
				CEPColeta = "[vazio]"
				LogColeta = "[vazio]"
				BairroColeta = "[vazio]"
				MunicipioColeta = "[vazio]"
				EstadoColeta = "[vazio]"
		End If

	End Sub


    '
    'peterson aquino 18-5-2014 id:22
    '
	Sub AprovarCliente(ID)
		Call RequestForm()
		StatusCliente = 1
		Call UpdateClienteAprovacaoAdm()
		Call AprovarContatoMaster(ID)
		Response.Write "<script>alert('Cliente aprovado com sucesso.\nPor favor verifique o contato e a Solicitação feita por este cliente!')</script>"

	  Dim sSql, arrCont, intCont, i
	  Dim Usuario, Senha, Destino

	  sSql = "SELECT " & _
					"[idContatos], " & _
					"[usuario], " & _
					"[senha], " & _
					"[email] " & _
					"FROM [marketingoki2].[dbo].[Contatos] " & _
					"WHERE [Clientes_idClientes] = "&ID&" " & _
					"AND [isMaster] = 1 " & _
					"AND [status_contato] = 1"
		Call search(sSql, arrCont, intCont)
		If intCont > -1 Then
			For i=0 To intCont
				Usuario = arrCont(1,i)
				Senha = arrCont(2,i)
				Destino = arrCont(3,i)
				Call EnviarEmail(True, ID, Usuario, Senha, Destino, "")
			Next
		End If
	End Sub

    '
    'reprova cliente
	Sub ReprovarCliente(ID)
	  Call RequestForm()
	  StatusCliente = 2 'status=2 é cliente reprovado
      Dim sSql
	  Dim arrCont, intCont, i
	  Dim Usuario, Senha, Destino

		MotivoStatus = Request.Form("txtMotivoStatus")

        'chama a mesma rotina de aprovação
		Call UpdateClienteAprovacaoAdm()

	  if left(request.ServerVariables("LOCAL_ADDR"), 2) = "10" then
		  sSql = "SELECT " & _
				"[idContatos], " & _
				"[usuario], " & _
				"[senha], " & _
				"[email] " & _
				"FROM [marketingoki2].[dbo].[Contatos] " & _
				"WHERE [Clientes_idClientes] = "&ID&" " & _
				"AND [isMaster] = 1"
			Call search(sSql, arrCont, intCont)
			If intCont > -1 Then
				For i=0 To intCont
					Usuario = arrCont(1,i)
					Senha = arrCont(2,i)
					Destino = arrCont(3,i)
					Call EnviarEmail(False, ID, Usuario, Senha, Destino, MotivoStatus)
				Next
			End If
		end if
		dim arrdelete, intdelete, idelete
		sSql = "SELECT [Solicitacao_coleta_idSolicitacao_coleta] " & _
			   "FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] where [Clientes_idClientes] = " & ID
'		response.write sSql
'		response.end
		call search(sSql, arrdelete, intdelete)
		if intdelete > -1 then
			for idelete=0 to intdelete
				sSql = "update [marketingoki2].[dbo].[Solicitacao_coleta] " & _
						"set status_coleta_idstatus_coleta = 3 " & _
						"where idsolicitacao_coleta = " & arrdelete(0,idelete)
'				response.write sSql
'				response.end
				call exec(sSql)
			next
		end if
'		sSql = "DELETE FROM cep_consulta_has_Clientes WHERE Clientes_idClientes = " & ID
'		Call exec(sSql)
'		sSql = "DELETE FROM Solicitacao_coleta_has_Clientes WHERE Clientes_idClientes = " & ID
'		Call exec(sSql)
'		sSql = "DELETE FROM Contatos WHERE Clientes_idClientes = " & ID
'		Call exec(sSql)
'		sSql = "DELETE FROM Clientes WHERE idClientes = " & ID
'		Call exec(sSql)
	End Sub


    '
    'envia e-mail de aprovação do cadastro
	Sub EnviarEmail(bAprovado, ID, Usuario, Senha, Destino, Motivo)
		Dim MsgBody

		If bAprovado Then
			MsgBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
								"<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
								"<head> " & _
								"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
								"<title>Email Okidata</title> " & _
								"</head> " & _
								"<body> " & _
									"<div id=""container"" align=""center""> " & _
										"<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
											"<p>Prezado(a) cliente;<br /> " & _
											"Primeiramente agradecemos pela sua colaboração e ingresso ao OKI Eco Program, o Programa de Coleta e Destinação de Cartuchos OKI.<br /><br /><br /> " & _
											"O seu cadastro foi aprovado, acesse a sua interface pelo site (<a href=""http://www.sustentabilidadeoki.com.br"">clique aqui</a>), com o login: <b style=""color:#990000"">"&Usuario&"</b> e senha: <b style=""color:#990000"">"&Senha&"</b>.<br /><br /> " & _
											"Caso tenha dúvidas, por favor fale conosco.<br />" & _
											"Grande São Paulo +55 (11) 3444-6747 <br />" & _
											"Demais localidades 0800-115577 <br />" & _
											"Horário de atendimento: Segunda a Sábado - das 8:00 às 20:00 <br />" & _
											"Atenciosamente;<br />" & _
											"<b style=""color:#990000"">OKI Printing Solutions</b> " & _
										"</div> " & _
										"<div id=""bottom"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;width:748px;""> " & _
											"<p><b>CONFIDENCIALIDADE DO CORREIO ELETRÔNICO</b> " & _
											"Esta mensagem, incluindo seus anexos, pode conter informação confidencial " & _
											"e/ou privilegiada. Caso você tenha recebido este e-mail por engano, não " & _
											"utilize, copie ou divulgue as informações nele contidas. E, por favor, avise " & _
											"imediatamente o remetente, respondendo ao e-mail, e em seguida apague-o.</p> " & _
											"<p><b>DISCLAIMER</b> " & _
											"This message, including its attachments, may contain confidential and/or " & _
											"privileged information. If you received this email by mistake, do not use, " & _
											"copy or disseminate any information here in contained. Please notify us " & _
											"immediately by replying to the sender and then delete it.</p> " & _
										"</div> " & _
									"</div> " & _
								"</body> " & _
								"</html>"
		Else
			MsgBody = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
								"<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
								"<head> " & _
								"<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
								"<title>Email Okidata</title> " & _
								"</head> " & _
								"<body> " & _
									"<div id=""container"" align=""center""> " & _
										"<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
											"<p>Prezado(a) cliente;<br /> " & _
											"Primeiramente agradecemos pelo seu interesse em ingressar ao OKI Eco Program, o Programa de Coleta e Destinação de Cartuchos OKI.<br /><br /><br /> " & _
											"<b style=""color:#990000"">A seu cadastro foi reprovado, devido ao motivo::</b> "&Motivo&"<br /><br /> " & _
											"Caso tenha dúvidas, por favor fale conosco.<br />" & _
											"Grande São Paulo +55 (11) 3444-6747 <br />" & _
											"Demais localidades 0800-115577 <br />" & _
											"Horário de atendimento: Segunda a Sábado - das 8:00 às 20:00 <br />" & _
											"Atenciosamente;<br />" & _
											"<b style=""color:#990000"">OKI Printing Solutions</b> " & _
										"</div> " & _
										"<div id=""bottom"" style=""font-family:Verdana, Arial, Helvetica, sans-serif;font-size:10px;width:748px;""> " & _
											"<p><b>CONFIDENCIALIDADE DO CORREIO ELETRÔNICO</b> " & _
											"Esta mensagem, incluindo seus anexos, pode conter informação confidencial " & _
											"e/ou privilegiada. Caso você tenha recebido este e-mail por engano, não " & _
											"utilize, copie ou divulgue as informações nele contidas. E, por favor, avise " & _
											"imediatamente o remetente, respondendo ao e-mail, e em seguida apague-o.</p> " & _
											"<p><b>DISCLAIMER</b> " & _
											"This message, including its attachments, may contain confidential and/or " & _
											"privileged information. If you received this email by mistake, do not use, " & _
											"copy or disseminate any information here in contained. Please notify us " & _
											"immediately by replying to the sender and then delete it.</p> " & _
										"</div> " & _
									"</div> " & _
								"</body> " & _
								"</html>"
		End If


        Dim objCDOSYSMail
        Dim objCDOSYSCon

        'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
        Set objCDOSYSMail = Server.CreateObject("CDO.Message")

        'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
        Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

    	'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Oki!321!" '"!nfe321!"        'senha
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

		'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "174.37.245.58"

		'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

		'PORTA DO CDO
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

		'TEMPO DE TIMEOUT (EM SEGUNDOS)
		'objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

		'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
		objCDOSYSCon.Fields.update
		Set objCDOSYSMail.Configuration = objCDOSYSCon

		'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
    	objCDOSYSMail.From = "sustentabilidadeoki@sustentabilidadeoki.com.br"

		''''NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
		objCDOSYSMail.To = Destino
        'objCDOSYSMail.CC = "peterson.aquino@hotmail.com"

		'ASSUNTO DA MENSAGEM
		objCDOSYSMail.Subject = "Okidata - Sistema de Gerenciamento de Recolhimento de Suprimentos"

		'CONTEÚDO DA MENSAGEM
		'objCDOSYSMail.TextBody = "Teste do componente CDOSYS"
		'PARA ENVIO DA MENSAGEM NO FORMATO HTML, ALTERE O TextBody PARA HtmlBody
		objCDOSYSMail.HtmlBody = MsgBody

		'ENVIA A MENSAGEM
		objCDOSYSMail.Send

		'DESTRÓI OS OBJETOS
		Set objCDOSYSMail = Nothing
		Set objCDOSYSCon = Nothing

    	'End If
	End Sub

	Sub RequestForm()

		' Endereço do Cliente
		cep_cliente				= request.Form("txtCEPEnderecoComumCliente")
		logradouro_cliente		= request.Form("txtLogradouro")
		bairro_cliente			= request.Form("txtBairro")
		municipio_cliente		= request.Form("txtMunicipio")
		estado_cliente			= request.Form("txtEstado")

		'Informações do Cliente
		'-----------------------
		TipoPessoa	  			= Request.Form("radioPessoa")
		If CInt(TipoPessoa) = 1 Then
			RazaoSocial   		= Request.Form("txtRazaoSocialCliente")
			NomeFantasia  		= Request.Form("txtNomeFantasiaCliente")
			CNPJ 		  		= Request.Form("txtCNPJCliente")
		Else
			RazaoSocial   		= Request.Form("txtRazaoSocialCliente")
			NomeFantasia  		= Request.Form("txtNomeCliente")
			CNPJ 		  		= Request.Form("txtCPFCliente")
		End If
		InscEstadual  			= Request.Form("txtInscricaoEstadualCliente")
		DDD 		  			= Request.Form("txtDDDCliente")
		Telefone 	  			= Request.Form("txtTelefoneCliente")
		Categoria 	  			= Request.Form("cbCategorias")
		'Grupo 		  			= Request.Form("cbGrupos")
		StatusCliente 			= Request.Form("cbStatusCliente")
		MotivoStatus  			= Request.Form("txtMotivoStatus")
		TipoColeta	  			= "1" 'Request.Form("cbTipoColeta") fixo não existe coleta domiciliar.

		'Endereço comum da Empresa
		'--------------------------
		CEP 								= Request.Form("txtCEPEnderecoComumCliente")
		CompLog 							= Request.Form("txtCompLogradouro")
		Numero 								= Request.Form("txtNumero")
		'isColetaDomiciliar 					= Request.Form("hiddenActionIsColetaDomiciliar")
        isColetaDomiciliar 					= 1 'peterson aquino 18-5-2014

		If isColetaDomiciliar = 1 Then

			'Endereço de Coleta
			'-------------------
			CEPColeta 				  	= Request.Form("txtCEPEnderecoComumClienteColeta")
			CompLogColeta 			  	= Request.Form("txtCompLogradouroColeta")
			NumeroColeta  			  	= Request.Form("txtNumeroColeta")
			ContatoRespColeta 		  	= Request.Form("txtRespColeta")
			DDDContatoRespColeta 	  	= Request.Form("txtDDDRespColeta")
			TelefoneContatoRespColeta	= Request.Form("txtTelefoneRespColeta")

			' Endereço de Coleta
			cep_coleta					= request.Form("txtCEPEnderecoComumClienteColeta")
			logradouro_coleta			= request.Form("txtLogradouroColeta")
			bairro_coleta				= request.Form("txtBairroColeta")
			municipio_coleta			= request.Form("txtMunicipioColeta")
			estado_coleta				= request.Form("txtEstadoColeta")
			cod_bonus					= request.form("cbBonus")
		Else

			'Endereço Ponto de Coleta
			'------------------------
			IDPontoColeta   = Request.Form("txtIDPontoDeColeta")
			NomePontoColeta = Request.Form("txtPontoDeColeta")
			CNPJPontoColeta = Request.Form("txtCNPJPontoDeColeta")
		End If
		CodCliConsolidador  = Request.Form("cbCodConsolidador")
		MinCartuchos		= Request.Form("txtQtdCartuchos")
		If CStr(MinCartuchos) = "" Then
			MinCartuchos = 0
		ElseIf CInt(MinCartuchos) = 0 Then
			MinCartuchos = 0
		End If

		if len(trim(NumeroColeta)) = 0 then
			NumeroColeta = 0
		end if
		if len(trim(Numero)) = 0 then
			Numero = 0
		end if
	End Sub

	Sub UpdateClienteAprovacaoAdm()
		Dim oCommand
		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.ActiveConnection = oConn
		oCommand.CommandTimeout = 200
		oCommand.CommandType = 4
		'oCommand.CommandText = "sp_UpdateAprovacaoCliente"
        oCommand.CommandText = "sp_UpdateAprovacaoClienteLc"
        oCommand.Parameters("@RazaoSocial")              = RazaoSocial
		oCommand.Parameters("@IDCliente") 				 = Request.Form("hiddenIDCliente")
		oCommand.Parameters("@NomeFantasia") 			 = NomeFantasia
		oCommand.Parameters("@InscEstadual") 			 = InscEstadual
		oCommand.Parameters("@DDD") 					 = DDD
		oCommand.Parameters("@Telefone") 				 = Telefone
		oCommand.Parameters("@CompEndereco") 			 = CompLog
		oCommand.Parameters("@CompEnderecoColeta")		 = CompLogColeta
		oCommand.Parameters("@NumEndereco") 			 = Numero
		oCommand.Parameters("@NumEnderecoColeta")		 = NumeroColeta
		oCommand.Parameters("@Categoria") 				 = Categoria
		'oCommand.Parameters("@Grupo") 					 = Grupo 'alterado peterson aquino 18-5-2014
		oCommand.Parameters("@CEP") 					 = CEP
		oCommand.Parameters("@COLETADOMICILIAR")		 = isColetaDomiciliar
		oCommand.Parameters("@CEPColeta") 				 = CEPColeta
		oCommand.Parameters("@RespColeta") 				 = ContatoRespColeta
		oCommand.Parameters("@DDDRespColeta") 			 = DDDContatoRespColeta
		oCommand.Parameters("@TelefoneRespColeta")		 = TelefoneContatoRespColeta
		oCommand.Parameters("@Status") 					 = StatusCliente
		oCommand.Parameters("@MotivoStatus") 			 = MotivoStatus
		oCommand.Parameters("@TipoColeta")				 = TipoColeta
		oCommand.Parameters("@MinCartuchos")			 = MinCartuchos
		oCommand.Parameters("@CodClienteConsolidador")	 = CodCliConsolidador
		oCommand.Parameters("@logradouro_cliente")		 = logradouro_cliente
		oCommand.Parameters("@logradouro_cliente_coleta")= logradouro_coleta
		oCommand.Parameters("@bairro_cliente")			 = bairro_cliente
		oCommand.Parameters("@bairro_cliente_coleta")	 = bairro_coleta
		oCommand.Parameters("@municipio_cliente")		 = municipio_cliente
		oCommand.Parameters("@municipio_cliente_coleta") = municipio_coleta
		oCommand.Parameters("@estado_cliente")			 = estado_cliente
		oCommand.Parameters("@estado_cliente_coleta")	 = estado_coleta
		oCommand.Parameters("@cod_bonus")				 = cod_bonus
		oCommand.Execute()
		Set oCommand = Nothing
'		response.write cint(Request.Form("cbTransp")) & "<br>"
'		response.end
		if cint(Request.Form("cbTransp")) = -1 or cint(Request.Form("cbTransp")) = 0 then
			sSql = "UPDATE Clientes " & _
							"SET " & _
							"[Transportadoras_idTransportadoras] = NULL " & _
							"WHERE [idClientes] = " & Request.Form("hiddenIDCliente")
		else
			sSql = "UPDATE Clientes " & _
							"SET " & _
							"[Transportadoras_idTransportadoras] = "&Request.Form("cbTransp")&" " & _
							"WHERE [idClientes] = " & Request.Form("hiddenIDCliente")
		end if
'		response.write sSql & "<br />"
'		response.end
		call exec(sSql)
	End Sub

	Sub AprovarContatoMaster(ID)
		Dim sSql
		sSql = "UPDATE [marketingoki2].[dbo].[Contatos] " & _
						"SET [status_contato] = 1 " & _
						"WHERE [Clientes_idClientes] = "&ID&" " & _
						"AND [isMaster] = 1"
		Call exec(sSql)
	End Sub

	Function GetIDTransp()
		Dim sSql, arrId, intId, i
		Dim Ret
		sSql = "SELECT Transportadoras_idTransportadoras FROM Clientes WHERE idClientes = " & Request.QueryString("idcliente")
'		Response.Write sSql
'		Response.End()
		Call search(sSql, arrId, intId)
		If intId > -1 Then
			For i=0 To intId
				Ret = arrId(0,i)
			Next
		End If
		GetIDTransp = Ret
	End Function

	Sub GetTransp()
		Dim sSql, arrTransp, intTransp, i
		Dim sSelected
		sSql = "SELECT [idTransportadoras] " & _
					 ",[nome_fantasia] " & _
					 "FROM [marketingoki2].[dbo].[Transportadoras] " & _
					 "WHERE [status] = 1"
		Call search(sSql, arrTransp, intTransp)
		If intTransp > -1 Then
			For i=0 To intTransp
				If GetIDTransp() = arrTransp(0,i) Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value="&arrTransp(0,i)&" "&sSelected&">"&arrTransp(1,i)&"</option>"
			Next
		End If
	End Sub

	Sub GetClientes()
		Dim sSql, arrCli, intCli, i
		Dim sSelected

		sSql = "SELECT idClientes, substring( upper(nome_fantasia),1,30) as nome_fantasia " & _
						"FROM [marketingoki2].[dbo].[Clientes] " & _
						"WHERE [status_cliente] = 1" & _
                        " order by nome_fantasia "
		Call search(sSql, arrCli, intCli)
		If intCli > -1 Then
			For i=0 To intCli
				If arrCli(0,i) = CodCliConsolidador Then
					sSelected = "selected"
				Else
					sSelected = ""
				End If
				Response.Write "<option value="&arrCli(0,i)&" "&sSelected&">"&arrCli(1,i)&"</option>"
			Next
		End If
	End Sub

	sub getBonus()
		dim sql, arr, intarr, i
		dim sSelected
		sql = "SELECT [cod_bonus] FROM [marketingoki2].[dbo].[Cadastro_Bonus] where aplicacao = 'CLI'"
		call search(sql, arr, intarr)
		if intarr > -1 then
			for i=0 to intarr
				if cod_bonus = arr(0,i) then
					sSelected = "selected"
				else
					sSelected = ""
				end if
				response.write "<option value="&arr(0,i)&" "&sSelected&">"&arr(0,i)&"</option>"
			next
		end if
	end sub

	If Request.QueryString("showclient") Then
		Call GetClientById(Request.QueryString("idcliente"))
	End If
	If Request.Form("hiddenActionForm") = "APROVAR" Then
		Call AprovarCliente(Request.Form("hiddenIDCliente"))
		Response.Redirect "frmCadastroClienteAdm.asp"
	End If
	If Request.Form("hiddenActionForm") = "REPROVAR" Then
		Call ReprovarCliente(Request.Form("hiddenIDCliente"))
		Response.Redirect "frmCadastroClienteAdm.asp"
	End If
	If Request.Form("hiddenActionForm") = "SALVAR" Then
		Call RequestForm()
		Call UpdateClienteAprovacaoAdm()
		Response.Redirect "frmCadastroClienteAdm.asp"
	End If

'	Response.Write GetIDTransp()

%>
<html>
<head>
<script src="../adm/js/frmCadastroClienteAdmLc.js" language="javascript"></script>
<link rel="stylesheet" type="text/css" href="../css/geral.css">
<title><%=TITLE%></title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style type="text/css">
        .oki-texto-titulo {
            text-align:right;
            font-family:Verdana;
            font-size:10px;
            width:258px;
        }
        .oki-texto-tituloTitle {
            text-align:center;
            font-family:Verdana;
            font-size:10px;
            width:258px;
        }
        .oki-texto-cabec {
            margin: 0px;
            padding: 0px;
            border: 0px;
            margin-left:10%;
            font-size: 15px;
            vertical-align: baseline;
            background-color: transparent; font-family: Arial, sans-serif; line-height: 24px; color: rgb(244, 121, 32); font-style: normal; font-variant: normal; letter-spacing: normal; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;
        }
    </style>
</head>

<body onload="tipopessoa()">
    <div id="container">
        <!--#include file="inc/i_header.asp" -->
        <div id="conteudo">
            <table cellspacing="0" cellpadding="0" width="775">
                <form action="frmEditCadastroClienteLc.asp" name="frmEditCadastroClienteLc" method="POST">
                    <input type="hidden" name="hiddenIDCliente" value="<%=Request.QueryString("idcliente")%>" />
                    <input type="hidden" name="hiddenIntEnderecoCep" value="" />
                    <input type="hidden" name="hiddenActionIsColetaDomiciliar" value="<%=isColetaDomiciliar%>" />
                    <input type="hidden" name="hiddenActionManagerProve" value="" />
                    <input type="hidden" name="hiddenActionForm" value="" />
                    <input type="hidden" name="hiddenIntQtdCartuchos" value="<%=MinCartuchos%>" />
                <tr>

                <td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
                <td id="conteudo">


                        <!--link para acesso ao cadastro de contatos-->
                        <table style="width: 100%;">
                            <tr>
                                <td id="explaintitle" align="center" colspan="7">Atalhos</td>
                            </tr>
                            <tr>
                                <td align="center" colspan="7"><a class="linkOperacional" href="javascript:window.location.href='frmContatosPorCliente.asp?idcliente=<%=Request.QueryString("idcliente")%>';">Atualiza Contatos</a></td>
                            </tr>
                        </table>


                        <!--cadastro do cliente-->
                        <table style="width: 100%;">
                            <tr>
                                <td id="explaintitle" align="center" colspan="7">Cadastro do Cliente - Saldo do Bônus:<%=getSaldoByCliente(Request.QueryString("idcliente"))%> </td>
                            </tr>
                            <tr>
                                <td align="right" colspan="7"><a class="linkOperacional" href="javascript:window.location.href='frmCadastroClienteAdm.asp';">&laquo Voltar</a></td>
                            </tr>
                        </table>

                        <table id="editTableCadClienteAlpha" style="width:100%;vertical-align:central;border-collapse: separate; border-spacing: 3px;">
                            <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                            </tr>
                            <tr>
                                <td style="text-align:right;"></td>
                                <td style="text-align:right;">Pessoa Física<input type="radio" name="radioPessoa" disabled="disabled" value="0" <% If CInt(TipoPessoa) = 0 Then %>checked<% End If %> onclick="tipopessoa()" /></td>
                                <td style="text-align:left;">Pessoa Jurídica<input type="radio" name="radioPessoa" disabled="disabled" value="1" <% If CInt(TipoPessoa) = 1 Then %>checked<% End If %> onclick="tipopessoa()" /></td>
                                <td style="text-align:right;"></td>
                            </tr>
                        </table>

                        <table id="pessoajuridica" style="border-collapse: separate; border-spacing: 3px; ">
                            <tr id="razaosocial">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Razão Social:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtRazaoSocialCliente" value="<%=RazaoSocial%>" size="40" /></td>
                            </tr>
                            <tr id="nomefantasia">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Nome Fantasia:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtNomeFantasiaCliente" value="<%=NomeFantasia%>" size="40" /></td>
                            </tr>
                            <tr id="cnpj">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">CNPJ:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" disabled="disabled" name="txtCNPJCliente" value="<%=CNPJ%>" size="22" maxlength="18" /></td>
                            </tr>
                            <tr id="inscestadual">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Inscrição Estadual:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtInscricaoEstadualCliente" value="<%=InscEstadual%>" size="18" maxlength="15" /></td>
                            </tr>
                        </table>

                        <table cellpading="3" cellspacing="2" id="pessoafisica">
                            <tr id="cpf">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">CPF:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" disabled="disabled" name="txtCPFCliente" value="<%=CNPJ%>" size="22" maxlength="18" /></td>
                            </tr>
                            <tr id="nomecliente">
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Nome:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" name="txtNomeCliente" value="<%=NomeFantasia%>" size="40" /></td>
                            </tr>
                        </table>

                        <table cellpading="3" cellspacing="2">
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">DDD:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" name="txtDDDCliente" value="<%=DDD%>" size="3" maxlength="3" /></td>
                            </tr>
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Telefone:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" name="txtTelefoneCliente" value="<%=Telefone%>" size="10" maxlength="8" /></td>
                            </tr>
                        </table>

                        <table cellpading="3" cellspacing="2" id="demais">
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">
                                    <label id="numidtransportadora">Cód. Consolidador: </label>
                                </td>
                                <td colspan="3">
                                    <select name="cbCodConsolidador" class="select">
                                        <option value="0">[Nenhum]</option>
                                        <%Call GetClientes()%>
                                    </select>
                                    <img src="img/cadcliente.gif" class="imgexpandeinfo" width="22" height="22" align="absmiddle" alt="Clientes" onclick="window.open('frmSearchCodConsolidador.asp','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
                                </td>
                            </tr>
                            <%If isColetaDomiciliar = 1 Then%>
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">
                                    <label id="numidtransportadora">Transportadora: </label>
                                </td>
                                <td colspan="3">
                                    <select name="cbTransp" class="select">
                                        <option value="-1">[Selecione]</option>
                                        <%Call GetTransp()%>
                                    </select>
                                    <img src="img/transportadoras.gif" class="imgexpandeinfo" width="25" height="25" align="absmiddle" alt="Buscar Transportadora" onclick="window.open('frmSearchTranspCliente.asp?idcli=<%= Request.QueryString("idcliente") %>','','width=410,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
                                </td>
                            </tr>
                            <%End If%>
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Categoria:</td>
                                <td align="left" colspan="3">
                                    <select name="cbCategorias" class="select">
                                        <option value="-1">[Selecione]</option>
                                        <%Call GetCategories()%>
                                    </select>
                                    <img src="img/categoria.gif" class="imgexpandeinfo" alt="Buscar Categoria" width="20" height="20" align="absmiddle" onclick="window.open('frmSearchCategoriaEditCliente.asp','','width=500,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
                                </td>
                            </tr>
                            <%If isColetaDomiciliar = 1 Then%>
                            <tr>
                                <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 258px;">Cód. Bônus:</td>
                                <td align="left" colspan="3">
                                    <select name="cbBonus" class="select" style="width: 200px;">
                                        <option value="">[Selecione]</option>
                                        <%Call getBonus()%>
                                    </select>
                                    <img src="img/bonus.gif" class="imgexpandeinfo" width="23" height="23" align="absmiddle" alt="Buscar Bônus" onclick="window.open('frmsearchbonuscliente.asp','','width=600,height=300,scrollbars=no,status=no,location=no,toolbar=no,menubar=no')" />
                                </td>
                            </tr>
                            <%end if%>
                            <tr>
                                <td colspan="4" align="center">&nbsp;
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" class="oki-texto-cabec"><b>&nbsp; Endereço do Cliente</b></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">CEP:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" name="txtCEPEnderecoComumCliente" value="<%=CEP%>" size="10" maxlength="8" /></td>
                                <td align="left">
                                    <img src="img/buscar.gif" class="imgexpandeinfo" align="middle" onclick="getEndereco()" />
                                    <!--<input id="btnBuscarCepColeta" alt="Buscar CEP" class="BotaoBuscar" name="btnBuscarCepColeta"  onclick="getEndereco() type="button" value="" /></td>-->
                                <td align="left" class="rotulos">
											<!--<img src="img/buscar.gif" class="imgexpandeinfo" align="middle" onclick="getEndereco()" />-->
                                    Preencha o CEP somente com 8 números Ex: 99999999</td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Logradouro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtLogradouro" value="<%=LogCliente%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Complemento Logradouro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtCompLogradouro" value="<%=CompLog%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Número:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtNumero" value="<%=Numero%>" size="10" maxlength="8" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Bairro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtBairro" value="<%=Bairro%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Município:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtMunicipio" value="<%=Municipio%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Estado:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtEstado" value="<%=Estado%>" size="2" /></td>
                            </tr>
                            <tr>
                                <td colspan="4" align="center">&nbsp;
                                </td>
                            </tr>

                            <tr>
                                <td colspan="4" class="oki-texto-cabec"><b>&nbsp; Endereço padrão para coleta domiciliar</b></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">CEP:</td>
                                <td align="left">
                                    <input type="text" class="textreadonly" name="txtCEPEnderecoComumClienteColeta" value="<%=CEPColeta%>" size="10" maxlength="8" />&nbsp;&nbsp;</td>
                                <td align="left">
                                    <img src="img/buscar.gif" class="imgexpandeinfo" align="middle" onclick="getEnderecoColeta()" />
                                    <!--<input id="btnBuscarCepColeta0" alt="Buscar CEP" button"="" class="BotaoBuscar" name="btnBuscarCepColeta0" onclick="getEnderecoColeta() type=" value="" /></td>-->
                                <td align="left" class="rotulos">
											<!--<img src="img/buscar.gif" class="imgexpandeinfo" align="middle" onclick="getEndereco()" />-->
                                    Preencha o CEP somente com 8 números Ex: 99999999</td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Logradouro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtLogradouroColeta" value="<%=LogColeta%>" size="40" />
                                    <!--<img src="img/buscar.gif" class="imgexpandeinfo" align="middle" onclick="getEnderecoColeta()" /></td>-->
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Complemento Logradouro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtCompLogradouroColeta" value="<%=CompLogColeta%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Número:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtNumeroColeta" value="<%=NumeroColeta%>" size="10" maxlength="8" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Bairro:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtBairroColeta" value="<%=BairroColeta%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Município:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtMunicipioColeta" value="<%=MunicipioColeta%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Estado:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtEstadoColeta" value="<%=EstadoColeta%>" size="2" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Resp. Coleta:</td>
                                <td align="left" colspan="3">
                                    <input name="txtRespColeta" type="text" class="textreadonly" id="txtRespColeta" value="<%=ContatoRespColeta%>" size="40" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">DDD Resp. Coleta:</td>
                                <td align="left" colspan="3">
                                    <input name="txtDDDRespColeta" type="text" class="textreadonly" id="txtDDDRespColeta" value="<%=DDDContatoRespColeta%>" size="2" maxlength="3" /></td>
                            </tr>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Telefone Resp. Coleta:</td>
                                <td align="left" colspan="3">
                                    <input type="text" class="textreadonly" name="txtTelefoneRespColeta" value="<%=TelefoneContatoRespColeta%>" size="10" maxlength="8" /></td>
                            </tr>
                            <tr id="motivostatus">
                                <td colspan="4" align="center">
                                &nbsp;
                            </tr>

                            <tr id="motivostatustext">
                                <td align="right" class="oki-texto-titulo">Motivo Status:</td>
                                <td align="left" colspan="3">
                                    <textarea name="txtMotivoStatus" style="height: 43px; width: 299px"><%=MotivoStatus%></textarea></td>
                            </tr>
                            <%If StatusCliente <> 0 Then%>
                            <tr>
                                <td align="right" class="oki-texto-titulo">Status:</td>
                                <td align="left" colspan="3">
                                    <select name="cbStatusCliente" class="select">
                                        <option value="2" <%If StatusCliente = 2 Then%>selected="selected" <%End If%>>Rejeitado</option>
                                        <option value="1" <%If StatusCliente = 1 Then%>selected="selected" <%End If%>>Aprovado</option>
                                        <option value="3" <%If StatusCliente = 3 Then%>selected="selected" <%End If%>>Inativo</option>
                                    </select>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="4" align="right">
                                    <input type="button" name="btnSalvar" class="btnformMaior" value="Salvar" onclick="validar()" /></td>
                            </tr>
                            <%Else%>
                            <tr>
                                <td colspan="4" align="right">
                                    <input type="button" name="btnAprovar" class="btnformMaior" value="Aprovar" onclick="aprovar()" />
                                    <input type="button" name="btnReprovar" class="btnformMaior" value="Rejeitar" onclick="reprovar()" />
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
