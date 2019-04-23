<!--#include file="_config/_config.asp" -->
<%Call open()%>
<%Call getSessionUser()%>
<%

	'// Ponto de Coleta
  '-----------------------------
	Dim QuantidadeCartuchos
	Dim NumeroSolicitacaoColeta
	'-----------------------------
	Dim IDPontoColeta
	Dim NomeFantasiaPontoColeta
	Dim LogradouroPontoColeta
	'-----------------------------

	'// Coleta Domiciliar
	'-----------------------------
	Dim CEPColeta
	Dim LogradouroColeta
	Dim BairroColeta
	Dim MunicipioColeta
	Dim EstadoColeta
	Dim TipoLogradouroColeta
	Dim CompLogradouro
	Dim NumeroColeta
	Dim ContatoRespColeta
	Dim DDDContatoRespColeta
	Dim TelefoneContatoRespColeta
	Dim RamalContatoRespColeta
	Dim DepContatoRespColeta
	'-----------------------------

	Dim NumColEnvio
    Dim Observacao
    Dim cDestino

    Dim NumCol

	Sub RequestForm()

		QuantidadeCartuchos 				= Request.Form("txtQtdCartuchos")
		IDPontoColeta 						= Request.Form("hiddenIntChangePontoColeta")
		CEPColeta 							= Request.Form("txtCepColeta")
		CompLogradouro 						= Request.Form("txtCompLogradouroColeta")
		NumeroColeta 						= Request.Form("txtNumeroColeta")
		ContatoRespColeta 					= Request.Form("txtRespColContato")
		DDDContatoRespColeta 				= Request.Form("txtDDDContatoRespColeta")
		TelefoneContatoRespColeta			= Request.Form("txtTelefoneContatoRespColeta")
		RamalContatoRespColeta				= Request.Form("txtRamalContatoRespColeta")
		DepContatoRespColeta				= Request.Form("txtDepContatoRespColeta")
		Observacao                          = Request.Form("txtObservacao")

		' Informações do endereço de coleta da empresa
		LogradouroColeta					= request.Form("txtLogradouroColeta")
		BairroColeta						= request.Form("txtBairroColeta")
		MunicipioColeta						= request.Form("txtMunicipioColeta")
		EstadoColeta						= request.Form("txtEstadoColeta")
	End Sub

    '
    'peterson aquino 18-5-2014
    Sub EnviaEmailResp(Destino)

        Dim objCDOSYSMail
        Dim objCDOSYSCon
	    Dim MsgBody

        MsgBody = 	"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd""> " & _
				    "<html xmlns=""http://www.w3.org/1999/xhtml""> " & _
				    "<head> " & _
				    "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /> " & _
				    "<title>Email Okidata</title> " & _
				    "</head> " & _
				    "<body> " & _
					    "<div id=""container"" align=""center""> " & _
						    "<div id=""conteudo"" style=""width:748px;font-family:Verdana, Arial, Helvetica, sans-serif;font-size:11px;"" > " & _
							    "<p>Prezado(a) " & Session("NomeContato")  & "<br /> " & _
							    "A solicitação número <b>" & NumCol & "</b> foi incluída com sucesso e está em análise.<br /><br /><br /> " & _
                                "Assim que tivermos uma posição entraremos em contato.<br /><br /> " & _
							    "Desde já nós agradecemos pelo contato.<br /><br /> " & _
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

		'CRIA A INSTÂNCIA COM O OBJETO CDOSYS
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")

		'CRIA A INSTÂNCIA DO OBJETO PARA CONFIGURAÇÃO DO SMTP
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")

		'SERVIDOR DE SMTP, USE smtp.SeuDominio.com OU smtp.hostsys.com.br
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.okidata.com.br"'"mail.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "nfe@okidata.com.br"'"sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!nfe321!"'"Oki7080! " '"!nfe321!"        'senha
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    	'PORTA PARA COMUNICAÇÃO COM O SERVIÇO DE SMTP
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

		'PORTA DO CDO
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

		'TEMPO DE TIMEOUT (EM SEGUNDOS)
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30

		'ATUALIZA A CONFIGURAÇÃO DO CDOSYS PARA ENVIO DO E-MAIL
		objCDOSYSCon.Fields.update

		Set objCDOSYSMail.Configuration = objCDOSYSCon

		'NOME DO REMETENTE, E-MAIL DO REMETENTE
		objCDOSYSMail.From = "sustentabilidadeoki@sustentabilidadeoki.com.br"

		'NOME DO DESINATÁRIO, E-MAIL DO DESINATÁRIO
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

    End Sub


	Sub SubmitForm()
	    If Request.ServerVariables("HTTP_METHOD") = "POST" Then
            Call RequestForm()

            If Request.Form("hiddenActionForm")         = 1 Then
                Response.Redirect "frmOperacionalCliente.asp"
            ElseIf Request.Form("hiddenActionForm")     = 3 Then
                Call AddSolColeta()
            Else
                Call AddSolColeta()
		    End If

            'envia e-mail
            cDestino = Session("Email")
            EnviaEmailResp(cDestino)
            .write "<script>alert('oi');</script>"
    		Response.Write "<script>alert('Você será direcionado para efetuar a emissão do relatório de coleta. Após a impressão, sua solicitação será finalizada e breve entraremos em contato para providenciar a coleta!');</script>"
		    Response.redirect "frmCartaDoacaoNF.asp?Acao=0&IdSolicitacaoColeta=" & NumCol & "&Adm=1&TipoColeta=" & Session("isColetaDomiciliar")

        End if

	End Sub

	Sub AddSolColeta()

		Dim oCommand, rs
		Dim NumeroSequencial
		Dim NumeroSolicitacaoColeta
		Dim IdentifiedCharacterColeta
		Dim DateMonthColeta
		Dim DateYearColeta

'		response.write IDPontoColeta
'		response.End()

		Set oCommand = Server.CreateObject("ADODB.Command")
		Set rs = Server.CreateObject("ADODB.Recordset") 'jadilson
		oCommand.CommandTimeout     = 200
		oCommand.ActiveConnection   = oConn
		oCommand.CommandType        = 4
		'oCommand.CommandText       = "sp_AddSolicitacaoColeta"  'comentado peterson 5/5/2014
		oCommand.CommandText        = "sp_AddSolColLc"

		If Session("isColetaDomiciliar") = 1 Then
			oCommand.Parameters("@IDPontoColeta")			= 0
			IdentifiedCharacterColeta						= "C"
		Else
			IdentifiedCharacterColeta						= "E"
			If Request.Form("hiddenIntChangePontoColeta") <> "" Then
				oCommand.Parameters("@IDPontoColeta")		= CInt(IDPontoColeta)
			Else
				oCommand.Parameters("@IDPontoColeta")		= CInt(Session("IDPontoColeta"))
			End If
		End If

		DateMonthColeta										= Month(Now())
		If Len(DateMonthColeta) = 1 Then
			DateMonthColeta									= "0" & DateMonthColeta
		End If
		DateYearColeta										= Right(Year(Now()), 2)
		NumeroSolicitacaoColeta								= IdentifiedCharacterColeta & DateYearColeta & DateMonthColeta
		NumeroSequencial									= getSequencial(False)
		NumeroSolicitacaoColeta								= NumeroSolicitacaoColeta & NumeroSequencial
		NumeroSolicitacaoColeta								= getDigitoControle(NumeroSolicitacaoColeta)

		oCommand.Parameters("@NumeroSolicitacaoColeta")		= NumeroSolicitacaoColeta
		oCommand.Parameters("@isColetaDomiciliar")			= CInt(Session("isColetaDomiciliar"))
		oCommand.Parameters("@QtdCartuchos")				= CInt(QuantidadeCartuchos)
		oCommand.Parameters("@IDClient")					= CInt(Session("IDCliente"))
		oCommand.Parameters("@IDContato")					= CInt(Session("IDContato"))

		'Peterson 5/5/2014
		oCommand.Parameters("@coleta_cep")					= Request.Form("txtCepColeta")
		oCommand.Parameters("@coleta_logradouro")			= Request.Form("txtLogradouroColeta")
		oCommand.Parameters("@coleta_complemento")			= Request.Form("txtCompLogradouroColeta")
		oCommand.Parameters("@coleta_numero")				= Request.Form("txtNumeroColeta")
		oCommand.Parameters("@coleta_bairro")				= Request.Form("txtBairroColeta")
		oCommand.Parameters("@coleta_municipio")			= Request.Form("txtMunicipioColeta")
		oCommand.Parameters("@coleta_estado")				= Request.Form("txtEstadoColeta")
		oCommand.Parameters("@coleta_contato")				= Request.Form("txtRespColContato")
		oCommand.Parameters("@coleta_ddd")					= Request.Form("txtDDDContatoRespColeta")
		oCommand.Parameters("@coleta_telefone")				= Request.Form("txtTelefoneContatoRespColeta")
		oCommand.Parameters("@coleta_ramal")				= Request.Form("txtRamalContatoRespColeta")
		oCommand.Parameters("@coleta_depto")				= Request.Form("txtDepContatoRespColeta")
        oCommand.Parameters("@Observacao")                  = Request.Form("txtObservacao")

        NumCol = NumeroSolicitacaoColeta

		rs.Open oCommand

		Set oCommand = Nothing

	End Sub

	Sub GetListPontoColeta()

		Dim sSql, arrListPontoColeta, intListPontoColeta, i

'		sSql = "SELECT " & _
'						"[A].[Pontos_coleta_idPontos_coleta], " & _
'						"[B].[Nome_fantasia], " & _
'						"[B].[Numero_endereco], " & _
'						"[B].[Complemento_endereco], " & _
'						"[C].[Tipologradouro], " & _
'						"[C].[Logradouro], " & _
'						"[C].[Bairro], " & _
'						"[C].[Municipio] " & _
'						"FROM [marketingoki2].[dbo].[Solicitacao_coleta_has_Clientes] AS [A] " & _
'						"LEFT JOIN [marketingoki2].[dbo].[Pontos_coleta] AS [B] " & _
'						"ON [A].[Pontos_coleta_idPontos_coleta] = [B].[idPontos_coleta] " & _
'						"LEFT JOIN [marketingoki2].[dbo].[cep_consulta] AS [C] " & _
'						"ON [B].[cep_consulta_idcep_consulta] = [C].[idcep_consulta] " & _
'						"WHERE [A].[Clientes_idClientes] = " & Session("IDCliente") & " AND " & _
'						"[B].[status_pontocoleta] = 1"

		sSql = "select " & _
				"a.pontos_coleta_idpontos_coleta, " & _
				"b.nome_fantasia, " & _
				"b.numero_endereco, " & _
				"b.complemento_endereco, " & _
				"b.logradouro, " & _
				"b.bairro, " & _
				"b.municipio " & _
				"from solicitacao_coleta_has_clientes as a " & _
				"left join pontos_coleta as b " & _
				"on a.pontos_coleta_idpontos_coleta = b.idpontos_coleta " & _
				"where a.clientes_idclientes = "&session("IDCliente")&" and b.status_pontocoleta = 1"

		'idponto coleta = 0
		'nome fantasia  = 1
		'numero endereco= 2
		'comp endereco	= 3
		'logradouro		= 4
		'bairro			= 5
		'municipio		= 6

		Call search(sSql, arrListPontoColeta, intListPontoColeta)
		If intListPontoColeta > -1 Then
			For i=0 To intListPontoColeta
				IDPontoColeta				= arrListPontoColeta(0,i)
				NomeFantasiaPontoColeta		= arrListPontoColeta(1,i)
				LogradouroPontoColeta		= arrListPontoColeta(4,i) & ", n° " & arrListPontoColeta(2,i) & ", " & arrListPontoColeta(5,i) & ", " & arrListPontoColeta(6,i)
			Next
		End If
	End Sub

	Sub GetListDomiciliar()

		Dim sSql, arrListDomiciliar, intListDomiciliar, i

		sSql = "SELECT " & _
						"[A].[cep], " & _
						"[A].[logradouro], " & _
						"[A].[bairro], " & _
						"[A].[municipio], " & _
						"[A].[estado], " & _
						"[C].[compl_endereco_coleta], " & _
						"[C].[numero_endereco_coleta], " & _
						"[C].[contato_respcoleta], " & _
						"[C].[ddd_respcoleta], " & _
						"[C].[telefone_respcoleta] " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS [A] " & _
						"LEFT JOIN [marketingoki2].[dbo].[Clientes] AS [C] " & _
						"ON [A].[Clientes_idClientes] = [C].[idClientes] " & _
						"WHERE [A].[isEnderecoColeta] = 1 AND " & _
						"[A].[Clientes_idClientes] = " & Session("IDCliente")

		Call search(sSql, arrListDomiciliar, intListDomiciliar)
		If intListDomiciliar > -1 Then
			For i=0 To intListDomiciliar
				CEPColeta						= arrListDomiciliar(0,i)
				LogradouroColeta				= arrListDomiciliar(1,i)
				BairroColeta					= arrListDomiciliar(2,i)
				MunicipioColeta					= arrListDomiciliar(3,i)
				EstadoColeta					= arrListDomiciliar(4,i)
				CompLogradouro					= arrListDomiciliar(5,i)
				NumeroColeta					= arrListDomiciliar(6,i)
				ContatoRespColeta				= arrListDomiciliar(7,i)
				DDDContatoRespColeta			= arrListDomiciliar(8,i)
				TelefoneContatoRespColeta		= arrListDomiciliar(9,i)
			Next
		End If
	End Sub

	Function GetCepEnderecoComum()
		Dim sSql, arrCep, intCep, i
		Dim CEP

		sSql = "SELECT " & _
						"A.cep " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 1"

		Call search(sSql, arrCep, intCep)

		If intCep > -1 Then
			For i=0 To intCep
				CEP = arrCep(0,i)
			Next
		Else
			CEP = Null
		End If

		GetCepEnderecoComum	= CEP

	End Function

	Function GetNumeroEnderecoCliente()
		Dim sSql, arrNumero, intNumero, i
		Dim Numero

		sSql = "SELECT [numero_endereco] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")

		Call search(sSql, arrNumero, intNumero)

		If intNumero > -1 Then
			For i=0 To intNumero
				Numero = arrNumero(0,i)
			Next
		Else
			Numero = Null
		End If

		GetNumeroEnderecoCliente = Numero

	End Function

	Function GetCompLogradouroEnderecoCliente()
		Dim sSql, arrComp, intComp, i
		Dim CompLogradouro

		sSql = "SELECT [compl_endereco] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")

		Call search(sSql, arrComp, intComp)

		If intComp > -1 Then
			For i=0 To intComp
				GetCompLogradouroEnderecoCliente = arrComp(0,i)
			Next
		Else
			GetCompLogradouroEnderecoCliente = Null
		End If

	End Function

	Sub UpdateAdressColect()
		Dim oCommand

		Set oCommand = Server.CreateObject("ADODB.Command")
		oCommand.CommandTimeout = 200
		oCommand.ActiveConnection = oConn
		oCommand.CommandType = 4
		oCommand.CommandText = "sp_UpdateClienteColeta"

		' informações sobre o cliente e as atualizações que estavam fazendo normalmente
		oCommand.Parameters("@IDCliente")						= Session("IDCliente")
		oCommand.Parameters("@CompEndereco")					= CompLogradouro
		oCommand.Parameters("@NumeroEndereco")					= CLng(NumeroColeta)
		oCommand.Parameters("@ContatoRespColeta")				= ContatoRespColeta
		oCommand.Parameters("@DDDContatoRespColeta")			= CInt(DDDContatoRespColeta)
		oCommand.Parameters("@TelefoneContatoRespColeta")		= CLng(TelefoneContatoRespColeta)
		oCommand.Parameters("@RamalContatoRespColeta")			= RamalContatoRespColeta
		oCommand.Parameters("@DepContatoRespColeta")			= DepContatoRespColeta

		' atualização que foi feita no endereço
		oCommand.Parameters("@cep_coleta")						= CEPColeta
		oCommand.Parameters("@logradouro_coleta")				= LogradouroColeta
		oCommand.Parameters("@bairro_coleta")					= BairroColeta
		oCommand.Parameters("@municipio_coleta")				= MunicipioColeta
		oCommand.Parameters("@estado_coleta")					= EstadoColeta

		oCommand.Execute()

		Set oCommand = Nothing

'		Response.Redirect "frmOperacionalCliente.asp"

	End Sub

	Function MinCartuchos()
		Dim sSql, arrMin, intMin, i

		sSql = "SELECT [minCartuchos] " & _
			   "FROM [marketingoki2].[dbo].[Clientes] " & _
			   "WHERE [idClientes] = " & Session("IDCliente")
		Call search(sSql, arrMin, intMin)
		If intMin > -1 Then
			For i=0 To intMin
				minCartuchos = arrMin(0,i)
			Next
		End If
	End Function

	Function GetCategoria()
		Dim sSql, arrCategoria, intCategoria, i
		Dim minCartuchos

		sSql = "SELECT B.minCartuchos FROM Clientes AS A " & _
				"LEFT JOIN Categorias AS B " & _
				"ON B.idCategorias = A.Categorias_idCategorias " & _
				"WHERE A.idClientes = " & Session("IDCliente")

		Call search(sSql, arrCategoria, intCategoria)

		For i=0 To intCategoria
			minCartuchos = arrCategoria(0,i)
		Next

		GetCategoria = minCartuchos
	End Function

	If Session("isColetaDomiciliar") = 1 Then
		Call GetListDomiciliar()
	Else
		Call GetListPontoColeta()
	End If

	' ************************************************************************************************************
	' Peterson 3/5/2014
	' Função traz o CEP padrão da Coleta
	' ************************************************************************************************************
	Function GethiddenCEPCol()
		Dim sSql, arrCep, intCep, i
		Dim CEP

		sSql = "SELECT " & _
						"A.cep " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 0"

		Call search(sSql, arrCep, intCep)

		If intCep > -1 Then
			For i=0 To intCep
				CEP = arrCep(0,i)
			Next
		Else
			CEP = Null
		End If

		GethiddenCEPCol	= CEP

	End Function


	Function GethiddenLogrCol()
		Dim sSql, arrCep, intCep, i
		Dim LOGRA

		sSql = "SELECT " & _
						"A.logradouro " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 0"

		Call search(sSql, arrCep, intCep)

		If intCep > -1 Then
			For i=0 To intCep
				LOGRA = arrCep(0,i)
			Next
		Else
			LOGRA = Null
		End If
		GethiddenLogrCol	= LOGRA

	End Function

	Function GethiddenBaiCol()
		Dim sSql, arrSql, intSql, i
		Dim sBai

		sSql = "SELECT " & _
						"A.bairro " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 0"

		Call search(sSql, arrSql, intSql)

		If intSql > -1 Then
			For i=0 To intSql
				sBai = arrSql(0,i)
			Next
		Else
			sBai = Null
		End If
		GethiddenBaiCol	= sBai
	End Function

	Function GethiddenMunCol()
		Dim sSql, arrSql, intSql, i
		Dim sMun

		sSql = "SELECT " & _
						"A.municipio " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 0"

		Call search(sSql, arrSql, intSql)

		If intSql > -1 Then
			For i=0 To intSql
				sMun = arrSql(0,i)
			Next
		Else
			sMun = Null
		End If
		GethiddenMunCol	= sMun
	End Function

	Function GethiddenEstCol()
		Dim sSql, arrSql, intSql, i
		Dim sRet

		sSql = "SELECT " & _
						"A.estado " & _
						"FROM [marketingoki2].[dbo].[cep_consulta_has_Clientes] AS A " & _
						"WHERE A.[Clientes_idClientes] = " & Session("IDCliente") & _
						"AND A.[isEnderecoComum] = 0"

		Call search(sSql, arrSql, intSql)

		If intSql > -1 Then
			For i=0 To intSql
				sRet = arrSql(0,i)
			Next
		Else
			sRet = "EE"
		End If
		GethiddenEstCol	= sRet
	End Function

	Function GethiddenComplCol()
		Dim sSql, arrSql, intSql, i
		Dim sRet

		sSql = "SELECT [compl_endereco_coleta] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")

		Call search(sSql, arrSql, intSql)

		If intSql > -1 Then
			For i=0 To intSql
				sRet = arrSql(0,i)
			Next
		Else
			sRet = Null
		End If
		GethiddenComplCol	= sRet
	End Function

	Function GethiddenNumCol()
		Dim sSql, arrSql, intSql, i
		Dim sRet

		sSql = "SELECT [numero_endereco_coleta] FROM [marketingoki2].[dbo].[Clientes] WHERE [idClientes] = " & Session("IDCliente")

		Call search(sSql, arrSql, intSql)

		If intSql > -1 Then
			For i=0 To intSql
				sRet = arrSql(0,i)
			Next
		Else
			sRet = Null
		End If
		GethiddenNumCol	= sRet
	End Function

	Call SubmitForm()

'	Response.Write "Categoria: " & GetCategoria() & "<br />"
'	Response.Write "Min. Cartuchos: " & MinCartuchos() & "<br />"
%>

<html>
<head>

<script src="js/frmAddSolicitacao.js"></script>
<script src="js/frmAddSolCol.js"></script>

<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style type="text/css">
        #txtObservacao {
            width: 242px;
        }
    </style>
</head>

<%If Session("isColetaDomiciliar") = 0 Then%>
<body>
<%else%>
<!--<body onLoad="loadInfoSameAdress()">-->
<body>
<%end if%>
<body onLoad="loadClear()">
<div id="container">
	<!--#include file="inc/i_header.asp" -->
	<div id="conteudo">
		<table cellspacing="0" cellpadding="0" width="775">

			<form action="frmAddSolicitacao.asp" name="frmAddSolicitacao" method="POST">

			<input type="hidden" name="hiddenSessionisColetaDomiciliar" value="<%=Session("isColetaDomiciliar")%>" />
			<input type="hidden" name="hiddenIntPontoColeta" value="" />
			<input type="hidden" name="hiddenIntChangePontoColeta" value="<%=IDPontoColeta%>" />
			<input type="hidden" name="hiddenIntEnderecoCepColeta" value="" />
			<input type="hidden" name="hiddenGetCepEnderecoComum" value="<%=GetCepEnderecoComum()%>" />
			<input type="hidden" name="hiddenGetNumeroEnderecoCliente" value="<%=GetNumeroEnderecoCliente()%>" />
			<input type="hidden" name="hiddenGetCompLogradouroEnderecoCliente" value="<%=GetCompLogradouroEnderecoCliente()%>" />
			<input type="hidden" name="hiddenActionForm" value="0" />
			<input type="hidden" name="hiddenMinCartuchos" value="<%If MinCartuchos() = "" Or MinCartuchos() = 0 Then Response.Write GetCategoria() Else Response.Write MinCartuchos() End If %>" />

			<input type="hidden" name="hiddenCEPCol" value="<%=GethiddenCEPCol()%>" />
			<input type="hidden" name="hiddenLogrCol" value="<%=GethiddenLogrCol()%>" />
			<input type="hidden" name="hiddenComplCol" value="<%=GethiddenComplCol()%>" />
			<input type="hidden" name="hiddenNumCol" value="<%=GethiddenNumCol()%>" />
			<input type="hidden" name="hiddenBaiCol" value="<%=GethiddenBaiCol%>" />
			<input type="hidden" name="hiddenMunCol" value="<%=GethiddenMunCol()%>" />
			<input type="hidden" name="hiddenEstCol" value="<%=GethiddenEstCol()%>" />

			<tr>
				<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
				<td id="conteudo">
					<table cellpadding="3" cellspacing="4" width="100%" id="tableAddSolicitacao" border="0">
						<tr>
							<td colspan="4" id="explaintitle" align="center">Nova Solicitação de Coleta</td>
						</tr>
						<tr>
							<td colspan="4" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>
						</tr>
						<tr>
							<td colspan="4"><b id="fontred">Atenção :</b>
										<b style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-size: 13px; vertical-align: baseline; background-color: transparent; color: rgb(55, 61, 69); font-family: Arial, sans-serif; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: 14px; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Os campos com (asterisco)* são de preenchimento obrigatório.</b></td>
						</tr>
						<tr>
							<!--<td colspan="3" align="right"><a class="linkOperacional" href="javascript:window.location.href='frmOperacionalCliente.asp';">&laquo Voltar</a></td>-->
							<td colspan="1">&nbsp;</td>
							<td align="right" colspan="3">&nbsp;</td>

						</tr>
						<tr>
							<td align="right" width="40%">Quantidade de Cartucho(s) a coletar*</td>
							<td align="left" colspan="3"><input type="text" class="text" name="txtQtdCartuchos" value="" size="5" /></td>
						</tr>
						<tr>
							<td colspan="4">&nbsp;</td>
							<!--<td colspan="1" align="right"><input type="button" class="btnform" name="btnSubmitSolicitacao" value="Enviar Solicitação" onClick="validaFormulario()" /></td>
							<td colspan="1" align="right"><input type="button" class="btnform" name="btnSubmitSolicitacao" value="Enviar Solicitação" onClick="validaFormulario()" /></td>-->
						</tr>
							<tr>
								<td colspan="4" id="explaintitle" align="center" class="oki-h1">Selecione abaixo o endereço onde quer que seja feita a coleta ou
                                    <br />
                                    para um novo endereço selecione: Outro Endereço.</td>
							</tr>
							<tr>
								<td align="center" colspan="4">
									<input type="radio" class="radio" id="chkMesmoEndereco" name="radiogroup" value="true" onClick="preencheMesmoEndereco()" />Mesmo endereço da Empresa&nbsp;
									<input type="radio" class="radio" id="chkMesmoEndereco" name="radiogroup" value="true" onClick="preencheEndColeta()" /> Endereço padrão de coleta (previamente cadastrado)
									<input type="radio" class="radio" id="chkNovoEndereco"  name="radiogroup" value="true" onClick="preencheNovoEndereco()" />Outro Endereço</td>
								<!--<td align="left">Usar mesmo endereço de cadastro da Empresa</td>-->
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td>&nbsp;</td>
                                <td>&nbsp;</td>
							</tr>

							<tr>
								<td align="right" width="25%">CEP de Coleta*</td>
								<td align="left" with="5%">
									<input type="text" class="textreadonly" name="txtCepColeta" id="txtCepColeta" value="<%=CEPColeta%>" size="14" maxlength="8" readonly="true"/></td>
								<td align="left" width="10%">
									<img align="absmiddle" style="cursor:pointer;" src="img/buscar.gif" name="btnBuscarCepColeta" id="btnBuscarCepColeta" alt="Buscar CEP" onClick="loadCepColeta()" /></td>
								<td align="left">
									Ex: 9999999 (apenas 8 números)</td>
							</tr>
							<tr>
								<td align="right" width="25%">Endereço*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtLogradouroColeta" value="<%=LogradouroColeta%>" size="40" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Complemento do Endereço</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" name="txtCompLogradouroColeta" value="<%=CompLogradouro%>" size="40" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Número*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" id="txtNumeroColeta" name="txtNumeroColeta" value="<%=NumeroColeta%>" size="10" maxlength="8" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Bairro*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" id="txtBairroColeta" name="txtBairroColeta" value="<%=BairroColeta%>" size="40" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Município*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" id="txtMunicipioColeta" name="txtMunicipioColeta" value="<%=MunicipioColeta%>" size="40" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Estado*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" id="txtEstadoColeta" name="txtEstadoColeta" value="<%=EstadoColeta%>" size="03" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">Contato para Retirada da Coleta*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" id="txtRespColContato" name="txtRespColContato" value="<%=ContatoRespColeta%>" size="40" readonly="true"/></td>
							</tr>
							<tr>
								<td align="right" width="25%">DDD do Contato*</td>
								<td align="left" width="10%"><input type="text" class="textreadonly" id="txtDDDContatoRespColeta" name="txtDDDContatoRespColeta" value="<%=DDDContatoRespColeta%>" size="3" maxlength="2" readonly="true"/></td>
								<td align="left" colspan="2">Ex: 11</td>
							</tr>
							<tr>
								<td align="right" width="25%">Telefone do Contato*</td>
								<td align="left"><input type="text" class="textreadonly" id="txtTelefoneContatoRespColeta" name="txtTelefoneContatoRespColeta" value="<%=TelefoneContatoRespColeta%>" size="15" maxlength="9" readonly="true"/></td>
								<td align="left" colspan="2">Ex: 999999999</td>
							</tr>
							<tr>
								<td align="right" width="25%">Ramal</td>
								<td align="left"><input type="text" class="textreadonly" id="txtRamalContatoRespColeta" name="txtRamalContatoRespColeta" value="<%=RamalContatoRespColeta%>" size="11" maxlength="4" readonly="true"/></td>
								<td align="left" colspan="2">Ex: 9999</td>
							</tr>

							<tr>
								<td align="right" width="25%">Departamento*</td>
								<td align="left" colspan="3"><input type="text" class="textreadonly" style="text-transform: uppercase;" id="txtDepContatoRespColeta" name="txtDepContatoRespColeta" value="<%=DepContatoRespColeta%>" size="30" maxlength="30" readonly="true" /></td>
							</tr>

							<tr>
								<td align="right" width="25%">&nbsp;</td>
								<td align="left" colspan="3">&nbsp;</td>
							</tr>

							<tr>
								<td align="right" width="25%">Observação quanto ao horário de entrega:</td>
								<td align="left" colspan="3">&nbsp;</td>
							</tr>

							<tr>
								<td align="right" width="25%">&nbsp;</td>
								<td align="left" colspan="3"><textarea name="txtObservacao"><%=txtObservacao%></textarea></td>
                                    <!--<input type="text" class="textreadonly" id="txtObservacao" name="txtObservacao" value="<%=txtObservacao%>" size="60" maxlength="80" /></td>-->
							</tr>

							<tr>
								<td align="right" width="25%">&nbsp;</td>
								<td align="right" colspan="3"><input type="button" class="btnform" name="btnSubmitSolicitacao" value="Salvar" onClick="validaFormulario()" /></td>
							</tr>

							<tr>
								<td colspan="1">&nbsp;</td>
								<!--<td colspan="2" align="center"><input type="button" class="btnform" name="btnChangeAdressColect" value="Alterar Endereço" onClick="authenticateUpdateAdress()" /></td>-->
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
