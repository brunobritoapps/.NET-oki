<!--#include file="_config/_config.asp" -->
<%
'|--------------------------------------------------------------------
'| Arquivo: frmCadCliente.asp
'| Autor: Leandro dos Santos (leandro.storoli@gmail.com)
'| Data Criação: 13/04/2007
'| Data Modificação : 15/04/2007
'| Descrição: Arquivo de Formulário para cadastro de Cliente (ASP)
'|--------------------------------------------------------------------
%>
<%Call open()%>
<%

    '
    '============================================================================================
    ' Coleta
	Dim IdentifiedCharacterColeta
	Dim DateMonthColeta
	Dim DateYearColeta
	Dim NumeroSolicitacaoColeta

    ' Hiddens
	Dim hiddenTipoColeta

	'Definição de Tipo de Coleta
	Dim Categoria
	Dim QuantidadeCartuchos

    Dim Grupos_idGrupos
	Dim Categorias_idCategorias
    Dim razao_social
	Dim nome_fantasia
	Dim cnpj
    Dim inscricao_estadual
	Dim ddd
	Dim telefone
	Dim compl_endereco
    Dim compl_endereco_coleta
    Dim numero_endereco
    Dim numero_endereco_coleta
    Dim contato_respcoleta
    Dim ddd_respcoleta
    Dim telefone_respcoleta
    Dim numero_sequencial
    Dim data_atualizacao_sequencial
    Dim minCartuchos
    Dim typeColect
    Dim status_cliente
    Dim motivo_status
    Dim bonus_type
    Dim Transportadoras_idTransportadoras
    Dim tipopessoa
    Dim cod_cli_consolidador
    Dim cod_bonus_cli
    Dim ramal_respcoleta
    Dim depto_respcoleta

    Dim hiddenTudoOk

    'dados do usuário master
    Dim Nome
    Dim Usuario
    Dim Senha
	Dim Email
    Dim Cpf

    Dim EmailAdm

    'dados para o cadastro do endereço principal e coleta default
    Dim cep_cliente
    Dim logradouro_cliente
    Dim bairro_cliente
    Dim municipio_cliente
    Dim estado_cliente

    '
    'Pega email Adm
    Function GetEmailAdm()

        Dim sSql, arrS, intS

		sSql = "SELECT [email] " & _
			   "FROM [marketingoki2].[dbo].[Administrator] " & _
			   "WHERE [aprovador] = 1 and [status] = 1 "

		Call search(sSql, arrS, intS)

        If intS > -1 Then
            EmailAdm = arrS(0,0)
        End if

        GetEmailAdm = EmailAdm

    End Function

    '
	'============================================================================================
	'| Sub que gera as Categorias para o cliente Selecionar
	'============================================================================================
	Sub geraCategorias()
		Dim sSql, arrCat, intCat, i
		sSql = "SELECT [idCategorias]" & _
			   ",[descricao]" & _
			   ",[ativo]" & _
			   ",[isColetaDomiciliar]" & _
			   ",[minCartuchos]" & _
			   "FROM [marketingoki2].[dbo].[Categorias] " & _
			   "WHERE [ativo] = 1"
		Call search(sSql, arrCat, intCat)

		With Response
			If intCat > -1 Then
				For i=0 To intCat
					.Write "<option value='"&arrCat(0,i)&"'>"&arrCat(1,i)&"</option>"
				Next
			Else
				.Write "<option value='-1'>Nenhuma Categoria cadastrada.</option>"
			End If
		End With
	End Sub

    '
    'envia email adm
    Sub EnviaEMailAdm()

        EmailAdm = GetEmailAdm()

        'objCDOSYSMail.CC    = GetEmailAdm()
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
							    "<p>Prezado(a) Administrador " & "" & "<br /><br /> " & _
							    "O cliente: " & Nome & " efetuou um cadastro em nosso sistema.<br /><br /> " & _
                                "Favor prosseguir com a análise para Aprovação ou Rejeição.<br /><br /> " & _
							    "<b style=""color:#990000"">Workflow OKI Printing Solutions</b> " & _
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
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Oki!321!" '"!nfe321!"        'senha
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
		objCDOSYSMail.To    = EmailAdm
        objCDOSYSMail.CC    = "peterson.aquino@hotmail.com"

		'ASSUNTO DA MENSAGEM
		objCDOSYSMail.Subject = "Workflow Okidata - Novo Cliente no Portal Sustentabilidade"

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



    '
    '
    '
    'envia email cliente
    Sub EnviaEmailCadastro()

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
							    "<p>Prezado(a) " & Nome  & "<br /> " & _
							    "Agradecemos primeiramente pelo cadastro.<br /><br /><br /> " & _
                                "Seu cadastro está em análise e assim que tivermos a liberação, enviaremos para você um e-mail com os dados para acesso.<br /><br /> " & _
                                "Guarde seu <b>Usuário</b>: " & Usuario & " e <b>Senha</b>: " & Senha & " <br /><br /> " & _
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
        objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.sustentabilidadeoki.com.br" '"mail.okidata.com.br"
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = "sustentabilidadeoki@sustentabilidadeoki.com.br" '"nfe@okidata.com.br" 'Email
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Oki!321!" '"!nfe321!"        'senha
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
		objCDOSYSMail.To    = Email
        'objCDOSYSMail.CC    = "peterson.aquino@hotmail.com"

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

    '
    '============================================================================================
	'| Sub que Controla o Submit do Form de Cadastro do Cliente
	'============================================================================================
	Sub SubmitForm()
		If Request.ServerVariables("HTTP_METHOD") = "POST" Then
            'if request.Form("hiddenTudoOk") = "true" Then
                Call RequestForm()
                Cpf                     = Request.Form("txtCPFnum")
                Call CadClienteLc()

                'response.Write "<script>alert('Cadastro efetuado com Sucesso! Por gentileza aguardar a aprovação do mesmo. Você receberá um e-mail de confirmação do cadastro e logo entraremos em contato!')</script>"
                Call EnviaEmailCadastro()
                Call EnviaEMailAdm()
                Response.redirect "index.asp?area=home&cadastro=ok"
            'End if
		End If
	End Sub

    '
    'inclusão: peterson 11-5-2014
    '
    Sub CadClienteLc()

		Dim objCommand
		Dim lTipoPessoa

        If Request.ServerVariables("HTTP_METHOD") = "POST" Then

    		On Error Resume Next

        	Set objCommand = Server.CreateObject("ADODB.Command")
	    	Set rs = Server.CreateObject("ADODB.Recordset")

		    objCommand.CommandTimeout = 200
		    objCommand.ActiveConnection = oConn
		    objCommand.CommandText = "sp_AddClientLc"
		    objCommand.CommandType = 4

        	objCommand.Parameters("@Grupos_idGrupos")           = 1
	        objCommand.Parameters("@Categorias_idCategorias")   = Categorias_idCategorias
	        objCommand.Parameters("@numero_sequencial")         = numero_sequencial
            objCommand.Parameters("@razao_social")              = razao_social
            objCommand.Parameters("@nome_fantasia")             = nome_fantasia
            objCommand.Parameters("@cnpj") = cnpj

            objCommand.Parameters("@inscricao_estadual") = inscricao_estadual

            objCommand.Parameters("@ddd") = ddd
            objCommand.Parameters("@telefone") = telefone
            objCommand.Parameters("@compl_endereco") = compl_endereco
            objCommand.Parameters("@compl_endereco_coleta") = compl_endereco_coleta
            objCommand.Parameters("@numero_endereco") = numero_endereco
            objCommand.Parameters("@numero_endereco_coleta") = numero_endereco_coleta
            objCommand.Parameters("@contato_respcoleta") = contato_respcoleta
            objCommand.Parameters("@ddd_respcoleta") = ddd
            objCommand.Parameters("@telefone_respcoleta") = telefone

            'objCommand.Parameters("@data_atualizacao_sequencial") = data_atualizacao_sequencial
            'objCommand.Parameters("@minCartuchos") = minCartuchos
            'objCommand.Parameters("@typeColect") = typeColect
            objCommand.Parameters("@status_cliente") = "0"
            objCommand.Parameters("@motivo_status") = ""
            'objCommand.Parameters("@bonus_type") = bonus_type
            'objCommand.Parameters("@Transportadoras_idTransportadoras") = Transportadoras_idTransportadoras

            objCommand.Parameters("@tipopessoa") = tipopessoa

            objCommand.Parameters("@cod_cli_consolidador") = "0"
            objCommand.Parameters("@cod_bonus_cli") = ""
            objCommand.Parameters("@ramal_respcoleta") = ramal_respcoleta
            objCommand.Parameters("@depto_respcoleta") = depto_respcoleta

            objCommand.Parameters("@Nome") = Nome
            objCommand.Parameters("@Usuario") = Usuario
            objCommand.Parameters("@Senha") = Senha
	        objCommand.Parameters("@Email") = Email

            objCommand.Parameters("@cep_cliente") = cep_cliente
			objCommand.Parameters("@logradouro_cliente") = logradouro_cliente
			objCommand.Parameters("@bairro_cliente") = bairro_cliente
			objCommand.Parameters("@municipio_cliente") = municipio_cliente
			objCommand.Parameters("@estado_cliente") = estado_cliente

    	    rs.Open objCommand

		    Set rs = nothing
		    Set objCommand = Nothing

        End If
    End Sub

	'============================================================================================
	'| Sub que pega todos os valores possíveis e necessários do Form via Request
	'============================================================================================
	Sub RequestForm()

		If Request.ServerVariables("HTTP_METHOD") = "POST" Then

	        Grupos_idGrupos         = "1"
            Cpf                     = Request.Form("txtCPFnum")
	        Categorias_idCategorias = Request.Form("cbCategorias")

			If CInt(Request.Form("radioPessoa")) = 0 Then
    			razao_social		= Request.Form("txtNome")
				nome_fantasia	    = Request.Form("txtNome")
				cnpj			    = Request.Form("txtCPFnum")
                Cpf                     = Request.Form("txtCPFnum")
                inscricao_estadual	= ""
                tipopessoa          = "0"
			Else
				razao_social		= Request.Form("txtRazaoSocial")
				nome_fantasia	    = Request.Form("txtFanta")
				cnpj			    = Request.Form("txtNCNPJ")
                inscricao_estadual	= Request.Form("txtIE")
                tipopessoa          = "1"
			End If
	        ddd                     = Request.Form("txtDDD")
	        telefone                = Request.Form("txtTelefone")
	        compl_endereco          = Request.Form("txtCompLogradouro")
	        compl_endereco_coleta   = Request.Form("txtCompLogradouro")
	        numero_endereco         = Request.Form("txtNumero")
	        numero_endereco_coleta  = Request.Form("txtNumero")
	        contato_respcoleta      = Request.Form("txtContatoColeta")
	        ddd_respcoleta          = Request.Form("txtDDD")
	        telefone_respcoleta     = Request.Form("txtTelefone")
	        numero_sequencial       = getSequencial(True)
	        data_atualizacao_sequencial = "NULL"
	        minCartuchos            = "0"
	        typeColect              = "1"
	        status_cliente          = "NULL"
	        motivo_status           = "NULL"
	        bonus_type              = "NULL"
	        Transportadoras_idTransportadoras = "0"
	        'tipopessoa              = CInt(TipoPessoa)
	        cod_cli_consolidador    = "NULL"
	        cod_bonus_cli           = "NULL"
	        ramal_respcoleta        = Request.Form("txtRamal")
	        depto_respcoleta        = Request.Form("txtDepartamento")

            'dados para cadastro do usuário master
            Nome = Request.Form("txtContatoColeta")
			Usuario				    = Request.Form("txtUsuario")
			Senha				    = Request.Form("txtSenha")
			Email				    = Request.Form("txtEmail")

            'dados para o cadastro do endereço principal e coleta default
    		cep_cliente			    = request.Form("txtCep")
			logradouro_cliente	    = request.Form("txtLogradouro")
			bairro_cliente		    = request.Form("txtBairro")
			municipio_cliente	    = request.Form("txtMunicipio")
			estado_cliente		    = request.Form("txtEstado")

            'Response.Write "<script>alert('" & RazaoSocial & "-" & NomeFantasia & CNPJ & InscricaoEstadual & "Contato:" & Contato & "Depto:" & Departamento & "DDD" & DDD & "')</script>"

        End If

	End Sub

	'============================================================================================
	'| Sub que Cadastra concretamente o Cliente de acordo com os parametros necessarios da PROC
	'| sp_AddClient
	'============================================================================================
	Sub CadCliente()
		On Error Resume Next

        'Response.Write "sp_AddClient @IDGrupos = 1 , @IDCategorias = "&Categoria&", "&_
        '"@RazaoSocial = '" & NomeFantasia & "', @NomeFantasia = '" & NomeFantasia & "', @CNPJ = '" & CNPJ & "',  "&_
        '"@InscricaoEstadual = '" & InscricaoEstadual & "', @DDD = '" & DDD & "', @Telefone = " & Telefone & ", @ComplementoEndereco = '" & CompLogradouro & "',  "&_
        '"@ComplementoEnderecoColeta = '" & CompLogradouroColeta & "', @NumeroEndereco = " & Numero & ", @NumeroEnderecoColeta = " & NumeroColeta & ", "&_
        '"@NumeroSequencial = '" & Mid(NumeroSolicitacaoColeta, 6, 5) & "', @BonusType = NULL, @CepEndereco = '" & Cep & "',  "&_
        '"@CepEnderecoColeta = '" & CepColeta & "', @IDPontoColeta = " & IDPontoDeColeta & ", @NumeroSolicitacaoColeta = '" & NumeroSolicitacaoColeta & "',  "&_
        '"@isColetaDomiciliar = '" & hiddenTipoColeta & "', @QtdCartuchos = '" & QuantidadeCartuchos & "', @Contato = '" & Contato & "',  "&_
        '"@Usuario = '" & Usuario & "', @Senha = '" & Senha & "', @Email = '" & Email & "', @ContatoResp = '" & ContatoRespColeta & "',  "&_
        '"@DDDContatoResp = '" & DDDContatoRespColeta & "', @TelefoneContatoResp = '" & TelefoneContatoRespColeta & "', @TIPOPESSOA = " & CInt(TipoPessoa) & ",  "&_
        '"@logradouro_cliente = '"&logradouro_cliente&"', @logradouro_cliente_coleta = '"&logradouro_coleta&"',  "&_
        '"@bairro_cliente = '"&bairro_cliente&"', @bairro_cliente_coleta = '"&bairro_coleta&"',  "&_
        '"@municipio_cliente = '"&municipio_cliente&"', @municipio_cliente_coleta = '"&municipio_coleta&"',  "&_
        '"@estado_cliente = '"&estado_cliente&"', @estado_cliente_coleta='"&estado_coleta&"' "

        'Response.End

		Dim objCommand
		Dim lTipoPessoa
		'dim sMsg

		Set objCommand = Server.CreateObject("ADODB.Command")

		objCommand.CommandTimeout = 200
		objCommand.ActiveConnection = oConn
		objCommand.CommandText = "sp_AddClientLc"
		objCommand.CommandType = 4

		objCommand.Parameters("@IDGrupos")						= 1
		objCommand.Parameters("@IDCategorias")					= Categoria
		If RazaoSocial <> "" then
		objCommand.Parameters("@RazaoSocial")					= RazaoSocial
		Else
		objCommand.Parameters("@RazaoSocial")					= NomeFantasia
		end if
		objCommand.Parameters("@NomeFantasia")					= NomeFantasia
		objCommand.Parameters("@CNPJ")							= CNPJ
		objCommand.Parameters("@InscricaoEstadual")				= InscricaoEstadual
		objCommand.Parameters("@DDD")						    = DDD
		objCommand.Parameters("@Telefone")						= Telefone
		objCommand.Parameters("@ComplementoEndereco")			= CompLogradouro
		objCommand.Parameters("@ComplementoEnderecoColeta")		= CompLogradouroColeta
		objCommand.Parameters("@NumeroEndereco")				= Numero
		objCommand.Parameters("@NumeroEnderecoColeta")			= NumeroColeta
		objCommand.Parameters("@NumeroSequencial")				= Mid(NumeroSolicitacaoColeta, 6, 5)
		objCommand.Parameters("@BonusType")						= "NULL"
		objCommand.Parameters("@CepEndereco")					= Cep
		objCommand.Parameters("@CepEnderecoColeta")				= CepColeta
		objCommand.Parameters("@IDPontoColeta")					= IDPontoDeColeta
		objCommand.Parameters("@NumeroSolicitacaoColeta")		= NumeroSolicitacaoColeta
		objCommand.Parameters("@isColetaDomiciliar")		    = hiddenTipoColeta
		objCommand.Parameters("@QtdCartuchos")					= QuantidadeCartuchos
		objCommand.Parameters("@Contato")						= Contato
		objCommand.Parameters("@Usuario")						= Usuario
		objCommand.Parameters("@Senha")							= Senha
		objCommand.Parameters("@Email")							= Email
		objCommand.Parameters("@ContatoResp")					= ContatoRespColeta
		objCommand.Parameters("@DDDContatoResp")				= DDDContatoRespColeta
		objCommand.Parameters("@TelefoneContatoResp")			= TelefoneContatoRespColeta
		objCommand.Parameters("@TIPOPESSOA")					= CInt(TipoPessoa)
		objCommand.Parameters("@logradouro_cliente")			= logradouro_cliente
		objCommand.Parameters("@logradouro_cliente_coleta")		= logradouro_coleta
		objCommand.Parameters("@bairro_cliente")				= bairro_cliente
		objCommand.Parameters("@bairro_cliente_coleta")			= bairro_coleta
		objCommand.Parameters("@municipio_cliente")				= municipio_cliente
		objCommand.Parameters("@municipio_cliente_coleta")		= municipio_coleta
		objCommand.Parameters("@estado_cliente")				= estado_cliente
		objCommand.Parameters("@estado_cliente_coleta")			= estado_coleta
    	objCommand.Parameters("@Departamento")      			= Departamento

		Dim rs, lIdSolicitacaoColeta

		Set rs = Server.CreateObject("ADODB.Recordset")

		rs.CursorType=1
		rs.Open objCommand


		'lIdSolicitacaoColeta = rs.Fields(1) 'comentado peterson: 10-5-2014 não terá mais retorno desta informação;
		''sMsg = rs.Fields(0)

		ser rs = nothing

		'objCommand.Execute()
		'If hiddenTipoColeta = 1 Then
		'	Response.Write "<script>alert('Seu cadastro foi submetido a aprovação, em breve entraremos em contato para " & _
		'								 "providenciar a coleta!');</script>"
		'Else
		'	Response.Write "<script>alert('Seu cadastro foi submetido a aprovação, em breve entraremos em contato para " & _
		'								 "autorizar a entrega do(s) cartucho(s) no ponto de coleta!')</script>"
		'End If
		Set objCommand = Nothing

		'response.redirect "index.asp?area=home"
		if len(trim(InscricaoEstadual)) = 0 then
			lTipoPessoa = 0 'PF
		else
			lTipoPessoa = 1 'PJ
		end if

		'If sMsg = "Cliente já cadastrado!" Then
		'	Response.Write "<script>alert('"&sMsg&"')</script>"
		'	Response.End()
		'End If

        'comentado: peterson 10-5-2014
        ''não terá mais esta informação, pois não fará nenhuma coleta até então;
		''response.redirect "frmCartaDoacaoNF.asp?IdSolicitacaoColeta=" & lIdSolicitacaoColeta & "&TipoPessoa=" & lTipoPessoa & "&TipoColeta=" & hiddenTipoColeta
		'Response.Write "frmCartaDoacaoNF.asp?IdSolicitacaoColeta=" & lIdSolicitacaoColeta & "&TipoPessoa=" & lTipoPessoa & "&TipoColeta=" & hiddenTipoColeta

		''If Err.number <> 0 Then
		''	Response.Write Err.Description
		''	Response.End()
		''End If

	End Sub
	'============================================================================================
	'| Chama a Operação de Cadastro
	'============================================================================================
	Call SubmitForm()

%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="css/geral.css">
<title><%=TITLE%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <style type="text/css">
        .auto-style5 {
            height: 28px;
        }
        .rotulos {
            font-family:Verdana;
            font-size:10px;
        }
        </style>
</head>
	<script language="javascript" type="text/javascript" src="js/frmCadCliente.js"></script>
		<div id="container">
			<!--#include file="inc/i_header.asp" -->
			<div id="conteudo">
				<table cellspacing="0" cellpadding="0" width="775">
					<tr>
						<td width="11" background="img/Bg_LatEsq.gif">&nbsp;</td>
						<td id="conteudo">

							<form action="frmCadCliente.asp" name="frmCadCliente" method="POST" onsubmit="return validaCadClienteContato()">
                            <!--<form name="frmCadCliente">-->
								<input type="hidden" name="hiddenMinCartuchos" value="" /> <input type="hidden" name="hiddenTypeColeta" value="" />
								<input type="hidden" name="hiddenIntEnderecoCep" value="" /> <input type="hidden" name="hiddenIntPontoColeta" value="" />
								<input type="hidden" name="hiddenIntEnderecoCepColeta" value="" /> <input type="hidden" name="hiddenIntChangePontoColeta" value="" />
								<input type="hidden" name="hiddenControleColeta" value="" />
								<input type="hidden" name="hiddenControleCartaNF" value="" />
                                <input type="hidden" name="hiddenTudoOk" value="false" />

								<table cellpadding="3" cellspacing="4" width="100%" id="tableTitle">
									<tr>
										<td colspan="2">
                                            <h3 style="margin: 0px; padding: 0px; border: 0px; font-size: 21px; vertical-align: baseline; background-color: transparent; font-family: Arial, sans-serif; line-height: 24px; color: rgb(244, 121, 32); font-style: normal; font-variant: normal; letter-spacing: normal; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Preencha os campos abaixo e cadastre-se</h3>
                                        </td>
									</tr>
									<tr>
										<td colspan="2"><b id="fontred">Atenção :</b>
										<b style="margin: 0px; padding: 0px; border: 0px; outline: 0px; font-size: 13px; vertical-align: baseline; background-color: transparent; color: rgb(55, 61, 69); font-family: Arial, sans-serif; font-style: normal; font-variant: normal; letter-spacing: normal; line-height: 14px; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Os campos com (asterisco)* são de preenchimento obrigatório.</b><td>
									</tr>
								</table>

								<table cellpadding="3" cellspacing="4" width="100%" id="tableCadClienteCategoria">
                                    <tr>
                                        <td width="25%">&nbsp;</td>
                                        <td width="75%">&nbsp;</td>
                                    </tr>
									<tr>
										<td colspan="2">
                                            <h4 style="margin: 0px; padding: 0px; border: 0px; font-size: 15px; vertical-align: baseline; background-color: transparent; font-family: Arial, sans-serif; line-height: 24px; color: rgb(244, 121, 32); font-style: normal; font-variant: normal; letter-spacing: normal; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Selecione uma categoria</h4>
										</td>
									</tr>

									<tr>
										<td align="right" class="auto-style4">Categoria:</td>
										<td align="left">
											<select name="cbCategorias" class="select" onchange="updateMinimo()">
												<option value="-1">Selecione uma Categoria</option>
												<%Call geraCategorias()%>
											</select>
											*
										</td>
									</tr>
                                </table>
                                <table style="width:100%;" id="Table1">
                                    <tr>
                                        <td style="width:50px;">&nbsp;</td>
                                        <td style="text-align:right;"><input type="radio" name="radioPessoa" value="0" onclick="checkPessoa()" /></td><td style="text-align:left; font-family:Verdana; font-size:10px; width:208px;">pessoa física</td>
                                        <td>&nbsp;</td>
                                        <td style="text-align:right;"><input type="radio" name="radioPessoa" checked="checked" value="1" onclick="checkPessoa()" /></td><td style="text-align:left; font-family:Verdana; font-size:10px; width:208px;">pessoa jurídica</td>
                                        <td>&nbsp;</td>
                                    </tr>
                                </table>

                                <table style="width:100%;" id="pessoajuridica">

                                    <tr id="razaosocial" style="display:block;">
                                        <td style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Razão social*</td>
                                        <td colspan="2"><input type="text"  style="text-transform: uppercase;" class="textreadonly" name="txtRazaoSocial" value="" size="40" /></td>
                                    </tr>

									<tr id="nomefantasia" style="display:block;">
										<td style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Nome Fantasia*</td>
										<td colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtFanta" value="" size="40" /></td>
									</tr>

                                    <tr id="cnpj" style="display:block;">
                                        <td style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">CNPJ*</td>
                                        <td align="left">
                                            <input type="text" class="textreadonly" name="txtNCNPJ" value="" size="22" maxlength="18" onkeypress="cnpj_format(this)" onBlur="checkCNPJEmpresa()" />
                                            <!--253.144.434-11-->
										</td>
                                        <td align="left" class="rotulos">
                                            Preencher com números Ex: 9999999999999</td>
                                    </tr>

                                    <tr id="inscestadual" style="display:block;">
                                        <td style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Inscrição Estadual*</td>
                                        <td align="left">
                                            <input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtIE" value="" size="18" maxlength="15" onkeypress="keypressIE(this.value)" onBlur="onblurie(this.value)" />
                                        </td>
                                        <td align="left" class="rotulos">
                                            Preencha Somente com Números ou com a palavra: ISENTO</td>
									</tr>
                                </table>

                                <table style="width: 100%;" id="pessoafisica">
                                    <tr id="nome" style="display: none;">
                                        <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 208px;">Nome*</td>
                                        <td colspan="2">
                                            <input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtNome" value="" size="40" /></td>
                                    </tr>

                                    <tr id="cpf" style="display: none;">
                                        <td style="text-align: right; font-family: Verdana; font-size: 10px; width: 208px;">CPF*</td>
                                        <td align="left">
                                            <input type="text" class="textreadonly" name="txtCPFnum" value="" size="22" maxlength="14" onkeypress="cpf_format(this)" onblur="checkCPF(this.value)" /></td>
                                        <td align="left" class="rotulos">
                                            Preencher com números Ex: 99999999999</td>
                                    </tr>
                                </table>

								<table style="width:100%;" id="endereco">

									<tr><td colspan="3">
										&nbsp;</td>
									</tr>
									<tr><td colspan="3">
										<h4 style="margin: 0px; padding: 0px; border: 0px; font-size: 15px; vertical-align: baseline; background-color: transparent; font-family: Arial, sans-serif; line-height: 24px; color: rgb(244, 121, 32); font-style: normal; font-variant: normal; letter-spacing: normal; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Endereço</h4>
                                        </td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">CEP*</td>
										<td align="left">
											<input type="text" class="textreadonly" name="txtCep" value="" size="10" maxlength="8"/>
											<INPUT type="button" class="BotaoBuscar" value="" id=btnBuscarCepColeta name=btnBuscarCepColeta alt="Buscar CEP" onClick="loadCepComum()"></td>
										<td align="left" class="rotulos">
											Preencha o CEP somente com 8 números Ex: 99999999</td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Endereço*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtLogradouro" value="" size="40" maxlength="50"/></td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Complemento</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtCompLogradouro" value="" size="40" maxlength="500"/></td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Número*</td>
										<td align="left" colspan="2"><input type="text" class="textreadonly" name="txtNumero" value="" size="10" maxlength="8" /></td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Bairro*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtBairro" value="" size="40" maxlength="50"/></td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Município*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtMunicipio" value="" size="40" maxlength="50" /></td>
									</tr>
									<tr>
										<td align="right" style="text-align:right; font-family:Verdana; font-size:10px; width:208px;">Estado*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtEstado" value="" size="2" maxlength="2" style="width:30px;"/></td>
									</tr>
									<tr>
										<td align="left" class="auto-style4">&nbsp;</td>
										<td align="right" colspan="2">&nbsp;</td>
									</tr>
								</table>

								<table cellpadding="3" cellspacing="4" width="100%" id="tableCadClienteContato" style="display:block;">
									<tr><td colspan="3">
										<h4 style="margin: 0px; padding: 0px; border: 0px; font-size: 15px; vertical-align: baseline; background-color: transparent; font-family: Arial, sans-serif; line-height: 24px; color: rgb(244, 121, 32); font-style: normal; font-variant: normal; letter-spacing: normal; orphans: auto; text-align: left; text-indent: 0px; text-transform: none; white-space: normal; widows: auto; word-spacing: 0px; -webkit-text-stroke-width: 0px; background-position: initial initial; background-repeat: initial initial;">Informação do principal contato da empresa</h4>
                                        </td>
									</tr>
									<tr>
										<td align="right" width="25%" class="auto-style5">Nome do Contato*</td>
										<td align="left" class="auto-style5" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtContatoColeta" value="" size="40" maxlength="50" /></td>
									</tr>
									<tr>
										<td align="right" width="25%">Login*</td>
										<td align="left"><input type="text" class="textreadonly" name="txtUsuario" value="" size="20" maxlength="20" onblur="checkUsuario()" /></td>
										<td align="left">Tamanho min 6, max 20 caracteres</td>
									</tr>
									<tr>
										<td align="right" width="25%">Senha*</td>
										<td align="left"><input type="password" class="textreadonly" name="txtSenha" value="" size="20" maxlength="20" /></td>
										<td align="left">Tamanho min 6, max 20 caracteres</td>
									</tr>
									<tr>
										<td align="right" width="25%">Confirmação da senha*</td>
										<td align="left" colspan="2"><input type="password" class="textreadonly" name="txtSenhaconfirma" value="" size="20" maxlength="20" /></td>
									</tr>
									<tr>
										<td align="right" width="25%">E-mail*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform:lowercase;" class="textreadonly" name="txtEmail" value="" size="40" maxlength="50"/></td>
									</tr>

									<tr>
										<td align="right" width="25%">Departamento*</td>
										<td align="left" colspan="2"><input type="text" style="text-transform: uppercase;" class="textreadonly" name="txtDepartamento" value="" size="40" maxlength="50" /></td>
									</tr>
									<tr>
										<td align="right" width="25%">DDD*</td>
										<td align="left"><input type="text" class="textreadonly" name="txtDDD" value="" size="04" maxlength="2" />
                                            </td>
										<td align="left">Desprezar o zero a esquerda Ex: 11</td>
									</tr>
									<tr>
										<td align="right" width="25%">Telefone*</td>
										<td align="left">
                                            <input type="text" class="textreadonly" name="txtTelefone" value="" size="15" maxlength="9" /></td>
										<td align="left">Preencher somente com números Ex: 999999999</td>
									</tr>
									<tr>
										<td align="right" width="25%">Ramal</td>
										<td align="left"><input type="text" class="textreadonly" name="txtRamal" value="" size="10" "/></td>
									</tr>
									<tr>
										<td align="left"></td>
										<td align="left">Preencher somente com números Ex: 999999999</td>
										<td align="right">&nbsp;</td>
									</tr>
									<tr>
										<td align="right" width="25%">&nbsp;</td>
										<td align="left" colspan="2">&nbsp;</td>
										<td align="right">&nbsp;</td>
									</tr>
									<tr>
										<td align="right" width="25%">&nbsp;</td>
										<td align="right" colspan="2">
                                            <!--<input type="button" class="btnformMaior" name="btnNextToSolicitacaoColeta" value="Concluir" alt="Concluir" onclick="validaCadClienteContato()"/></td>-->
                                            <input type="submit" class="btnformMaior" name="btnNextToSolicitacaoColeta" value="Concluir" alt="Concluir"/></td>
										<td align="right">&nbsp;</td>
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
<!--#include file="_config/colectobject.asp" -->
